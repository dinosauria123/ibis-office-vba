// SPDX-License-Identifier: AGPL-3.0-only
// Copyright (c) ibis inc.
// Additional terms under AGPL-3.0 section 7(b): see NOTICE.
//! VBA plugin for ibisOffice: macros in .xlsm/.xlsb/.docm/.pptm files are
//! read, listed, and (for data-processing macros, headlessly) executed,
//! without Microsoft Office.
//!
//! Two libraries it wraps (both MIT):
//! - `ovba` (tim-weis): a 100%-safe-Rust [MS-OVBA] parser. It extracts a
//!   VBA project's modules and their source from `vbaProject.bin`.
//! - `elixcee` (kent-tokyo): a headless Excel VBA runtime (its own VM,
//!   deterministic budgets by default). It loads a workbook, runs `Sub`
//!   entry points on its sheets, and saves the result.
//!
//! Scope mirrors elixcee's: common data-processing VBA (`If`/`For`/`Do`/
//! `Select Case`/`With`, arrays, `Range`/`Cells`, worksheet functions)
//! against .xlsx/.xlsm workbooks. Word/PowerPoint object models and UI
//! automation are out of scope here.
//!
//! Module kinds: ovba's [MS-OVBA] MODULETYPE record only separates
//! procedural modules from the document/class/designer umbrella. The
//! umbrella is split heuristically by each module's `Attribute VB_Base`
//! GUID, as [MS-OVBA] stores the host class a document module extends
//! (Excel's Workbook `00020819`/Worksheet `00020820`/`00020821`, Word's
//! `00020906`, PowerPoint's `91493462`) and VBA stores the generic
//! `Object` base (`FCFB3D2A…`) for class modules. `ThisWorkbook` is
//! always a document module.

use std::io::Read;

const VBA_PROJECT_PATH: &str = "xl/vbaProject.bin";

/// One VBA module read out of a macro-enabled Office file.
#[derive(Debug)]
pub struct VbaModule {
    /// The module name (as shown in the VBE's project tree).
    pub name: String,
    /// `standard`, `document`, `class` or `designer`.
    pub kind: &'static str,
    /// The module's source code, decoded with the project's code page.
    pub source: String,
}

/// Reads the VBA project of an OOXML package (`.xlsm`, `.docm`, `.pptm`…).
/// A file with no macros at all answers `Ok(vec![])`.
pub fn read_modules(package_path: &std::path::Path) -> Result<Vec<VbaModule>, String> {
    let file = std::fs::File::open(package_path)
        .map_err(|e| format!("cannot open {}: {e}", package_path.display()))?;
    let mut archive = zip::ZipArchive::new(file)
        .map_err(|e| format!("{} is not an Open XML package: {e}", package_path.display()))?;
    let Some(mut entry) = archive.by_name(VBA_PROJECT_PATH).ok() else {
        return Ok(Vec::new());
    };
    let mut raw = Vec::new();
    entry
        .read_to_end(&mut raw)
        .map_err(|e| format!("cannot read {VBA_PROJECT_PATH}: {e}"))?;
    let project = ovba::open_project(raw)
        .map_err(|e| format!("cannot parse the VBA project ({VBA_PROJECT_PATH}): {e}"))?;
    let mut modules = Vec::with_capacity(project.modules.len());
    for module in &project.modules {
        let source = project
            .module_source(&module.name)
            .map_err(|e| format!("cannot read the source of module {}: {e}", module.name))?;
        let kind = match module.module_type {
            ovba::ModuleType::Procedural => "standard",
            ovba::ModuleType::DocClsDesigner => classify_doc_cls(&module.name, &source),
        };
        modules.push(VbaModule { name: module.name.clone(), kind, source });
    }
    Ok(modules)
}

/// Splits ovba's document/class/designer umbrella by the module's
/// `Attribute VB_Base` GUID; see the crate documentation. `class` is the
/// fallback: data-processing macros lose little from a guessed class, but
/// a document module's host events matter to the run.
fn classify_doc_cls(name: &str, source: &str) -> &'static str {
    if name.eq_ignore_ascii_case("ThisWorkbook") {
        return "document";
    }
    // The bases the hosts name: Excel's Workbook (00020819) and
    // Worksheet (00020820/21), Word's Document (00020906/07),
    // PowerPoint's Presentation (91493462). VBA's generic class base
    // (FCFB3D2A…, the Object class) names every plain class module.
    let host_bases = ["00020819", "00020820", "00020821", "00020906", "00020907", "91493462"];
    let Some(line) = source.lines().find(|l| l.trim_start().starts_with("Attribute VB_Base")) else {
        return "class";
    };
    match line.split('{').nth(1).map(|guid| guid.chars().take(8).collect::<String>().to_ascii_uppercase()) {
        Some(base) if host_bases.contains(&base.as_str()) => "document",
        _ => "class",
    }
}

/// Compiles every module that can run: standard modules as they are,
/// class modules marked as classes (elixcee's field reclassification,
/// as its own CLI does for `.cls` files), and document modules so their
/// host events (`Workbook_Open`, `Worksheet_Change`...) are found. A
/// module elixcee cannot compile does not stop the run; it is reported
/// (with its reason) in the run's `skipped` instead.
fn compile_all(
    package_path: &std::path::Path,
) -> Result<(Vec<(String, elixcee::parser::ast::Program)>, Vec<String>), String> {
    let modules = read_modules(package_path)?;
    if modules.is_empty() {
        return Err("the file has no VBA project (no macros were stored in it)".into());
    }
    let (mut programs, mut skipped) = (Vec::new(), Vec::new());
    for module in modules {
        let compiled = elixcee::parser::parse(&module.source).map(|mut program| {
            if module.kind == "class" {
                // elixcee's parser cannot tell classes from standard
                // modules (the stream form has no header): mark them and
                // give the instances their fields, as the elixcee CLI
                // does for exported `.cls` files.
                program.is_class_module = true;
                let mut fields = program.module_variables;
                program.module_variables = Default::default();
                program.class_fields.append(&mut fields);
            }
            (module.name.to_lowercase(), program)
        });
        match compiled {
            Ok(compiled) => programs.push(compiled),
            Err(reason) => skipped.push(format!("{} {}", module.kind, module.name)),
        }
    }
    Ok((programs, skipped))
}

/// What a macro run left behind on the sheets and where.
#[derive(Debug, Default)]
pub struct RunReport {
    /// Sheet names the VM saw (as read from the file).
    pub sheets: Vec<String>,
    /// The text `MsgBox`/`Debug.Print` produced, if any.
    pub messages: Vec<String>,
    /// The modules the run went on without: those elixcee could not compile.
    pub skipped: Vec<String>,
    /// Path the result workbook was written to.
    pub saved: String,
}

/// A run's default instruction budget: elixcee's own deterministic
/// default. `--max-vba-instructions` (and the UI) pass what the user says.
pub const DEFAULT_MAX_VBA_INSTRUCTIONS: Option<u64> = Some(elixcee::vm::DEFAULT_MAX_VBA_INSTRUCTIONS);

/// Runs `entry` ("SubName" or "Module.SubName") inside elixcee's VM on a
/// copy of the workbook, writing the result to `output_path` (`.xlsx`,
/// `.xlsm` or `.ods`). No window is opened. The run may not exceed
/// elixcee's default instruction budget; pass another budget (or `None`
/// for an unbounded, trusted run) via [`run_macro_bounded`].
pub fn run_macro(
    package_path: &std::path::Path,
    entry: &str,
    output_path: &std::path::Path,
) -> Result<RunReport, String> {
    run_macro_bounded(package_path, entry, output_path, DEFAULT_MAX_VBA_INSTRUCTIONS)
}

/// [`run_macro`] with a caller-set instruction budget: `Some(n)` caps one
/// run at `n` instructions, `None` removes the cap for a trusted,
/// long-running macro. The VM still refuses external effects
/// (`CreateObject`, shell and file access, `Save`/`Close`) either way.
pub fn run_macro_bounded(
    package_path: &std::path::Path,
    entry: &str,
    output_path: &std::path::Path,
    max_instructions: Option<u64>,
) -> Result<RunReport, String> {
    let (programs, skipped) = compile_all(package_path)?;
    if programs.is_empty() {
        return Err(match skipped.first() {
            Some(first) => format!("no module could run (first failure: {first})"),
            None => "no module could run".into(),
        });
    }
    let mut vm = elixcee::vm::Vm::new();
    vm.max_instructions = max_instructions;
    let sheets = vm
        .load_workbook_file(&package_path.display().to_string())
        .map_err(|e| format!("cannot load the workbook for VBA: {e}"))?;
    vm.run_sub_multi(&programs, entry)
        .map_err(|e| format!("the macro failed: {e}"))?;
    let messages = [vm.take_messages(), vm.take_debug_output()].concat();
    elixcee::save_workbook(&vm, &output_path.display().to_string())
        .map_err(|e| format!("cannot save the result: {e}"))?;
    Ok(RunReport { sheets, messages, skipped, saved: output_path.display().to_string() })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn no_macros_answers_empty() {
        let path = std::env::temp_dir().join("ibis-vba-no-macros.docx");
        std::fs::write(&path, b"not really a package").unwrap();
        assert!(matches!(read_modules(&path), Err(_)));
    }
}
