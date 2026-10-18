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

pub mod forms;

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
        // An exported module may declare dotted, object-library types
        // (`As MSForms.ReturnBoolean`); elixcee's parser stops on them,
        // so a Variant stand-in goes in first.
        let scrubbed = forms::run::scrub_unknown_types(&module.source);
        let compiled = elixcee::parser::parse(&scrubbed).map(|mut program| {
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
    /// Where the result workbook was written (`None`: the run reported
    /// its messages only and wrote nothing).
    pub saved: Option<String>,
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
    run_macro_bounded(package_path, entry, Some(output_path), DEFAULT_MAX_VBA_INSTRUCTIONS)
}

/// [`run_macro`] with a caller-set instruction budget: `Some(n)` caps one
/// run at `n` instructions, `None` removes the cap for a trusted,
/// long-running macro. The VM still refuses external effects
/// (`CreateObject`, shell and file access, `Save`/`Close`) either way.
pub fn run_macro_bounded(
    package_path: &std::path::Path,
    entry: &str,
    output_path: Option<&std::path::Path>,
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
    vm.run_sub_multi(&programs, entry).map_err(|e| {
        // A not-found entry the file actually declares names the
        // modules that did not compile (their declarations are there;
        // the parse stopped before them).
        if e.contains("not found") {
            let wanted = entry.split('.').next_back().unwrap_or(entry);
            let owners: Vec<String> = crate::read_modules(package_path)
                .unwrap_or_default()
                .iter()
                .filter(|m| !programs.iter().any(|(name, _)| name.eq_ignore_ascii_case(&m.name)))
                .filter(|m| {
                    m.source.to_lowercase().contains(&format!(
                        "sub {}(",
                        wanted.to_lowercase()
                    ))
                })
                .map(|m| m.name.clone())
                .collect();
            if !owners.is_empty() {
                let word = if owners.len() == 1 { "module" } else { "modules" };
                return format!(
                    "the macro failed: {e} (the {} {word} declaring it did not compile — run it through the form window, or simplify the module's declarations)",
                    owners.join(", ")
                );
            }
            format!("the macro failed: {e}")
        } else {
            format!("the macro failed: {e}")
        }
    })?;
    let messages = [vm.take_messages(), vm.take_debug_output()].concat();
    let saved = match output_path {
        Some(out) => {
            elixcee::save_workbook(&vm, &out.display().to_string())
                .map_err(|e| format!("cannot save the result: {e}"))?;
            Some(out.display().to_string())
        }
        None => None,
    };
    Ok(RunReport { sheets, messages, skipped, saved })
}


// ── save-back (Excel-style overwrite) ────────────────────────────────────────

/// Save the run's workbook **over the original file**, Excel-style: the
/// cells the run wrote replace the original's, and — for a macro
/// workbook (`.xlsm`) — the file keeps its VBA project (elixcee's save
/// carries `xl/vbaProject.bin` through when the output name is
/// macro-enabled). The write is atomic: a temporary package beside the
/// original is renamed over it, and the plane's scratch sheet rides
/// along, so the next run's state chains from the saved file.
///
/// The workbook is regenerated by the runtime, so the formatting it
/// does not model simplifies; the values, the sheets and the VBA
/// project survive.
pub fn save_back_to_package(original: &std::path::Path, vm: &elixcee::vm::Vm) -> Result<(), String> {
    let ext = match original.extension().and_then(|e| e.to_str()).map(|e| e.to_ascii_lowercase()) {
        Some(ext) if ext == "xlsx" || ext == "xlsm" => ext,
        _ => return Err(format!("save-back supports .xlsx and .xlsm only: {}", original.display())),
    };
    let dir = original.parent().filter(|p| !p.as_os_str().is_empty()).unwrap_or(std::path::Path::new("."));
    let stem = original
        .file_stem()
        .and_then(|s| s.to_str())
        .ok_or_else(|| format!("the file has no name: {}", original.display()))?;
    let temp = dir.join(format!(".{stem}-ibis-save.{ext}"));
    elixcee::save_workbook(vm, &temp.display().to_string())?;
    std::fs::rename(&temp, original)
        .map_err(|e| format!("cannot replace {}: {e}", original.display()))?;
    Ok(())
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
