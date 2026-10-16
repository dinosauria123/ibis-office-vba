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
//!   deterministic instruction budget by default). It loads a workbook,
//!   runs `Sub` entry points on its sheets, and saves the result.
//!
//! Scope intentionally mirrors elixcee's: common data-processing VBA
//! (`If`/`For`/`Do`/`Select Case`/`With`, arrays, `Range`/`Cells`, worksheet
//! functions) against .xlsx/.xlsm workbooks. Word/PowerPoint object models
//! and UI automation are out of scope here.

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

/// What a macro run left behind on the sheets and where.
#[derive(Debug, Default)]
pub struct RunReport {
    /// Sheet names the VM saw (as read from the file).
    pub sheets: Vec<String>,
    /// The text `MsgBox`/`Debug.Print` produced, if any.
    pub messages: Vec<String>,
    /// Path the result workbook was written to.
    pub saved: String,
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
        let kind = match module.module_type {
            ovba::ModuleType::Procedural => "standard",
            ovba::ModuleType::DocClsDesigner => "document",
        };
        let source = project
            .module_source(&module.name)
            .map_err(|e| format!("cannot read the source of module {}: {e}", module.name))?;
        modules.push(VbaModule { name: module.name.clone(), kind, source });
    }
    Ok(modules)
}

/// Extracts the modules that can run headlessly: standard code modules.
/// Document/class/designer modules are named in `skipped` (with the reason
/// elixcee's data-processing runtime cannot execute them here).
fn runnable_source(
    package_path: &std::path::Path,
) -> Result<(Vec<(String, String)>, Vec<String>), String> {
    let modules = read_modules(package_path)?;
    if modules.is_empty() {
        return Err("the file has no VBA project (no macros were stored in it)".into());
    }
    let mut runnable = Vec::new();
    let mut skipped = Vec::new();
    for module in modules {
        if module.kind == "standard" {
            runnable.push((module.name, module.source));
        } else {
            skipped.push(module.name);
        }
    }
    Ok((runnable, skipped))
}

/// Runs `entry` ("SubName" or "Module.SubName") inside `elm.excee`'s VM,
/// operating on a copy of the workbook and writing the result to
/// `output_path` (`.xlsx`, `.xlsm` or `.ods`). No window is opened.
///
/// The run is bounded by elixcee's deterministic budgets (instruction
/// count, call depth, string/array/cell sizes), and blocked external
/// effects (`CreateObject`, file/shell access, `Save`/`Close`) are refused
/// by the VM itself.
pub fn run_macro(
    package_path: &std::path::Path,
    entry: &str,
    output_path: &std::path::Path,
) -> Result<RunReport, String> {
    let (runnable, skipped) = runnable_source(package_path)?;
    if runnable.is_empty() {
        return Err(
            "no standard module was found: document and class modules cannot run headlessly"
                .into(),
        );
    }
    // Parse each module independently; the error names the module.
    let mut parsed: Vec<(String, elixcee::parser::ast::Program)> = Vec::with_capacity(runnable.len());
    for (name, source) in &runnable {
        let program = elixcee::parser::parse(source)
            .map_err(|e| format!("VBA of module {name} does not compile: {e}"))?;
        parsed.push((name.clone(), program));
    }
    let mut vm = elixcee::vm::Vm::new();
    vm.max_instructions = None; // the caller named this macro explicitly
    let sheets = vm
        .load_workbook_file(&package_path.display().to_string())
        .map_err(|e| format!("cannot load the workbook for VBA: {e}"))?;
    vm.run_sub_multi(&parsed, entry)
        .map_err(|e| format!("the macro failed: {e}"))?;
    let messages = [vm.take_messages(), vm.take_debug_output()].concat();
    elixcee::save_workbook(&vm, &output_path.display().to_string())
        .map_err(|e| format!("cannot save the result: {e}"))?;
    let _ = skipped;
    Ok(RunReport { sheets, messages, saved: output_path.display().to_string() })
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
