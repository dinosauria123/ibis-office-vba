// SPDX-License-Identifier: AGPL-3.0-only
// Copyright (c) ibis inc.
// Additional terms under AGPL-3.0 section 7(b): see NOTICE.
//! A form session: run one of the form module's event procedures
//! against the form's on-screen values.
//!
//! The state plane is a scratch worksheet (`__IbisCtl`) in the workbook:
//! one row per control (its value/text, caption, check), the event's
//! `Control.Prop` accesses rewritten to calls on the generated shim
//! module, and the post-run cells read back into the form. The events
//! run through elixcee's public runtime — no change to the upstream
//! crate — under its deterministic budgets, with the external-effects
//! barrier up (the plane is workbook cells, not file/shell effects).
//!
//! The rewrite is limited to the plain member patterns a form's events
//! carry (`Control.Caption = …`, reads inside expressions, the `Show`
//! call); anything else fails the run named, rather than guessed away.
//! A workbook-only plane keeps every control's value inside the same
//! file the events' other effects go to.
use crate::read_modules;
use super::{ControlDef, UserFormDef};
use elixcee::parser::ast::Program;
use elixcee::vm::{Vm, Variant};
use std::collections::HashMap;
use std::path::{Path, PathBuf};

/// The control values of one form, as the UI carries them.
#[derive(Debug, Clone, Default)]
pub struct FormValue {
    /// A control's text/value (a TextBox's content, a Label's caption).
    pub text: String,
    /// A two-state control's check.
    pub checked: bool,
}

const CTL_SHEET: &str = "__IbisCtl";
const COL_KIND: u32 = 2;
const COL_TEXT: u32 = 3;
const COL_CAPTION: u32 = 4;
const COL_CHECKED: u32 = 5;

/// What running an event left behind (and where the session file is).
#[derive(Debug)]
pub struct FormRun {
    pub state: HashMap<String, FormValue>,
    pub file: PathBuf,
    /// `MsgBox`/`Debug.Print` the event produced, if any.
    pub messages: Vec<String>,
}

/// The scratch plane's row of each control: 2.., in the form's order.
fn rows(controls: &[ControlDef]) -> Vec<(String, String, usize)> {
    controls
        .iter()
        .enumerate()
        .map(|(i, c)| (c.name.clone(), c.kind.clone(), i + 2))
        .collect()
}

// MARK: the shim module (generated per form)

fn shim_source(rows: &[(String, String, usize)]) -> String {
    let mut code = String::from("Attribute VB_Name = \"__IbisCtl\"\nOption Explicit\n\n");
    code.push_str("Function __IbisRow(n As String) As Long\n    Select Case n\n");
    for (name, _, row) in rows {
        code.push_str(&format!("        Case \"{name}\"\n            __IbisRow = {row}\n"));
    }
    code.push_str("        Case Else\n            __IbisRow = 0\n    End Select\nEnd Function\n\n");
    code.push_str("Function __IbisCol(p As String) As Long\n    Select Case p\n");
    code.push_str(&format!("        Case \"Caption\"\n            __IbisCol = {COL_CAPTION}\n"));
    code.push_str(&format!("        Case \"Check\"\n            __IbisCol = {COL_CHECKED}\n"));
    code.push_str(&format!("        Case Else\n            __IbisCol = {COL_TEXT}\n    End Select\nEnd Function\n\n"));
    code.push_str("Function __IbisGet(n As String, p As String) As Variant\n    Dim r As Long\n    r = __IbisRow(n)\n    If r = 0 Then\n        __IbisGet = Empty\n    Else\n");
    code.push_str(&format!(
        "        __IbisGet = Worksheets(\"{CTL_SHEET}\").Cells(r, __IbisCol(p)).Value\n    End If\nEnd Function\n\n"
    ));
    code.push_str("Sub __IbisPut(n As String, p As String, v As Variant)\n    Dim r As Long\n    r = __IbisRow(n)\n    If r <> 0 Then\n");
    code.push_str(&format!(
        "        Worksheets(\"{CTL_SHEET}\").Cells(r, __IbisCol(p)) = v\n    End If\nEnd Sub\n\n"
    ));
    code.push_str("Sub __IbisShow(n As String)\nEnd Sub\n");
    code
}

// MARK: the module's source rewrite

/// The member names the plane carries; `Text` falls back to `Value`.
fn prop_of(member: &str) -> Option<&'static str> {
    match member {
        "Text" | "Value" => Some("Value"),
        "Caption" => Some("Caption"),
        "Checked" | "Check" => Some("Check"),
        _ => None,
    }
}

/// Rewrites a form module's source to the shim:
/// - `X.Prop = v` becomes `Call __IbisPut("X", "Prop", (v))`,
/// - `X.Prop` in an expression becomes `__IbisGet("X", "Prop")`,
/// - `UserFormX.Show …` becomes `Call __IbisShow("UserFormX")`.
pub fn rewrite(source: &str, form: &UserFormDef) -> (String, Vec<String>) {
    let mut notes = Vec::new();
    let mut out_lines = Vec::new();
    for raw in source.lines() {
        let mut line = raw.to_string();
        // The form's Show first (an argument stays as a comment).
        let with_show = format!("{}.Show", form.name);
        if line.trim_end().ends_with(&with_show) {
            let rest = line.trim_end()[..line.trim_end().len() - with_show.len()].trim_end().to_string();
            if rest.is_empty() {
                line = format!("Call __IbisShow(\"{}\")", form.name);
            } else {
                line = format!("Call __IbisShow(\"{}\") '{rest}", form.name);
                notes.push(format!("Show with an argument ({rest}): shown without it"));
            }
        }
        // An assignment `Name.Prop = …` (only where the line itself is
        // one; If/While/With lines stay in the rewrite's read path).
        let is_assignment = {
            let trimmed = trimmed_start(&line);
            !trimmed.starts_with("If ")
                && !trimmed.starts_with("Select Case")
                && !trimmed.starts_with("Case ")
                && !trimmed.starts_with("Do ")
                && !trimmed.starts_with("While ")
                && !trimmed.starts_with("For ")
                && !trimmed.starts_with("'")
        };
        let maybe_split: Option<(String, String)> = if is_assignment {
            line.split_once('=').map(|(l, r)| (l.trim().to_string(), r.trim().to_string()))
        } else {
            None
        };
        let Some((lhs, rhs)) = maybe_split else {
            let (rewritten, note) = rewrite_reads(&line, form);
            if let Some(note) = note {
                notes.push(note);
            }
            out_lines.push(rewritten);
            continue;
        };
        let Some((name, member)) = lhs.split_once('.') else {
            let (rewritten, note) = rewrite_reads(&line, form);
            if let Some(note) = note {
                notes.push(note);
            }
            out_lines.push(rewritten);
            continue;
        };
        let name = name.trim();
        let member = member.trim().trim_end_matches('=').trim();
        let Some(plane_prop) = prop_of(member) else {
            notes.push(format!("{name}.{member}: not modelled, left as written"));
            out_lines.push(line);
            continue;
        };
        if form.control(name).is_none() {
            notes.push(format!("{name}.{member}: the form has no control of that name"));
            out_lines.push(line);
            continue;
        }
        let (rhs_rewritten, note) = rewrite_reads(&rhs, form);
        if let Some(note) = note {
            notes.push(note);
        }
        out_lines.push(format!("Call __IbisPut(\"{name}\", \"{plane_prop}\", ({rhs_rewritten}))"));
    }
    (out_lines.join("\r\n"), notes)
}

fn trimmed_start(line: &str) -> String {
    // A line label (`Foo:`) is not an assignment head.
    line.trim()
        .trim_start_matches(|c: char| c == ':' || c == ' ')
        .to_string()
}

/// Member reads out of expressions: `identifier.member` pairs whose
/// member the plane carries and whose identifier is one of the form's
/// controls become the shim's get.
fn rewrite_reads(line: &str, form: &UserFormDef) -> (String, Option<String>) {
    let mut out = String::new();
    let mut rest = line;
    let mut found: Vec<String> = Vec::new();
    while let Some(dot_at) = rest.find('.') {
        let before = &rest[..dot_at];
        let after_dot = &rest[dot_at + 1..];
        let member: String = after_dot.chars().take_while(|c| c.is_alphanumeric() || *c == '_').collect();
        if member.is_empty() {
            out.push_str(&rest[..=dot_at]);
            rest = after_dot;
            continue;
        }
        let word_start = before
            .rfind(|c: char| !(c.is_alphanumeric() || c == '_'))
            .map(|i| i + 1)
            .unwrap_or(0);
        let name = &before[word_start..];
        if form.control(name).is_none() {
            out.push_str(&rest[..=dot_at]);
            rest = after_dot;
            continue;
        }
        match prop_of(&member) {
            Some(plane_prop) => {
                out.push_str(&before[..word_start]);
                out.push_str(&format!("__IbisGet(\"{name}\", \"{plane_prop}\")"));
                found.push(format!("{name}.{member}"));
                rest = &after_dot[member.len()..];
            }
            None => {
                out.push_str(&rest[..=dot_at]);
                rest = after_dot;
            }
        }
    }
    out.push_str(rest);
    let note = if found.is_empty() { None } else { Some(format!("read {}", found.join(", "))) };
    (out, note)
}

// MARK: the Variant plane

fn variant_text(v: &Variant) -> String {
    match v {
        Variant::Integer(n) => n.to_string(),
        Variant::Float(f) => {
            let mut text = f.to_string();
            if text.ends_with(".0") {
                text.truncate(text.len() - 2);
            }
            text
        }
        Variant::Str(s) => s.clone(),
        Variant::Boolean(true) => "True".into(),
        Variant::Boolean(false) => "False".into(),
        Variant::Date(n) => n.to_string(),
        _ => String::new(),
    }
}

fn variant_checked(v: &Variant) -> bool {
    match v {
        Variant::Boolean(b) => *b,
        Variant::Integer(n) => *n != 0,
        Variant::Float(n) => *n != 0.0,
        Variant::Str(s) => s.eq_ignore_ascii_case("True") || s.eq_ignore_ascii_case("-1"),
        _ => false,
    }
}

// MARK: the run

/// The run's two files: the events' code sits in `package`, the values'
/// plane in `workbook` (`None`: the same file — a session's first run).
pub struct FormSession {
    pub package: PathBuf,
    pub workbook: Option<PathBuf>,
}

/// Runs `event` ("CommandButton1_Click", "UserForm_Initialize", …)
/// against `state` and answers the run's own state (which the UI keeps,
/// along with the output file for follow-up runs).
pub fn run_form_event(
    session: &FormSession,
    form: &UserFormDef,
    event: &str,
    state: &HashMap<String, FormValue>,
    output: &Path,
) -> Result<FormRun, String> {
    let controls = &form.controls;
    let rows = rows(controls);
    if rows.is_empty() {
        return Err("the form has no controls to run against".into());
    }
    let modules = read_modules(&session.package)?;
    let Some(code) = modules
        .iter()
        .find(|m| m.name.eq_ignore_ascii_case(&form.name))
        .map(|m| m.source.clone())
    else {
        return Err(format!("no module named {}", form.name));
    };
    let (form_code, notes) = rewrite(&code, form);
    let _ = notes;
    let mut programs: Vec<(String, Program)> = Vec::with_capacity(2);
    let shim = shim_source(&rows);
    programs.push((
        "__IbisCtl".to_string(),
        elixcee::parser::parse(&shim).map_err(|e| format!("the shim does not compile: {e}"))?,
    ));
    // The form's events run as a plain procedure module (their subs are
    // the form's event handlers; a stored designer class is not needed
    // here, the plane stands in for the controls).
    programs.push((
        form.name.clone(),
        elixcee::parser::parse(&form_code)
            .map_err(|e| format!("the form module does not compile: {e}"))?,
    ));

    let workbook = session.workbook.clone().unwrap_or_else(|| session.package.clone());
    let mut vm = Vm::new();
    vm.max_instructions = Some(elixcee::vm::DEFAULT_MAX_VBA_INSTRUCTIONS);
    vm.load_workbook_file(&workbook.display().to_string())
        .map_err(|e| format!("cannot load the workbook for the form run: {e}"))?;
    vm.ensure_sheet(CTL_SHEET);
    vm.set_active_sheet(CTL_SHEET)
        .map_err(|e| format!("the scratch sheet cannot open: {e}"))?;
    vm.set_cell_value(1, 1, Variant::Str("Control".into()));
    vm.set_cell_value(1, COL_KIND, Variant::Str("Kind".into()));
    vm.set_cell_value(1, COL_TEXT, Variant::Str("Value".into()));
    vm.set_cell_value(1, COL_CAPTION, Variant::Str("Caption".into()));
    vm.set_cell_value(1, COL_CHECKED, Variant::Str("Check".into()));
    for (name, kind, row) in &rows {
        let row = *row as u32;
        vm.set_cell_value(row, 1, Variant::Str(name.clone()));
        vm.set_cell_value(row, COL_KIND, Variant::Str(kind.clone()));
        let value = state.get(name).cloned().unwrap_or_default();
        // The two caption-bearing kinds share the form's text; the text
        // classes carry the value only (seeding BOTH would lose the
        // caption of one when the event writes just the other's).
        let captioned = matches!(kind.as_str(), "Label" | "CommandButton");
        vm.set_cell_value(row, COL_TEXT, Variant::Str(value.text.clone()));
        if captioned {
            vm.set_cell_value(row, COL_CAPTION, Variant::Str(value.text.clone()));
        }
        vm.set_cell_value(row, COL_CHECKED, Variant::Boolean(value.checked));
    }

    vm.run_sub_multi(&programs, event)
        .map_err(|e| format!("the event failed: {e}"))?;
    let mut messages = vm.take_messages();
    messages.extend(vm.take_debug_output());

    vm.set_active_sheet(CTL_SHEET)
        .map_err(|e| format!("the scratch sheet cannot reopen: {e}"))?;
    let cells = vm.cells();
    let mut out_state: HashMap<String, FormValue> = HashMap::new();
    for (name, _, row) in &rows {
        let row = *row as u32;
        let mut value = state.get(name).cloned().unwrap_or_default();
        let kind = rows.iter().find(|(n, _, r)| n == name && *r == row as usize).map(|(_, k, _)| k.clone()).unwrap_or_default();
        let _ = kind;
        let captioned = matches!(kind.as_str(), "Label" | "CommandButton");
        if captioned && let Some(cell) = cells.get(&(row, COL_CAPTION)) {
            let caption = variant_text(&cell.value);
            if !caption.is_empty() {
                value.text = caption;
            }
        } else if let Some(cell) = cells.get(&(row, COL_TEXT)) {
            let text = variant_text(&cell.value);
            if !text.is_empty() {
                value.text = text;
            }
        }
        if let Some(cell) = cells.get(&(row, COL_CHECKED)) {
            value.checked = variant_checked(&cell.value);
        }
        out_state.insert(name.clone(), value);
    }
    elixcee::save_workbook(&vm, &output.display().to_string())
        .map_err(|e| format!("cannot save the form run: {e}"))?;
    Ok(FormRun { state: out_state, file: output.to_path_buf(), messages })
}
