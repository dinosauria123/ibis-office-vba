// SPDX-License-Identifier: AGPL-3.0-only
// Copyright (c) ibis inc.
// Additional terms under AGPL-3.0 section 7(b): see NOTICE.
//! A form session: run one of the form module's event procedures
//! against the form's on-screen values.
//!
//! The state plane is a scratch worksheet (`__IbisCtl`) in the workbook:
//! one row per control (its value/text, caption, check and — for list
//! controls — the selection), the event's `Control.Prop` accesses
//! rewritten to calls on the generated shim module, and the post-run
//! cells read back into the form. The events run through elixcee's
//! public runtime — no change to the upstream crate — under its
//! deterministic budgets, with the external-effects barrier up (the
//! plane is workbook cells, not file/shell effects).
//!
//! The rewrite's tables cover the patterns the demo form modules carry:
//! control member chains (`Control.Caption = …`, reads inside
//! expressions), a form-object reference (`Set x = oForm`, then
//! `x.Control`), aListBox's `ListIndex`/`ListCount`/`List(r, c)`/
//! `AddItem`/`Clear`/`ColumnCount` reads and writes, and a table plane
//! (`Set t = Worksheets("Data").ListObjects("TableData")`, then
//! `t.Range.Rows.Count`, `t.ListRows.Add/n.dataGridView`, and the row's
//! `Range(n) = v` writes inside the With). A pattern with no rule below
//! keeps its line as written; the run's own compile then reports
//! anything that does not hold together.
//!
//! `Give a form data's head` — the workbook's own ListObjects xml
//! carries the range the table spans, which the rewriter bakes into the
//! absolute cells the shim writes ([MS-OFORMS]/OOXML 2.2 tables).
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
    /// A list control's rows, one line per row (the cells joined with a
    /// tab).
    pub list: Vec<String>,
    /// The row the list plane has selected (`ListIndex`).
    pub selected: Option<usize>,
}

const CTL_SHEET: &str = "__IbisCtl";
const COL_KIND: u32 = 2;
const COL_TEXT: u32 = 3;
const COL_CAPTION: u32 = 4;
const COL_CHECKED: u32 = 5;
const COL_LIST_INDEX: u32 = 11;
/// The list rows' area: row 50 + list row, column 5 + column index —
/// one column per member (five here).
const LIST_ROW_BASE: usize = 50;
/// The list rows go in per-control slots (a form may carry several
/// list controls — three ComboBoxes fill a wide one), 64 rows each; a
/// control's slot starts at `LIST_ROW_BASE + (row - 2) * SLOT_ROWS`.
const SLOT_ROWS: usize = 64;
const LIST_COL_BASE: u32 = 6;
/// `selected: <n>` rides the state answer as line 1.
const LIST_MAX_COLS: usize = 5;
const LIST_MAX_ROWS: usize = 64;

/// A list control's slot head on the plane: its own region of rows
/// (the shim calls pass it, so several list controls coexist).
fn list_base(control_row: usize) -> usize {
    LIST_ROW_BASE + (control_row.saturating_sub(2)) * SLOT_ROWS
}

/// The kinds the plane carries as list controls (a ListBox and its
/// single-line cousin, the ComboBox).
fn is_list_kind(kind: &str) -> bool {
    kind == "ListBox" || kind == "ComboBox"
}

/// A module function's zero-arg call the source spells without parens
/// (`BuatID`): elixcee answers those as variables (an empty cell), so
/// the calls gain the parens. The function's own head, assignment and
/// `End` lines, Attribute rows and comments leave.
fn add_zero_arg_parens(body: &str, extra_fns: &[String]) -> String {
    let mut fn_names: Vec<String> = extra_fns.to_vec();
    for line in body.lines() {
        let trimmed = line.trim_start();
        let head = trimmed.to_lowercase();
        if head.starts_with("function ") || head.starts_with("private function ") || head.starts_with("public function ") {
            // The name follows the `Function` keyword: a `Private` or
            // `Public` modifier shifts one token over.
            let after_keyword: &str = after_modifier(trimmed);
            let name: String = after_keyword
                .split_whitespace()
                .nth(1)
                .map(|w| w.chars().take_while(|c| *c == '_' || c.is_alphanumeric()).collect::<String>())
                .unwrap_or_default();
            // A zero-arg declaration only: the first parens, when the
            // head has one, hold nothing but whitespace.
            let empty = trimmed.find('(').map(|at| {
                let inner = &trimmed[at + 1..];
                inner.split(')').next().unwrap_or(message_placeholder()).trim().is_empty()
            });
            if !name.is_empty() && empty.unwrap_or(true) && !fn_names.iter().any(|n| n.eq_ignore_ascii_case(&name)) {
                fn_names.push(name);
            }
        }
    }
    if fn_names.is_empty() {
        return body.to_string();
    }
    let mut spaced: Vec<String> = Vec::new();
    for line in body.lines() {
        let head = line.trim_start();
        let own_assignment = fn_names
            .iter()
            .any(|name| head.to_lowercase().starts_with(&name.to_lowercase()));
        let kept = line.starts_with("Attribute ")
            || line.trim().starts_with('\'')
            || own_assignment
            || {
                let lh = head.to_lowercase();
                lh.starts_with("function ")
                    || lh.starts_with("sub ")
                    || lh.starts_with("private ")
                    || lh.starts_with("public ")
                    || lh.starts_with("end ")
            };
        let mut out = line.to_string();
        if !kept {
            // A Call line's callee keeps its shape; the arguments and
            // other name positions gain the parens.
            let search_from: usize = if out.to_lowercase().starts_with("call ") {
                out.find('(').unwrap_or(out.len())
            } else {
                0
            };
            for name in &fn_names {
                let bare = name.to_lowercase();
                let mut scan = search_from;
                while let Some(ix) = out[scan..].to_lowercase().find(&bare) {
                    let at = scan + ix;
                    let lower = out.to_lowercase();
                    let after_is_paren = lower[at + name.len()..].starts_with('(');
                    let before_ok = at == 0
                        || !(out[..at].chars().next_back().is_some_and(|c: char| c.is_alphanumeric() || c == '_'));
                    let after_ok = out[at + name.len()..]
                        .chars()
                        .next()
                        .is_none_or(|c: char| !(c.is_alphanumeric() || c == '_'));
                    if after_is_paren || !(before_ok && after_ok) {
                        scan = at + name.len();
                        continue;
                    }
                    let tail = out[at + name.len()..].to_string();
                    out = format!("{}{}(){}", &out[..at], name, tail);
                    scan = at + name.len() + 2;
                }
            }
        }
        spaced.push(out);
    }
    spaced.join("\r\n")
}


/// A Sheet (quoted) binding's name without its quotes; the `Name`
/// member's literal lives at the binding.
fn message_placeholder() -> &'static str {
    ""
}

/// A module's zero-arg `Function` declarations (the calls the form's
/// code may spell bare).
fn zero_arg_fns(source: &str) -> Vec<String> {
    let mut names = Vec::new();
    for line in source.lines() {
        let trimmed = line.trim_start();
        let head = trimmed.to_lowercase();
        if !(head.starts_with("function ") || head.starts_with("private function ") || head.starts_with("public function ")) {
            continue;
        }
        let after = after_modifier(trimmed);
        let Some(name) = after.split_whitespace().nth(1).map(|w| {
            w.chars().take_while(|c| *c == '_' || c.is_alphanumeric()).collect::<String>()
        }) else { continue };
        let empty = trimmed.find('(').map(|at| {
            let inner = &trimmed[at + 1..];
            inner.split(')').next().unwrap_or("").trim().is_empty()
        });
        if !name.is_empty() && empty.unwrap_or(true) {
            names.push(name);
        }
    }
    names
}

/// The text after a `Private`/`Public`/`Static` modifier, else as is.
fn after_modifier(line: &str) -> &str {
    let lowered = line.to_lowercase();
    for m in ["private ", "public ", "static "] {
        if lowered.starts_with(m) {
            return &line[m.len()..];
        }
    }
    line
}

fn binding_sheet_name(binding: &Binding) -> String {
    match binding {
        Binding::Sheet { name_expr, .. } => name_expr.trim_matches('"').to_string(),
        _ => String::new(),
    }
}

/// What running an event left behind (and where the session file is).
#[derive(Debug)]
pub struct FormRun {
    pub state: HashMap<String, FormValue>,
    pub file: PathBuf,
    /// `MsgBox`/`Debug.Print` the event produced, if any.
    pub messages: Vec<String>,
}

/// The run's two files: the events' code sits in `package`, the values'
/// plane in `workbook` (`None`: the same file — a session's first run).
pub struct FormSession {
    pub package: PathBuf,
    pub workbook: Option<PathBuf>,
    /// Excel-style save: the run's result replaces the original file
    /// (its VBA project stays; the plane rides along for chaining).
    /// Off: the copy beside the original (`-form-run.xlsx`) takes it.
    pub save_back: bool,
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

fn shim_source(
    rows: &[(String, String, usize)],
    table: &Option<(String, String, usize, usize, usize)>,
    sheets: &[String],
) -> String {
    let has_list = rows.iter().any(|(_, kind, _)| is_list_kind(kind));
    let mut code = String::from("Attribute VB_Name = \"__IbisCtl\"\nOption Explicit\n\n");
    code.push_str("Function __IbisRow(n As String) As Long\n    Select Case n\n");
    for (name, _, row) in rows {
        code.push_str(&format!("        Case \"{name}\"\n            __IbisRow = {row}\n"));
    }
    code.push_str("        Case Else\n            __IbisRow = 0\n    End Select\nEnd Function\n\n");
    code.push_str("Function __IbisCol(p As String) As Long\n    Select Case p\n");
    code.push_str(&format!("        Case \"Caption\"\n            __IbisCol = {COL_CAPTION}\n"));
    code.push_str(&format!("        Case \"Check\"\n            __IbisCol = {COL_CHECKED}\n"));
    code.push_str(&format!("        Case \"ListIndex\"\n            __IbisCol = {COL_LIST_INDEX}\n"));
    code.push_str(&format!(
        "        Case Else\n            __IbisCol = {COL_TEXT}\n    End Select\nEnd Function\n\n"
    ));
    code.push_str("Function __IbisGet(n As String, p As String) As Variant\n    Dim r As Long\n    r = __IbisRow(n)\n    If r = 0 Then\n        __IbisGet = Empty\n    Else\n");
    code.push_str(&format!(
        "        __IbisGet = Worksheets(\"{CTL_SHEET}\").Cells(r, __IbisCol(p)).Value\n    End If\nEnd Function\n\n"
    ));
    code.push_str("Sub __IbisPut(n As String, p As String, v As Variant)\n    Dim r As Long\n    r = __IbisRow(n)\n    If r <> 0 Then\n");
    code.push_str(&format!(
        "        Worksheets(\"{CTL_SHEET}\").Cells(r, __IbisCol(p)) = v\n    End If\nEnd Sub\n\n"
    ));
    code.push_str("Sub __IbisShow(n As String)\nEnd Sub\n\n");
    // The workbook's sheets by index (the session read them in the
    // load's order); the demos that enumerate `Worksheets(i).Name` run
    // on this list (elixcee's Worksheets itself takes only literals).
    code.push_str("Function __IbisSheetName(n As Long) As String\n    Select Case n\n");
    for (ix, sheet) in sheets.iter().enumerate() {
        code.push_str(&format!("        Case {}\n            __IbisSheetName = \"{}\"\n", ix + 1, sheet));
    }
    code.push_str("        Case Else\n            __IbisSheetName = \"\"\n    End Select\nEnd Function\n\n");
    // The table plane: the shim carries the session table's literal
    // sheet (the row/col math bakes in the rewrite; a session's shim is
    // per-form).
    if let Some((_, table_sheet, table_head, _, _)) = table {
        let table_sheet = table_sheet.to_lowercase();
        let _ = table_sheet;
        code.push_str(&format!(
            "Function __IbisTableHead() As Long\n    __IbisTableHead = {table_head}\nEnd Function\n\n"
        ));
        code.push_str("Function __IbisDataRows() As Long\n    Dim r As Long\n    r = 0\n    Do\n");
        code.push_str(&format!(
            "        If Worksheets(\"{table_sheet}\").Cells(__IbisTableHead() + 1 + r, 1).Value = vbNullString Then\n            Exit Do\n        End If\n        r = r + 1\n    Loop While r < 500\n    __IbisDataRows = r\nEnd Function\n\n"
        ));
        code.push_str("Sub __IbisDataPut(row As Long, col As Long, v As Variant)\n");
        code.push_str(&format!(
            "    Worksheets(\"{table_sheet}\").Cells(row, col) = v\nEnd Sub\n\n"
        ));
        code.push_str("Function __IbisDataGet(row As Long, col As Long) As Variant\n");
        code.push_str(&format!(
            "    __IbisDataGet = Worksheets(\"{table_sheet}\").Cells(row, col).Value\nEnd Function\n\n"
        ));
        let dsh = table_sheet.replace('"', "");
        code.push_str(&format!(
            "Sub __IbisDataDeleteRow(row As Long)\n    Dim r As Long\n    Dim c As Long\n    r = row\n    Do\n        For c = 1 To 5\n            Worksheets(\"{dsh}\").Cells(r, c) = Worksheets(\"{dsh}\").Cells(r + 1, c).Value\n        Next c\n        If Worksheets(\"{dsh}\").Cells(r + 1, 1).Value = vbNullString Then\n            Exit Do\n        End If\n        r = r + 1\n    Loop While r < __IbisTableHead() + 500\nEnd Sub\n\n"
        ));
    }
    code.push_str("Function __IbisMin(a As Variant, b As Variant) As Variant\n    If a < b Then\n        __IbisMin = a\n    Else\n        __IbisMin = b\n    End If\nEnd Function\n\n");
    if has_list {
        code.push_str("Sub __IbisListPut(base As Long, r As Long, c As Long, v As Variant)\n");
        code.push_str(&format!(
            "    Worksheets(\"{CTL_SHEET}\").Cells(base + r, {LIST_COL_BASE} + c) = v\nEnd Sub\n\n"
        ));
        code.push_str("Function __IbisListGet(base As Long, r As Long, c As Long) As Variant\n");
        code.push_str(&format!(
            "    __IbisListGet = Worksheets(\"{CTL_SHEET}\").Cells(base + r, {LIST_COL_BASE} + c).Value\nEnd Function\n\n"
        ));
        code.push_str("Sub __IbisListClear(base As Long)\n    Dim r As Long\n    Dim c As Long\n    For r = 0 To 63\n        For c = 0 To 4\n            Worksheets(\"__IbisCtl\").Cells(base + r, 6 + c) = vbNullString\n        Next c\n    Next r\nEnd Sub\n\n");
        code.push_str("Function __IbisListRows(base As Long) As Long\n    Dim r As Long\n    r = 0\n    Do\n        If __IbisListRowEmpty(base, r) Then\n            Exit Do\n        End If\n        r = r + 1\n    Loop While r < 63\n    __IbisListRows = r\nEnd Function\n\n");
        code.push_str("Function __IbisListRowEmpty(base As Long, r As Long) As Boolean\n    Dim c As Long\n    Dim any As Boolean\n    any = False\n    For c = 0 To 4\n        If Worksheets(\"__IbisCtl\").Cells(base + r, 6 + c).Value <> vbNullString Then\n            any = True\n        End If\n    Next c\n    __IbisListRowEmpty = Not any\nEnd Function\n\n");
        // `AddItem expr`: the next free row's first column (the writers
        // that follow fill the row's other columns).
        code.push_str("Sub __IbisListAddItem(base As Long, v As Variant)\n    Dim r As Long\n    r = __IbisListRows(base)\n    Worksheets(\"__IbisCtl\").Cells(base + r, 6) = v\nEnd Sub\n\n");
        code.push_str("Function __IbisMax(a As Variant, b As Variant) As Variant\n    If a < b Then\n        __IbisMax = b\n    Else\n        __IbisMax = a\n    End If\nEnd Function\n\n");
    }
    code
}

// MARK: the module's source rewrite

#[derive(Debug, Clone)]
enum Binding {
    /// A form object (`Set myUserForm = oUserForm`): chains through it
    /// re-route to the control plane.
    Form,
    /// A worksheet object (`Set ws = ThisWorkbook.Worksheets("Data")` —
    /// or an indexed `Worksheets(i)`, the name a runtime expression):
    /// its `Name` answers literal (when quoted) and its Cells/Range
    /// re-qualify.
    Sheet { name_expr: String, quoted: bool },
    /// An indexed worksheet spread across iterations (`Worksheets(i)`):
    /// elixcee's Worksheets takes a literal, so only the `Name` reads
    /// survive — through the session's per-index Case list; the
    /// Cells/Range chains drop with a note.
    SheetIndexed { index_expr: String },
    /// A ListObject: its sheet and the range the table says.
    Table { sheet: String, head_row: usize, head_col: usize, cols: usize },
    /// A ListRow carrier: writes go to the table's absolute cells at
    /// the row expression the carrier holds.
    Row { table: String, head_row: usize, head_col: usize, row_expr: String },
}

struct ModuleRewrite {
    vars: HashMap<String, Binding>,
    form: UserFormDef,
    notes: Vec<String>,
    /// The session's other standard modules (a `Call mCode.Read` is
    /// theirs; elixcee's cross-module resolution takes the bare name).
    module_names: Vec<String>,
    /// The variables whose `Set` lines dropped (a dropped Set means the
    /// variable's later chains cannot run: those lines drop too, so
    /// the module still parses and the run reports what it leaned on).
    dropped_vars: std::collections::HashSet<String>,
}

/// Rewrites one module source for the form session. `states` bools the
/// run's own registrations: the form's and its locals.
impl ModuleRewrite {
    fn new(form: &UserFormDef) -> Self {
        Self { vars: Default::default(), form: form.clone(), notes: Vec::new(), module_names: Vec::new(), dropped_vars: Default::default() }
    }

    fn note(&mut self, note: String) {
        self.notes.push(note);
    }
}

/// The rewrite's own per-line tables; returns whether the line stays.
fn keep_as_written(line: &str) -> bool {
    line.starts_with("Attribute ")
        || line.starts_with("Option ")
        || trim_line(line).is_empty()
        || trim_line(line).starts_with("'")
}

fn trim_line(line: &str) -> &str {
    line.split(':').next().unwrap_or(line).trim()
}

/// A declaration's dotted type (an object library's `X.Y`): a Variant
/// stand-in parses (the procedures here never read the object). The
/// plain macro path scrubs a module's declarations before its own
/// compile; the form session's rewrite folds this into its pass.
pub fn scrub_unknown_types(source: &str) -> String {
    let mut out: Vec<String> = Vec::new();
    for raw in source.lines() {
        let mut line = raw.to_string();
        let head = line.trim_start().to_lowercase();
        if head.starts_with("dim ") {
            line = dim_of(&line);
        } else if head.starts_with("private sub ")
            || head.starts_with("public sub ")
            || head.starts_with("sub ")
        {
            if line.contains(" As ") {
                let as_at = line.find(" As ").unwrap();
                let after = line[as_at + 4..].to_string();
                let dotted_len = after
                    .find(|c: char| c == ')' || c == ',')
                    .unwrap_or(after.len());
                if after[..dotted_len].contains('.') {
                    line = format!("{}Variant{}", &line[..as_at + 4], &after[dotted_len..]);
                }
            }
        }
        out.push(line);
    }
    out.join("\r\n")
}

/// Drops the `As <type>` suffix from a `Dim` when the type is dotted
/// (`Excel.listObject`) or one elixcee's object set does not carry
/// (`Worksheet`, `Range`, `CommandBar` — the session models their
/// engine through the shim, not their type objects).
fn dim_of(source_line: &str) -> String {
    if let Some(as_at) = source_line.find(" As ") {
        let mut after = &source_line[as_at + 4..];
        if after.len() > 1 && after.ends_with(',') {
            after = &after[..after.len() - 1];
        }
        let type_name: String = after
            .chars()
            .take_while(|c| c.is_alphanumeric() || *c == '.' || *c == '_')
            .collect();
        let builtin = [
            "variant", "string", "long", "integer", "boolean", "double", "single", "byte", "date",
            "currency", "object",
        ];
        let known = builtin.iter().any(|b| type_name.to_lowercase() == *b);
        if after.contains('.') || (!known && !type_name.is_empty()) {
            return format!("{} As Variant", &source_line[..as_at]);
        }
    }
    source_line.to_string()
}

/// The `Set <var> = <rhs>` line: registers and answers whether the
/// source stays (the rewritten setter lines are dropped after).
fn set_of(
    line: &str,
    vars: &mut HashMap<String, Binding>,
    form: &UserFormDef,
    table: &Option<(String, String, usize, usize, usize)>,
    notes: &mut Vec<String>,
) -> Option<bool> {
    let rest = line.strip_prefix("Set ")?;
    let (name, expr) = rest.split_once('=')?;
    let name = name.trim();
    let expr = expr.trim();
    let expr = expr.strip_prefix("ThisWorkbook.").unwrap_or(expr);
    if expr == form.name {
        vars.insert(name.to_string(), Binding::Form);
        notes.push(format!("Set {name} = {} (form reference; the lines through it re-route)", form.name));
        return Some(false);
    }
    // The object's chain, through one dot at most for binding vars.
    if let Some(pos) = expr.find("ListObjects(") {
        let table_name = expr[pos..]
            .trim_start_matches("ListObjects(")
            .split(')')
            .next()
            .unwrap_or("")
            .trim_matches('"');
        let sheet = expr
            .find("Worksheets(")
            .and_then(|at| {
                expr[at..]
                    .trim_start_matches("Worksheets(")
                    .split(')')
                    .next()
                    .map(|s| s.trim_matches('"').to_string())
            })
            .unwrap_or_else(|| form.name.clone());
        let (head_row, head_col, cols) = table
            .as_ref()
            .filter(|(name, ..)| name == table_name)
            .map(|(_, _, head_row, head_col, cols)| (*head_row, *head_col, *cols))
            .unwrap_or((1, 1, 1));
        vars.insert(
            name.to_string(),
            Binding::Table { sheet: sheet.to_lowercase(), head_row, head_col, cols },
        );
        notes.push(format!(
            "Set {name} = ListObjects(\"{table_name}\") on sheet {sheet}: the table's plane is its cells"
        ));
        return Some(false);
    }
    // A worksheet binding: `Worksheets("<literal>")` bakes a literal in;
    // an indexed `Worksheets(i)` keeps the runtime expression (the
    // Name reads ride the shim's sheet list).
    if expr.contains("Worksheets(") {
        let at = expr.find("Worksheets(").unwrap();
        let arg = expr[at + "Worksheets(".len()..]
            .split(')')
            .next()
            .unwrap_or("")
            .trim()
            .to_string();
        if !arg.is_empty() {
            if arg.starts_with('"') {
                let literal = arg.trim_matches('"').to_string();
                vars.insert(
                    name.to_string(),
                    Binding::Sheet { name_expr: format!("\"{literal}\""), quoted: true },
                );
                notes.push(format!(
                    "Set {name} = Worksheets(\"{literal}\"): the sheet's chains re-qualify"
                ));
            } else {
                vars.insert(
                    name.to_string(),
                    Binding::SheetIndexed { index_expr: arg.clone() },
                );
                notes.push(format!(
                    "Set {name} = Worksheets({arg}): an indexed sheet — its Name reads the session's sheet list"
                ));
            }
            return Some(false);
        }
    }
    None
}

/// The `Set <var> = <rowvar>.ListRows…` lines, registered after the
/// table binding (their own pass).
fn set_row(
    line: &str,
    vars: &mut HashMap<String, Binding>,
    notes: &mut Vec<String>,
) -> Option<bool> {
    let rest = line.strip_prefix("Set ")?;
    let (name, expr) = rest.split_once('=')?;
    let name = name.trim();
    let expr = expr.trim();
    // The carried variable's table binding.
    let carrier = vars.keys().find(|key| expr.starts_with(*key)).cloned()?;
    let Some(Binding::Table { sheet, head_row, head_col, cols }) = vars.get(&carrier).cloned() else {
        return None;
    };
    let after = expr[carrier.len()..].trim_start();
    let Some(after) = after.strip_prefix(".ListRows") else {
        return None;
    };
    // The row expression: Add answers the table's next row; else the
    // caller's index carries relative to the range's head.
    let row_expr = match after.trim() {
        ".Add" | "Add" => format!("({head_row} + 1 + __IbisDataRows())"),
        other => {
            let ix = other.trim_start_matches('(').split(')').next().unwrap_or("1").trim().to_string();
            format!("({head_row} + {ix} - 1)")
        }
    };
    let _ = (head_col, cols);
    vars.insert(
        name.to_string(),
        Binding::Row { table: sheet.clone(), head_row, head_col, row_expr: row_expr.clone() },
    );
    notes.push(format!("Set {name} = ListRows {} : the row is {row_expr}", after.trim()));
    Some(false)
}

/// One module source, rewritten for the form session.
pub fn rewrite_module(
    _module: &str,
    source: &str,
    form: &UserFormDef,
    table: &Option<(String, String, usize, usize, usize)>,
    data_rows: Option<usize>,
    others: &[String],
    extra_fns: &[String],
) -> (String, Vec<String>) {
    // A `_` line continuation: elixcee's parser does not carry one, the
    // lines join first.
    let mut joined: Vec<String> = Vec::new();
    let mut carried: Option<String> = None;
    for raw in source.lines() {
        match &mut carried {
            Some(head) => {
                head.push(' ');
                head.push_str(raw.trim());
                if raw.trim().ends_with('_') {
                    continue;
                }
                let merged = head.trim().trim_end_matches('_').to_string();
                carried = None;
                if merged.ends_with('_') {
                    carried = Some(merged.trim_end_matches('_').to_string());
                } else {
                    joined.push(merged);
                }
            }
            None => {
                if raw.trim().ends_with('_') {
                    carried = Some(raw.trim().trim_end_matches('_').to_string());
                } else {
                    joined.push(raw.to_string());
                }
            }
        }
    }
    if let Some(tail) = carried {
        joined.push(tail);
    }
    let mut rewriter = ModuleRewrite {
        vars: Default::default(),
        form: form.clone(),
        notes: Vec::new(),
        module_names: others.to_vec(),
        dropped_vars: Default::default(),
    };
    let rewriter = &mut rewriter;
    let mut out: Vec<String> = Vec::new();
    let mut with_stack: Vec<String> = Vec::new();
    for raw in joined {
        let mut line = raw.trim().to_string();
        if keep_as_written(&line) {
            out.push(raw.to_string());
            continue;
        }
        // A declaration's dotted type (an object library's): a Variant
        // stand-in parses (the procedures here never read the field).
        if line.starts_with("Private Sub ") || line.starts_with("Public Sub ") || line.starts_with("Sub ") {
            if let Some(as_at) = line.find(" As ") {
                let after = line[as_at + 4..].to_string();
                if after.contains('.') && !after.starts_with("( ") {
                    let (before, dotted) = line.split_at(as_at + 4);
                    let dotted_only: String = dotted.chars().take_while(|c| *c != ')' && *c != ',').collect();
                    if dotted_only.contains('.') {
                        line = format!("{before}Variant{}", &dotted[dotted_only.len()..]);
                    }
                }
            }
        }
        // With: the members expand the context; the With line drops.
        if let Some(expr) = line.strip_prefix("With ") {
            let mut inner = expr.to_string();
            // The form-reference hop strips: `With x.Ctl` binds the ctl.
            if with_stack.is_empty() {
                // `With <tblvar>` (a row carrier) stays bound, as the
                // member writes go through its row expr.
            }
            with_stack.push(inner.clone());
            let _ = inner;
            rewriter.note(format!("With {expr}: the members carry it"));
            continue;
        }
        if line == "End With" {
            with_stack.pop();
            continue;
        }
        // The Set rules first (drops or keeps; the drop notes the run).
        if let Some(keep) = set_of(&line, &mut rewriter.vars, form, table, &mut rewriter.notes) {
            if !keep {
                out.push(format!("' {line} (the form session re-route)"));
            } else {
                out.push(line);
            }
            continue;
        }
        if let Some(keep) = set_row(&line, &mut rewriter.vars, &mut rewriter.notes) {
            if !keep {
                out.push(format!("' {line} (the form session re-route)"));
            } else {
                out.push(line);
            }
            continue;
        }
        // Dim lines: an unknown dotted type becomes a Variant.
        if line.starts_with("Dim ") {
            out.push(dim_of(&line));
            continue;
        }
        // Named-argument calls (`Type:=msoControlButton`), a Set from a
        // With member (`Set x = .Controls.Add(...)`), and skipped
        // arguments (`Find(v, , xlValues, ...)`) are host glue;
        // elixcee's grammar stops on all three. A `Range("… & …")`
        // computed address parses a literal only, and the members after
        // one (`.Resize(`) do not either.
        let range_computed = {
            let mut at = 0usize;
            let mut hit = false;
            while let Some(ix) = line[at..].find(".Range(\"") {
                let args_start = at + ix + ".Range(\"".len();
                if let Some(close) = line[args_start..].find(')') {
                    if line[args_start..args_start + close].contains('&')
                        || line[args_start + close..].contains(".Resize(")
                        || line[args_start + close..].contains(".Offset(")
                    {
                        hit = true;
                        break;
                    }
                }
                at = args_start;
            }
            hit
        };
        if line.contains(":=")
            || line.contains(", ,")
            || (line.starts_with("Set ") && line.contains(" = ."))
            || range_computed
        {
            if let Some(set_to) = line.strip_prefix("Set ").and_then(|r| r.split_once('=')) {
                rewriter.dropped_vars.insert(set_to.0.trim().to_lowercase());
            }
            out.push(format!("' {line} (host-object glue the session does not model)"));
            continue;
        }
        // A dropped Set's variable chains: the object the variable held
        // is not modelled (`HC.EntireRow.Delete`, `Baris = HC.Row`).
        let lowered = line.to_lowercase();
        let dropped = rewriter
            .dropped_vars
            .iter()
            .any(|v| lowered.starts_with(&format!("{v}.")) || lowered.contains(&format!(" {v}.")));
        if dropped {
            out.push(format!("' {line} (the variable's Set was not modelled; the line drops)"));
            continue;
        }
        // Application events (the UI's own command path carries the
        // event semantics anyway), and an Unload (the window stays).
        if line.starts_with("Application.") && line.contains('=') {
            out.push(format!("' {line}"));
            continue;
        }
        if line.starts_with("Application.") || line.eq_ignore_ascii_case("Unload Me") {
            out.push(format!("' {line} (the session carries it)"));
            continue;
        }
        // The member line's context: a stack With expands it first. A
        // context the session cannot model (host objects with
        // `Type:=` named args, CommandBars' children) keeps the line a
        // comment, so the module itself parses.
        let line = match (line.starts_with('.'), with_stack.last().cloned()) {
            (true, Some(ctx))
                if ctx.contains(":=") || ctx.contains("Controls(") || ctx.contains("CommandBars") =>
            {
                out.push(format!("' {ctx}{line} (a host-object With the session does not model)"));
                continue;
            }
            (true, Some(ctx)) => format!("{ctx}{line}"),
            _ => line,
        };
        let mut line_out = line.clone();
        for module in rewriter.module_names.clone() {
            let prefix = format!("Call {module}.");
            if line_out.starts_with(&prefix) {
                line_out = format!("Call {}", &line_out[prefix.len()..]);
                rewriter.note(format!("Call {module}.: the bare name runs"));
                break;
            }
            let bare = format!("{module}.");
            if let Some(pos) = line_out.find(&bare) {
                line_out = format!("{}{}", &line_out[..pos], &line_out[pos + bare.len()..]);
                rewriter.note(format!("{module}.: the bare name runs"));
                break;
            }
        }
        let (new_line, with_note) = rewrite_line_chain(&line_out, &mut rewriter.vars, form);
        if let Some(note) = with_note {
            rewriter.note(note);
        }
        out.push(new_line);
    }
    // elixcee's cross-module Function calls answer Empty; the count the
    // bound needs is a number the session knows (the plane's own read in
    // Rust), so the bound becomes a literal here.
    let mut body = out.join("\r\n");
    if let Some(count) = data_rows {
        // The bound gains rows on top of the count: a Create in the same
        // run appends a row the pre-run count cannot know; the readback
        // scan skips the empty rows.
        body = body
            .replace("__IbisDataRows() + 1", &format!("{} + 11", count))
            .replace("__IbisDataRows()", &count.to_string());
    }
    let body = add_zero_arg_parens(&body, extra_fns);
    let mut notes = std::mem::take(&mut rewriter.notes);
    (body, notes)
}

/// One line's rewrite, in the rewriter's known bindings: (source kept
/// text, notes). The patterns, most particular first:
/// - `x.Range.Rows.Count` → `__IbisDataRows()`,
/// - `x.Range.Cells(r, c)` → `Worksheets("<sheet>").Cells(…).Value`,
/// - `x.Range(n) = …` and `x.Delete` (the row carrier),
/// - a form reference's `x.<control>.<member>` chains,
/// - the control's members directly (`Control.Caption = …` & reads).
fn rewrite_line_chain(
    line: &str,
    vars: &mut HashMap<String, Binding>,
    form: &UserFormDef,
) -> (String, Option<String>) {
    let mut notes: Vec<String> = Vec::new();
    let mut out = line.to_string();
    // The WorksheetFunction members the shim answers (min/max; elixcee
    // has no Application object), then the leftover `Application.`
    // qualifiers inside expressions.
    while let Some(ix) = out.find("WorksheetFunction.Min(") {
        out.replace_range(ix..ix + "WorksheetFunction.Min(".len(), "__IbisMin(");
    }
    while let Some(ix) = out.find("WorksheetFunction.Max(") {
        out.replace_range(ix..ix + "WorksheetFunction.Max(".len(), "__IbisMax(");
    }
    while let Some(ix) = out.find("WorksheetFunction.") {
        let after = &out[ix + "WorksheetFunction.".len()..];
        let member: String = after.chars().take_while(|c| c.is_alphanumeric() || *c == '_').collect();
        out.replace_range(ix..ix + "WorksheetFunction.".len(), format!("__IbisWF{member}").as_str());
    }
    out = out.replace("Application.", "");
    // The form reference's chains strip: `x.<control>…` (the With's
    // member lines built the chain top before this pass).
    for (var, binding) in vars.iter().map(|(k, v)| (k.clone(), v.clone())) {
        if !matches!(binding, Binding::Form) {
            continue;
        }
        let prefix = format!("{var}.");
        while let Some(ix) = out.find(&prefix) {
            out.replace_range(ix..ix + prefix.len(), "");
            notes.push(format!("form reference {var}. stripped"));
        }
    }
    // The form's Show: the window is open already.
    let with_show = format!("{}.Show", form.name);
    if out.trim_end().ends_with(&with_show) {
        let rest = out.trim_end()[..out.trim_end().len() - with_show.len()].trim_end().to_string();
        out = if rest.is_empty() {
            format!("Call __IbisShow(\"{}\")", form.name)
        } else {
            format!("Call __IbisShow(\"{}\") '{rest}", form.name)
        };
        notes.push("Show: the window carries it".to_string());
        return (out, Some(notes.join("; ")));
    }
    // A list control's method lines (Clear/AddItem) right after the
    // strip; the base is the control's own slot on the plane.
    for (ix, control) in form.controls.iter().enumerate() {
        if !is_list_kind(&control.kind) {
            continue;
        }
        let base = list_base(ix + 2);
        let clear = format!("{}.Clear", control.name);
        if out.trim() == clear {
            return (format!("Call __IbisListClear({base})"), None);
        }
        let add = format!("{}.AddItem", control.name);
        let head = out.trim_start();
        if head.eq_ignore_ascii_case(&add) {
            return ("' (list item: the writes carry the row)".to_string(), None);
        }
        // `X.AddItem expr` appends the expression to the slot's next
        // free row (the writers that follow fill the row's columns).
        if let Some(rest) = head.strip_prefix(&add) {
            let item = rest.trim();
            if !item.is_empty() {
                let item = rewrite_sheet_expr(item, vars);
                let (item_out, _) = rewrite_reads(&item, form);
                return (
                    format!("Call __IbisListAddItem({base}, ({item_out}))"),
                    Some(format!("{}. AddItem", control.name)),
                );
            }
        }
    }
    // 1. The bindings, longest name first (the same head: `x` and `xs`).
// 1. The bindings, longest name first (the same head: `x` and `xs`).
    let mut names: Vec<(String, Binding)> = vars
        .iter()
        .map(|(k, v)| (k.clone(), v.clone()))
        .collect();
    names.sort_by(|a, b| b.0.len().cmp(&a.0.len()));
    for (var, binding) in &names {
        match binding {
            Binding::Sheet { name_expr, .. } => {
                // The sheet's chains re-qualify (`var.Cells(` and
                // `var.Range("` aim the sheet); `var.Name` answers
                // through rewrite_sheet_expr on the expression level,
                // not a per-line put.
                let prefix = format!("{var}.");
                let qualified = format!("Worksheets({name_expr}).");
                for head in [format!("{prefix}Cells("), format!("{prefix}Range(\"")] {
                    while let Some(ix) = out.find(&head) {
                        out.replace_range(ix..ix + prefix.len(), &qualified);
                    }
                }
                let name_member = format!("{prefix}Name");
                while let Some(ix) = out.find(&name_member) {
                    let literal = format!("\"{}\"", binding_sheet_name(binding));
                    out.replace_range(ix..ix + prefix.len() + "Name".len(), &literal);
                }
            }
            Binding::Table { sheet, head_row, head_col, .. } => {
                let rows_count = format!("{var}.Range.Rows.Count");
                if let Some(ix) = out.find(&rows_count) {
                    let repl = "(__IbisDataRows() + 1)".to_string();
                    out.replace_range(ix..ix + rows_count.len(), &repl);
                    notes.push(format!("{var}.Range.Rows.Count"));
                    continue;
                }
                // `var.Range.Cells(r, c)` reads/writes (the offset's own
                // head: the range's table row is r 1-based, so the
                // absolute is head + r - 1; columns likewise).
                let cells = format!("{var}.Range.Cells(");
                while let Some(ix) = out.find(&cells) {
                    let args_start = ix + cells.len();
                    let Some(close) = out[args_start..].find(')') else { break };
                    let args = out[args_start..args_start + close].to_string();
                    let (row, col) = split_two(&args);
                    let repl = format!(
                        "Worksheets(\"{sheet}\").Cells(({head_row} + {row} - 1), ({head_col} + {col} - 1))"
                    );
                    // Replace `Cells(args)`'s end: the span from ix to args_start + close.
                    let end = args_start + close + 1;
                    out.replace_range(ix..end, &repl);
                    notes.push(format!("{var}.Range.Cells"));
                }
            }
            Binding::Row { table: _, head_row: _, head_col, row_expr } => {
                let delete = format!("{var}.Delete");
                if out.trim() == delete {
                    notes.push(format!("{var}.Delete: one row comes off"));
                    return (format!("Call __IbisDataDeleteRow(({row_expr}))"), None);
                }
                // `var.Range(n) = rhs` (1-based within the row's range):
                // the shim's sheet write at the carried row, head + n - 1.
                let trimmed = out.trim();
                if let Some(after) = trimmed.strip_prefix(&format!("{var}.Range(")) {
                    let Some((n, rhs)) = after.split_once(')') else { continue };
                    let Some(rhs) = rhs.trim_start().strip_prefix('=') else { continue };
                    let lens_n = n.trim().to_string();
                    // The form-reference hop strips before the reads run.
                    let mut rhs = rhs.trim().to_string();
                    for (var, binding) in vars.iter().map(|(k, v)| (k.clone(), v.clone())) {
                        if !matches!(binding, Binding::Form) {
                            continue;
                        }
                        let prefix = format!("{var}.");
                        while let Some(ix) = rhs.find(&prefix) {
                            rhs.replace_range(ix..ix + prefix.len(), "");
                        }
                    }
                    let (rhs, _) = rewrite_reads(&rhs, form);
                    let col = *head_col as usize + parse_number(&lens_n).unwrap_or(1) - 1;
                    let repl = format!("Call __IbisDataPut(({row_expr}), {col}, ({rhs}))");
                    let head_note = format!("{var}.Range({lens_n})");
                    notes.push(head_note);
                    return (repl, None);
                }
            }
            Binding::Form => {}
            Binding::SheetIndexed { index_expr } => {
                // elixcee's Worksheets takes only a literal, so the
                // chains drop; the `Name` reads ride the shim's sheet
                // list.
                let prefix = format!("{var}.");
                let name_member = format!("{prefix}Name");
                while let Some(ix) = out.find(&name_member) {
                    let call = format!("__IbisSheetName({index_expr})");
                    out.replace_range(ix..ix + prefix.len() + "Name".len(), &call);
                    notes.push(format!("{prefix}Name reads the session's sheet list"));
                }
                if out.contains(&prefix) {
                    notes.push(format!("chains through {prefix} drop: the indexed sheet's cells are not modelled"));
                    return (format!("' {out} (the indexed sheet's cells are not modelled)"), None);
                }
            }
        }
    }
    // 2. The form reference's chains strip: `x.<control>…`.
    for (var, binding) in &names {
        if !matches!(binding, Binding::Form) {
            continue;
        }
        let prefix = format!("{var}.");
        while let Some(ix) = out.find(&prefix) {
            out.replace_range(ix..ix + prefix.len(), "");
            notes.push(format!("form reference {var}. stripped"));
        }
    }
    // 3. The control members: reads and assignments. The line's own
    // shape decides (the shim's put when the member head matches a
    // control and the tail is an assignment).
    let is_assignment = {
        let trimmed = out.trim();
        !trimmed.starts_with("If ")
            && !trimmed.starts_with("Select Case")
            && !trimmed.starts_with("Case ")
            && !trimmed.starts_with("Do ")
            && !trimmed.starts_with("While ")
            && !trimmed.starts_with("For ")
            && !trimmed.starts_with("'")
            && trimmed.contains('=')
    };
    if is_assignment {
        let (lhs, rhs) = out.split_once('=').expect("assignment");
        // `X.Member = …` or `X.List(a, b) …`.
        let lhs_trim = lhs.trim();
        let mut put: Option<(String, String)> = None;
        for (control_index, control) in form.controls.iter().enumerate() {
            let name = &control.name;
            if let Some(after) = lhs_trim.strip_prefix(name.as_str()).and_then(|r| r.strip_prefix('.')) {
                let member = after.trim().trim_end_matches('=').trim();
                let plane = prop_of(member);
                if let Some(plane) = plane {
                    put = Some((name.clone(), plane.to_string()));
                    break;
                }
                if is_list_kind(&control.kind) {
                    let base = list_base(control_index + 2);
                    if member.eq_ignore_ascii_case("columncount") {
                        // The plane carries up to five columns.
                        notes.push(format!("{name}.ColumnCount: the plane has five"));
                        return (format!("' {line} (ColumnCount)"), None);
                    }
                    if let Some(args) = member.strip_prefix("List(").map(|a| a.trim_end()) {
                        let args = args.strip_suffix(')').unwrap_or(args).trim();
                        let (row, col) = split_two(args);
                        let (rhs_out, _) = rewrite_reads(rhs.trim(), form);
                        return (
                            format!("Call __IbisListPut({base}, {row}, {col}, ({rhs_out}))"),
                            None,
                        );
                    }
                    // A designer binding whose value the plane does not
                    // carry (`RowSource = …`, list styles) drops with a
                    // note: the parses it keeps would stop the module.
                    const UNMODELLED: [&str; 6] = [
                        "RowSource", "ListStyle", "BoundColumn", "TextColumn", "MatchEntry",
                        "ColumnWidths",
                    ];
                    if let Some(hit) = UNMODELLED.iter().find(|m| member.eq_ignore_ascii_case(m)) {
                        notes.push(format!("{name}.{hit}: not modelled, the line drops"));
                        return (format!("' {line} (the {hit} is not modelled)"), None);
                    }
                    notes.push(format!("{name}.{member}: not modelled"));
                    break;
                }
                notes.push(format!("{name}.{member}: not modelled"));
                break;
            }
        }
        if let Some((name, plane)) = put {
            let (rhs_out, _) = rewrite_reads(rhs.trim(), form);
            return (format!("Call __IbisPut(\"{name}\", \"{plane}\", ({rhs_out}))"), None);
        }
    }
    // Reads: the known members (a control's props, a ListBox's
    // `ListIndex`/`ListCount`/`List(r,c)`) and a method line
    // (`X.Clear`, `X.AddItem`).
    let (reads_out, note) = rewrite_reads(out.as_str(), form);
    return (reads_out, note);
}

fn split_two(args: &str) -> (String, String) {
    let args = args.trim();
    // A top-level split on the first comma (nesting in the demo is zero).
    match args.split_once(',') {
        Some((a, b)) => (a.trim().to_string(), b.trim().to_string()),
        None => (args.to_string(), "0".to_string()),
    }
}

fn parse_number(text: &str) -> Option<usize> {
    text.parse::<usize>().ok()
}

fn prop_of(member: &str) -> Option<&'static str> {
    match member {
        "Text" | "Value" => Some("Value"),
        "Caption" => Some("Caption"),
        "Checked" | "Check" => Some("Check"),
        "ListIndex" => Some("ListIndex"),
        _ => None,
    }
}

/// A sheet binding's members inside one expression: `ws.Cells(`,
/// `ws.Range("` and (a quoted set's) `ws.Name` re-qualify before the
/// reads run.
fn rewrite_sheet_expr(expr: &str, vars: &HashMap<String, Binding>) -> String {
    let mut out = expr.to_string();
    for (var, binding) in vars.iter().map(|(k, v)| (k.clone(), v.clone())) {
        let prefix = format!("{var}.");
        match &binding {
            Binding::Sheet { name_expr, quoted } => {
                let qualified = format!("Worksheets({name_expr}).");
                for head in [format!("{prefix}Cells("), format!("{prefix}Range(\"")] {
                    while let Some(ix) = out.find(&head) {
                        out.replace_range(ix..ix + prefix.len(), &qualified);
                    }
                }
                let name_member = format!("{prefix}Name");
                while let Some(ix) = out.find(&name_member) {
                    let literal = if *quoted {
                        format!("\"{}\"", name_expr.trim_matches('"'))
                    } else {
                        format!("Worksheets({name_expr}).Name")
                    };
                    out.replace_range(ix..ix + prefix.len() + "Name".len(), &literal);
                }
            }
            Binding::SheetIndexed { index_expr } => {
                let name_member = format!("{prefix}Name");
                while let Some(ix) = out.find(&name_member) {
                    out.replace_range(
                        ix..ix + prefix.len() + "Name".len(),
                        &format!("__IbisSheetName({index_expr})"),
                    );
                }
                if out.contains(&prefix) {
                    out = "' (the indexed sheet's cells are not modelled)".to_string();
                    break;
                }
            }
            _ => {}
        }
    }
    out
}

/// Reads: `identifier.member` pairs whose member the plane carries and
/// whose identifier is one of the form's controls become the shim's
/// get. The ListBox members route to the list plane.
fn rewrite_reads(line: &str, form: &UserFormDef) -> (String, Option<String>) {
    let mut out = line.to_string();
    // The list kinds' read members first (ListCount and List(a, b)).
    for (ix, control) in form.controls.iter().enumerate() {
        if !is_list_kind(&control.kind) {
            continue;
        }
        let base = list_base(ix + 2);
        let count = format!("{}.ListCount", control.name);
        if let Some(at) = out.find(&count) {
            out.replace_range(at..at + count.len(), &format!("__IbisListRows({base})"));
            continue;
        }
        let list = format!("{}.List(", control.name);
        while let Some(at) = out.find(&list) {
            let args_start = at + list.len();
            let Some(close) = out[args_start..].find(')') else { break };
            let args = out[args_start..args_start + close].trim().to_string();
            let (row, col) = split_two(args.trim_end_matches(')').trim());
            let end = args_start + close + 1;
            out.replace_range(at..end, &format!("__IbisListGet({base}, {row}, {col})"));
            break;
        }
    }
    // The control member reads.
    let mut rest = out.as_str();
    let mut body = String::new();
    let mut found: Vec<String> = Vec::new();
    while let Some(dot_at) = rest.find('.') {
        let before = &rest[..dot_at];
        let after_dot = &rest[dot_at + 1..];
        let member: String = after_dot.chars().take_while(|c| c.is_alphanumeric() || *c == '_').collect();
        if member.is_empty() {
            body.push_str(&rest[..=dot_at]);
            rest = after_dot;
            continue;
        }
        let word_start = before
            .rfind(|c: char| !(c.is_alphanumeric() || c == '_'))
            .map(|i| i + 1)
            .unwrap_or(0);
        let name = &before[word_start..];
        if form.control(name).is_none() {
            body.push_str(&rest[..=dot_at]);
            rest = after_dot;
            continue;
        }
        match prop_of(&member) {
            Some(plane_prop) => {
                body.push_str(&before[..word_start]);
                body.push_str(&format!("__IbisGet(\"{name}\", \"{plane_prop}\")"));
                found.push(format!("{name}.{member}"));
                rest = &after_dot[member.len()..];
            }
            None => {
                body.push_str(&rest[..=dot_at]);
                rest = after_dot;
            }
        }
    }
    body.push_str(rest);
    let note = if found.is_empty() { None } else { Some(format!("read {}", found.join(", "))) };
    (body, note)
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
    // The count the For-literal rewrite answers: load the workbook and
    // count the table's data rows in Rust (deterministic; the run's own
    // bound then does not need a cross-module Function call).
    // #[cfg(test)]
    let table = read_table_range(&session.package);
    let data_rows: Option<usize> = (|| -> Option<usize> {
        let (_, _, head_row, _, _) = table.clone()?;
        let workbook = session.workbook.clone().unwrap_or_else(|| session.package.clone());
        let mut vm = Vm::new();
        vm.load_workbook_file(&workbook.display().to_string()).ok()?;
        // elixcee's sheet keys are lowercased (`sheet_names` answers
        // them); the table's own sheet carries the range's head cell.
        let candidates = vm.sheet_names();
        let mut found = None;
        for candidate in &candidates {
            // The plane sheet is skipped: its own control rows would
            // otherwise look like a table.
            if candidate.starts_with("__ibis") {
                continue;
            }
            vm.set_active_sheet(candidate).ok()?;
            let text = vm
                .cells()
                .get(&(head_row as u32, 1))
                .map(|c| variant_text(&c.value))
                .unwrap_or_default();
            if !text.is_empty() {
                found = Some(candidate.clone());
                break;
            }
        }
        let Some(found) = found else { return None };
        vm.set_active_sheet(&found).ok()?;
        // The count rides the table's first column below the header; a
        // missing row ends the scan.
        let cells = vm.cells();
        let mut count = 0usize;
        while count < 500 {
            let Some(cell) = cells.get(&((head_row + 1 + count) as u32, 1)) else {
                break;
            };
            let text = variant_text(&cell.value);
            if text.is_empty() {
                break;
            }
            count += 1;
        }
        Some(count)
    })();
    // One early load answers both the table's data-row count (the
    // For-bound's literal) and the sheet list (the enumerated
    // Worksheets(i).Name reads ride the shim's Case list).
    let table = read_table_range(&session.package);
    let (data_rows, sheets): (Option<usize>, Vec<String>) = (|| -> Option<(Option<usize>, Vec<String>)> {
        let workbook = session.workbook.clone().unwrap_or_else(|| session.package.clone());
        let mut vm = Vm::new();
        vm.load_workbook_file(&workbook.display().to_string()).ok()?;
        let sheets = vm.sheet_names();
        // elixcee's sheet keys are lowercased (`sheet_names` answers
        // them); the table's own sheet carries the range's head cell.
        let Some((_, _, head_row, _, _)) = &table else { return Some((None, sheets.clone())) };
        let candidates = vm.sheet_names();
        let mut found = None;
        for candidate in &candidates {
            // The plane sheet is skipped: its own control rows would
            // otherwise look like a table.
            if candidate.starts_with("__ibis") {
                continue;
            }
            vm.set_active_sheet(candidate).ok()?;
            let text = vm
                .cells()
                .get(&(*head_row as u32, 1))
                .map(|c| variant_text(&c.value))
                .unwrap_or_default();
            if !text.is_empty() {
                found = Some(candidate.clone());
                break;
            }
        }
        let Some(found) = found else { return Some((None, sheets.clone())) };
        vm.set_active_sheet(&found).ok()?;
        // The count rides the table's first column below the header; a
        // missing row ends the scan.
        let cells = vm.cells();
        let mut count = 0usize;
        while count < 500 {
            let Some(cell) = cells.get(&((head_row + 1 + count) as u32, 1)) else { break };
            let text = variant_text(&cell.value);
            if text.is_empty() {
                break;
            }
            count += 1;
        }
        Some((Some(count), sheets))
    })(
)
    .unwrap_or((None, Vec::new()));
    let modules = crate::read_modules(&session.package)?;
    let table = read_table_range(&session.package);
    let Some(code) = modules
        .iter()
        .find(|m| m.name.eq_ignore_ascii_case(&form.name))
        .map(|m| m.source.clone())
    else {
        return Err(format!("no module named {}", form.name));
    };
    let others: Vec<String> = modules
        .iter()
        .filter(|m| m.kind == "standard")
        .map(|m| m.name.clone())
        .collect();
    // Zero-arg functions the workbook's standard modules declare: the
    // form's own code may spell their calls bare.
    let std_fns: Vec<String> = modules
        .iter()
        .filter(|m| m.kind == "standard")
        .flat_map(|m| zero_arg_fns(&m.source))
        .collect();
    let (form_code, _notes) = rewrite_module(&form.name, &code, form, &table, data_rows, &others, &std_fns);
    let mut programs: Vec<(String, Program)> = Vec::with_capacity(modules.len() + 2);
    let shim = shim_source(&rows, &table, &sheets);
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
            .map_err(|e| format!("the form module does not compile: {e}\n--- rewritten:\n{form_code}\n---"))?,
    ));
    // The other standard modules: elixcee's cross-module Function
    // calls answer Empty, so the modules merge into one scope (a VBA
    // project's standard modules share the namespace anyway) — same
    // rewritings, one compile; a name clash between modules shows as
    // a compile error, reported whole.
    let mut merged = String::from("Attribute VB_Name = \"__IbisStd\"\r\nOption Explicit\r\n");
    for module in modules.iter().filter(|m| m.kind == "standard") {
        // The module's own name also strips (a self-qualified
        // `mCode.Read` inside mCode keeps the bare call).
        let own = vec![module.name.clone()];
        let (rewritten, _note) = rewrite_module(&module.name, &module.source, form, &table, data_rows, &own, &std_fns);
        for line in rewritten.lines() {
            if !line.starts_with("Attribute ") {
                merged.push_str(line);
                merged.push_str("\r\n");
            }
        }
    }
    #[cfg(test)]
    if std::env::var_os("IBIS_DUMP2").is_some() {
        let _ = std::fs::write(std::env::temp_dir().join("ibis-merged-std.txt"), &merged);
    }
    // The merged scope carries every standard module's functions, so
    // the zero-arg paren pass runs over the whole of it (a form
    // module's bare `GetBaris` then finds its declaration).
    let merged = add_zero_arg_parens(&merged, &[]);
    let standard = elixcee::parser::parse(&merged).map_err(|e| {
        format!(
            "the workbook's standard modules do not compile together: {e}\n--- source:\n{merged}\n---"
        )
    })?;
    programs.push(("__IbisStd".to_string(), standard));

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
        // classes carry the value only.
        let captioned = matches!(kind.as_str(), "Label" | "CommandButton");
        vm.set_cell_value(row, COL_TEXT, Variant::Str(value.text.clone()));
        if captioned {
            vm.set_cell_value(row, COL_CAPTION, Variant::Str(value.text.clone()));
        }
        vm.set_cell_value(row, COL_CHECKED, Variant::Boolean(value.checked));
        let _ = vm.set_cell_value(row, COL_LIST_INDEX, Variant::Integer(state.get(name).and_then(|v| v.selected).unwrap_or(0) as i64));
        // The list plane's rows come from the state (one cell per row),
        // cols 1..5, joined.
        if is_list_kind(&kind) {
            let base = list_base(row as usize) as u32;
            for (list_row, line) in value.list.iter().enumerate() {
                for (list_col, cell) in line.split('\t').enumerate() {
                    let _ = vm.set_cell_value(base + list_row as u32, LIST_COL_BASE + list_col as u32, Variant::Str(cell.to_string()));
                }
            }
        }
    }

    vm.run_sub_multi(&programs, event)
        .map_err(|e| format!("the event failed: {e}\n--- code:\n{form_code}\n---"))?;
    let mut messages = vm.take_messages();
    messages.extend(vm.take_debug_output());

    vm.set_active_sheet(CTL_SHEET)
        .map_err(|e| format!("the scratch sheet cannot reopen: {e}"))?;
    let cells = vm.cells();
    let mut out_state: HashMap<String, FormValue> = HashMap::new();
    for (name, kind, row) in &rows {
        let kind = kind.clone();
        let row = *row as u32;
        let mut value = state.get(name).cloned().unwrap_or_default();
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
        if is_list_kind(&kind) {
            // Gap-tolerant: the run's own Create appends rows past the
            // pre-run count, and the empty ListPut rows in between do
            // not stop the scan.
            let base = list_base(row as usize) as u32;
            let mut list: Vec<String> = Vec::new();
            for list_row in 0..LIST_MAX_ROWS as u32 {
                let mut line: Vec<String> = Vec::new();
                for list_col in 0..LIST_MAX_COLS as u32 {
                    let cell = cells
                        .get(&(base + list_row, LIST_COL_BASE + list_col))
                        .map(|c| variant_text(&c.value))
                        .unwrap_or_default();
                    if cell.is_empty() {
                        break;
                    }
                    line.push(cell);
                }
                if line.is_empty() {
                    continue;
                }
                list.push(line.join("\t"));
            }
            value.list = list;
            value.selected = cells
                .get(&(row, COL_LIST_INDEX))
                .map(|c| variant_text(&c.value).parse::<usize>().ok())
                .unwrap_or_else(|| state.get(name).and_then(|v| v.selected));
        }
        out_state.insert(name.clone(), value);
    }
    let file = if session.save_back {
        crate::save_back_to_package(&session.package, &vm)
            .map_err(|e| format!("cannot save over the original: {e}"))?;
        session.package.clone()
    } else {
        elixcee::save_workbook(&vm, &output.display().to_string())
            .map_err(|e| format!("cannot save the form run: {e}"))?;
        output.to_path_buf()
    };
    Ok(FormRun { state: out_state, file, messages })
}

/// The workbook's ListObject range: the table's own xml carries the
/// range; the sheet the table lives on comes from the worksheet's rels
/// (the shim writes the sheet's name, so a qualified sheet resolves).
fn read_table_range(package: &Path) -> Option<(String, String, usize, usize, usize)> {
    let Ok(file) = std::fs::File::open(package) else { return None };
    let Some(mut archive) = zip::ZipArchive::new(file).ok() else { return None };
    // The tables: (name, range).
    let names: Vec<String> = archive.file_names().map(|n| n.to_string()).collect();
    let mut tables: Vec<(String, String, String)> = Vec::new(); // (name, ref, xml path)
    for name in &names {
        if !name.starts_with("xl/tables/") {
            continue;
        }
        let Some(mut bytes) = archive.by_name(name).ok() else { continue };
        // zip's ZipFile implements std::io::Read.
        use std::io::Read as _;
        let mut buf: Vec<u8> = Vec::new();
        if bytes.read_to_end(&mut buf).is_err() {
            continue;
        }
        let xml = String::from_utf8_lossy(&buf).into_owned();
        let Some(table_name) = between(&xml, "name=\"", "\"") else { continue };
        let Some(range) = between(&xml, "ref=\"", "\"") else { continue };
        tables.push((table_name, range, name.clone()));
    }
    // The table file's name tells the worksheet's rels which sheet it
    // belongs to; the rels' sheet parent names the sheet file, whose
    // own sheet name sits in workbook.xml in order — matching the two
    // by the sheet file that carries the rels entry.
    // (Sheets' rels: xl/worksheets/_rels/sheetN.xml.rels → Target
    // "../tables/<file>".)
    let Ok(file2) = std::fs::File::open(package) else { return None };
    let Ok(mut archive2) = zip::ZipArchive::new(file2) else { return None };
    let mut sheet_of: HashMap<String, String> = HashMap::new(); // tables/<file> → sheetN
    for name in &names {
        if !name.starts_with("xl/worksheets/_rels/") {
            continue;
        }
        let Some(mut rels) = archive2.by_name(name).ok() else { continue };
        let mut buf: Vec<u8> = Vec::new();
        use std::io::Read as _;
        if rels.read_to_end(&mut buf).is_err() {
            continue;
        }
        let text = String::from_utf8_lossy(&buf).into_owned();
        let sheet_file = name
            .split('/')
            .nth(2)
            .and_then(|s| s.strip_suffix(".rels"))
            .unwrap_or("")
            .to_string();
        for table in &tables {
            if text.contains(format!("/{}", table.2.split('/').next_back().unwrap_or("")).as_str()) {
                sheet_of.insert(table.2.split('/').next_back().unwrap_or("").to_string(), sheet_file.clone());
            }
        }
    }
    // The sheet files' own names: workbook.xml lists them in order with
    // the same rels ids... matching the file names to the sheets'
    // arithmetic order is not exposed; the demo's is the first:
    let Ok(file3) = std::fs::File::open(package) else { return None };
    let Ok(mut archive3) = zip::ZipArchive::new(file3) else { return None };
    let Some(mut wb_xml) = archive3.by_name("xl/workbook.xml").ok() else { return None };
    let mut buf: Vec<u8> = Vec::new();
    use std::io::Read as _;
    if wb_xml.read_to_end(&mut buf).is_err() {
        return None;
    }
    let wb = String::from_utf8_lossy(&buf).into_owned();
    let mut sheet_names: Vec<String> = Vec::new();
    let mut pos = 0usize;
    while let Some(at) = wb[pos..].find("<sheet name=\"") {
        let from = pos + at + "<sheet name=\"".len();
        let after = &wb[from..];
        let width = after.find('"').unwrap_or(0);
        sheet_names.push(after[..width].to_string());
        pos = from + width;
    }
    let _ = archive2;
    for table in &tables {
        let file_name = table.2.split('/').next_back().unwrap_or("");
        let Some(sheet_file) = sheet_of.get(file_name) else { continue };
        // sheetN.xml → the workbook's own <sheet name> Nth? The
        // worksheets' rels' sheet file order is not the names' order;
        // the r:id mapping is: resolve through the rel ids of the
        // workbook (rId1..) — for the demo's shape, the sheet file that
        // carries the table's rels is the sheet the table lives on, and
        // its name is the sheet whose cells the range's head cell owns;
        // for now the first sheet name list is matched to the file's
        // ordinal.
        let ordinal: usize = sheet_file
            .trim_start_matches("sheet")
            .trim_end_matches(".xml")
            .parse()
            .unwrap_or(1);
        let sheet_name = sheet_names.get(ordinal.saturating_sub(1)).cloned().unwrap_or_default();
        // "A4:C7": the range's head.
        let head = ref_parse(&table.1);
        return Some((table.0.clone(), sheet_name, head.0, head.1, 5));
    }
    None
}

fn ref_parse(range: &str) -> (usize, usize) {
    let mut col = 0usize;
    let mut chars = range.char_indices();
    let mut row_start = 0usize;
    for (ix, c) in chars.by_ref() {
        if c.is_ascii_digit() {
            row_start = ix;
            break;
        }
        col = col * 26 + (c.to_ascii_uppercase() as usize - 'A' as usize + 1);
    }
    let row: usize = range[row_start..]
        .split(':')
        .next()
        .and_then(|row| row.parse().ok())
        .unwrap_or(1);
    (row, col)
}

/// The first worksheet's name of a package (a list-object table lives
/// on one sheet; the demo's is the workbook's first).
fn first_sheet_name(package: &Path) -> String {
    let Ok(file) = std::fs::File::open(package) else { return String::new() };
    let Some(mut archive) = zip::ZipArchive::new(file).ok() else { return String::new() };
    let Some(mut xml) = archive.by_name("xl/workbook.xml").ok() else { return String::new() };
    let mut buf = Vec::new();
    use std::io::Read as _;
    if xml.read_to_end(&mut buf).is_err() {
        return String::new();
    }
    let text = String::from_utf8_lossy(&buf).into_owned();
    let Some(start) = text.find("<sheet name=") else { return String::new() };
    let after = &text[start + 12..]; // `<sheet name=` is 12 chars
    let after = after.trim_start_matches(|c: char| c == '"' || c == ' ');
    after.find('"').map(|end| after[..end].to_string()).unwrap_or_default()
}

pub fn __read_table_range_debug(package: &Path) -> Option<(String, String, usize, usize, usize)> {
    read_table_range(package)
}

fn between<'a>(text: &'a str, open: &str, close: &str) -> Option<String> {
    let start = text.find(open)? + open.len();
    let rest = &text[start..];
    let end = rest.find(close)?;
    Some(rest[..end].to_string())
}
