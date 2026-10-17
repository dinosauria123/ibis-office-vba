// SPDX-License-Identifier: AGPL-3.0-only
// Copyright (c) ibis inc.
// Additional terms under AGPL-3.0 section 7(b): see NOTICE.
//! ツール > マクロ… (the macros window): one dialog shaped like the
//! host's own — the macro names (including the file's UserForms), the
//! 実行 button (which runs a macro headlessly, or shows a form), the
//! 保存先 line, and the source the 編集 button unfolds.
//!
//! Safety: nothing runs until the user has pressed 実行. Every run is
//! bounded by ibis-vba's default instruction budget, and the VM refuses
//! external effects (CreateObject, shell and file access, Save/Close).
use std::path::PathBuf;
use std::sync::Arc;

use gpui_kit::component::button::{Button, ButtonVariants as _};
use gpui_kit::component::input::{Input, InputState};
use gpui_kit::component::searchable_list::SearchableListItem as _;
use gpui_kit::component::select::{Select, SelectEvent, SelectState};
use gpui_kit::component::{ActiveTheme as _, Disableable as _, IndexPath, Root, Sizable as _, h_flex, v_flex};
use gpui_kit::prelude::FluentBuilder as _;
use gpui_kit::*;

use crate::shell::editor::SlideEditor;
use crate::shell::i18n;

/// One name the dropdown and the list offer: a runnable macro
/// (`Module.Sub` entry point) or a UserForm (the 実行 of a form opens
/// its window).
#[derive(Clone)]
struct MacroEntryItem {
    name: String,
    is_form: bool,
    title: SharedString,
}

impl component::searchable_list::SearchableListItem for MacroEntryItem {
    type Value = String;
    fn title(&self) -> SharedString {
        self.title.clone()
    }
    fn value(&self) -> &String {
        &self.name
    }
}

/// The dropdown's data: the found names of the file.
struct MacroEntriesDelegate {
    items: Arc<Vec<MacroEntryItem>>,
}

impl component::searchable_list::SearchableListDelegate for MacroEntriesDelegate {
    type Item = MacroEntryItem;

    fn items_count(&self, _section: usize) -> usize {
        self.items.len()
    }
    fn item(&self, ix: component::IndexPath) -> Option<&Self::Item> {
        self.items.get(ix.row)
    }
    fn position<V>(&self, value: &V) -> Option<component::IndexPath>
    where
        Self::Item: component::searchable_list::SearchableListItem<Value = V>,
        V: PartialEq,
    {
        let row = self.items.iter().position(|item| *item.value() == *value)?;
        Some(component::IndexPath { section: 0, row, column: 0 })
    }
}

thread_local! {
    /// The macros window, while it is open (one at a time).
    static OPEN: std::cell::RefCell<Option<WindowHandle<Root>>> = const { std::cell::RefCell::new(None) };
}

pub(crate) struct MacrosWindow {
    focus: FocusHandle,
    /// The file's UserForms (designer data), for the 実行-that-shows path.
    forms: Arc<Vec<ibis_vba::forms::UserFormDef>>,
    /// The open file the modules were read from (None or unreadable: a note).
    path: Option<PathBuf>,
    /// The editor that opened the file: the result workbook opens in it
    /// (None: the app-menu path; the opening step hides).
    editor: Option<WeakEntity<SlideEditor>>,
    /// The window whose file the user came from (for opening the result in it).
    parent: Option<AnyWindowHandle>,
    /// The modules read at open (`Arc` so the loaded list can move into
    /// the run task even if the window closes meanwhile).
    modules: Arc<Vec<ibis_vba::VbaModule>>,
    /// What reading the project said when it failed.
    load_error: Option<SharedString>,
    /// The macro picker (the found names of the file): the type-to-search
    /// dropdown integrated in the entry box.
    select: Entity<SelectState<MacroEntriesDelegate>>,
    /// The picker's Confirm event, while the window lives.
    _select_subscription: gpui::Subscription,
    /// The names with their kinds (the list beside the picker shows it).
    entries: Arc<Vec<MacroEntryItem>>,
    /// The chosen entry (name, whether it is a form).
    chosen: Option<(String, bool)>,
    /// A run is going on (the buttons while disabled).
    running: bool,
    /// The result of the last run: whether it was an error, and its text.
    status: Option<(bool, SharedString)>,
    /// The workbook of the last run, for 実行結果を開く.
    result: Option<PathBuf>,
    /// The source the 編集 button unfolds (the chosen entry's module).
    show_source: bool,
}

/// Open the macros window for the open file's VBA project (the file
/// opened in the editor window the user is in; no file, a note).
pub(crate) fn open_over(
    path: Option<PathBuf>,
    editor: Option<WeakEntity<SlideEditor>>,
    parent_window: Option<AnyWindowHandle>,
    parent: Option<Bounds<Pixels>>,
    display: Option<DisplayId>,
    cx: &mut App,
) {
    if let Some(handle) = OPEN.with(|o| *o.borrow())
        && handle.update(cx, |_, window, _| window.activate_window()).is_ok()
    {
        return;
    }
    let modules = path.as_deref().map(ibis_vba::read_modules);
    let forms_loaded = path.as_deref().map(ibis_vba::forms::read_userforms);
    let (forms, forms_error): (Vec<ibis_vba::forms::UserFormDef>, Option<String>) = match forms_loaded {
        Some(Ok(forms)) => (forms, None),
        Some(Err(e)) => (Vec::new(), Some(e)),
        None => (Vec::new(), None),
    };
    let (modules, load_error) = match modules {
        Some(Ok(modules)) => (modules, None),
        Some(Err(e)) => (Vec::new(), Some(e.into())),
        None => (Vec::new(), Some(i18n::text("macros.open-a-file"))),
    };
    // The found names: every parameterless `Sub` of the standard and
    // document modules, then the file's forms (the 実行 of a form
    // opens its window).
    let macro_names: Vec<String> = modules
        .iter()
        .filter(|m| m.kind != "class")
        .flat_map(|m| subs_of(&m.name, &m.source))
        .collect();
    let mut it: Vec<MacroEntryItem> = macro_names
        .iter()
        .map(|name| MacroEntryItem { name: name.clone(), is_form: false, title: name.clone().into() })
        .collect();
    for form in &forms {
        let title: SharedString =
            i18n::format("macros.form-suffix", &[("name", &form.name)]).into();
        it.push(MacroEntryItem { name: form.name.clone(), is_form: true, title });
    }
    let entries = Arc::new(it);
    let (width, height) = (px(560.), px(420.));
    let bounds = match parent {
        Some(p) => Bounds::new(point(p.origin.x + (p.size.width - width) / 2.0, p.origin.y + (p.size.height - height) / 2.0), size(width, height)),
        None => Bounds::centered(None, size(width, height), cx),
    };
    let title = i18n::text("tools.macros");
    let options = WindowOptions {
        window_bounds: Some(WindowBounds::Windowed(bounds)),
        display_id: display,
        titlebar: Some(TitlebarOptions { title: Some(title.clone()), ..Default::default() }),
        window_decorations: crate::server_decorations(),
        ..Default::default()
    };
    let _ = cx.open_window(options, |window, cx| {
        // The picker's data: the found names, as items.
        let delegate = MacroEntriesDelegate { items: entries.clone() };
        let selected_index = (!entries.is_empty()).then(|| IndexPath { section: 0, row: 0, column: 0 });
        let select = cx.new(|cx| SelectState::new(delegate, selected_index, window, cx).searchable(true));
        let chosen = entries
            .first()
            .map(|e| (e.name.clone(), e.is_form));
        let view = cx.new(|cx| {
            // The picker's choice names the run.
            let subscription = cx.subscribe_in(
                &select,
                window,
                |this: &mut MacrosWindow,
                 _: &Entity<SelectState<MacroEntriesDelegate>>,
                 ev: &SelectEvent<MacroEntriesDelegate>,
                 _: &mut Window,
                 _: &mut Context<MacrosWindow>| {
                    if let SelectEvent::Confirm(Some(name)) = ev {
                        this.chosen = this
                            .entries
                            .iter()
                            .find(|e| &e.name == name)
                            .map(|e| (e.name.clone(), e.is_form));
                    }
                },
            );
            MacrosWindow {
                focus: cx.focus_handle(),
                path,
                editor,
                parent: parent_window,
                forms: Arc::new(forms),
                modules: Arc::new(modules),
                load_error: load_error.map(Into::into),
                select: select.clone(),
                _select_subscription: subscription,
                entries,
                chosen,
                running: false,
                status: None,
                result: None,
                show_source: false,
            }
        });
        let focus = view.read(cx).focus.clone();
        focus.focus(window, cx);
        cx.new(|cx| Root::new(view, window, cx))
    });
}

/// ツール > マクロ… (the app menu's, on macOS).
pub(crate) fn open(cx: &mut App) {
    let parent = cx.active_window().and_then(|w| w.update(cx, |_, window, cx| (window.bounds(), window.display(cx).map(|d| d.id()))).ok());
    open_over(None, None, None, parent.map(|p| p.0), parent.and_then(|p| p.1), cx);
}

/// Every `Sub` name in a module's source, as `Module.Sub`: the entries a
/// run can name (class modules' eventless Subs stay out: elixcee resolves
/// entry points across standard and document modules only).
fn subs_of(module_name: &str, source: &str) -> Vec<String> {
    let mut found = Vec::new();
    for line in source.lines() {
        let line = line.trim();
        for head in ["Public Sub ", "Private Sub ", "Friend Sub ", "Sub "] {
            if let Some(rest) = line.strip_prefix(head) {
                let name = rest.split(['(', ' ', '\t']).next().unwrap_or_default();
                // Subs with parameters cannot run as entry points (the
                // events; the caller supplies their arguments): stay out.
                // `BuildReport()` (nothing between the parentheses) can.
                let parametered = rest
                    .find('(')
                    .is_some_and(|open| rest[open + 1..].chars().take_while(|c| *c != ')').any(|c| !c.is_whitespace()));
                if !name.is_empty() && !parametered {
                    found.push(format!("{module_name}.{name}"));
                }
            }
        }
    }
    found
}

impl MacrosWindow {
    /// 実行結果を開く: the output workbook of the last run (a file beside
    /// the original) opens in the editor the user came from.
    fn open_result(&self, cx: &mut Context<Self>) {
        let (Some(result), Some(editor), Some(parent)) = (&self.result, self.editor.clone(), self.parent) else {
            return;
        };
        let result = result.clone();
        cx.defer(move |cx| {
            let _ = parent.update(cx, |_, window, cx| {
                let _ = editor.update(cx, |e, cx| e.open_path(result, None, window, cx));
            });
        });
    }

    /// The path the run saves its copy beside.
    fn output_path(&self) -> Option<PathBuf> {
        let path = self.path.as_ref()?;
        let dir = path.parent().map(|p| p.to_path_buf()).unwrap_or_default();
        let stem = path.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_else(|| "out".into());
        Some(dir.join(format!("{stem}-macro-run.xlsx")))
    }

    /// 実行: one button for both kinds — a form's entry opens its
    /// window (values plane, events), a macro runs on the disk file and
    /// opens the result.
    fn run_clicked(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        if self.load_error.is_some() || self.running {
            return;
        }
        let Some((entry, is_form)) = self.chosen.clone() else {
            self.status = Some((true, i18n::text("macros.enter-an-entry")));
            return cx.notify();
        };
        if is_form {
            let Some(package) = self.path.clone() else { return };
            let Some(form) = self.forms.iter().find(|f| f.name == entry).cloned() else {
                return;
            };
            self.running = true;
            cx.notify();
            crate::shell::forms_window::open_form(
                self.editor.clone(),
                Some(window.bounds()),
                window.display(cx).map(|d| d.id()),
                package,
                form,
                cx,
            );
            self.running = false;
            self.status = Some((false, i18n::text("macros.form-running")));
            return cx.notify();
        }
        let Some(path) = self.path.clone() else { return };
        let Some(out) = self.output_path() else { return };
        self.running = true;
        self.result = None;
        self.status = Some((false, i18n::text("macros.running")));
        cx.notify();
        let budget = ibis_vba::DEFAULT_MAX_VBA_INSTRUCTIONS;
        let bounded = ibis_vba::run_macro_bounded(&path, &entry, &out, budget);
        let finished = cx.background_executor().spawn(async move { bounded });
        cx.spawn_in(window, async move |this, cx| {
            let result = finished.await;
            let _ = this.update(cx, |this, cx| {
                this.running = false;
                this.result = result.as_ref().ok().map(|_| out.clone());
                this.status = Some(match result {
                    Ok(report) => {
                        let mut text = i18n::format("macros.done", &[("path", &report.saved)]);
                        if !report.skipped.is_empty() {
                            text += &format!(" ({})", i18n::format("macros.skipped", &[("count", &report.skipped.len().to_string())]));
                        }
                        (false, text.into())
                    }
                    Err(e) => (true, e.into()),
                });
                if this.result.is_some() {
                    // One click: the result workbook opens in the editor
                    // at once, unless the app-menu path (no editor).
                    this.open_result(cx);
                }
                cx.notify();
            });
        })
        .detach();
    }

    /// 編集: the chosen entry's module source folds out (a form's
    /// module's events too).
    fn source_of(&self) -> String {
        if let Some((name, is_form)) = &self.chosen {
            if *is_form {
                return i18n::format("macros.form-suffix", &[("name", name)]).to_string() + &format!("\n{}", i18n::text("macros.runs-on-disk"));
            }
            if let Some(module) = self.modules.iter().find(|m| {
                format!("{}.", m.name) == format!("{}.{}", name.split('.').next().unwrap_or(""), name.split('.').next().unwrap_or(""))
            }) {
                return module.source.clone();
            }
        }
        String::new()
    }
}

impl Render for MacrosWindow {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let theme = cx.theme().clone();
        let mut body = v_flex().size_full().gap_0().bg(theme.background).text_color(theme.foreground).track_focus(&self.focus);
        // The decorations: when the WM drew no bar (its fallback), the
        // window's own bar carries the drag and the buttons; the run's
        // panel then sits under it.
        if let Some(bar) = crate::shell::chrome::floating_titlebar(i18n::text("tools.macros"), window) {
            body = body.child(bar);
        }
        let mut panel = v_flex().flex_1().min_h_0().gap_2().p_4();
        
        // The names list: the picker (searchable) over the plain rows,
        // beside the 実行 column — the host's dialog layout.
        let is_form_chosen = self.chosen.as_ref().is_some_and(|(_, f)| *f);
        let mut names_col = v_flex().flex_1().min_w_0().min_h_0().gap_1()
            .child(div().text_size(px(12.)).child(i18n::text("macros.name")))
            .child(
                Select::new(&self.select)
                    .id("macros-entry")
                    .small()
                    .flex_shrink_0()
                    .disabled(self.running || self.load_error.is_some())
                    .placeholder(i18n::text("macros.entry-placeholder"))
                    .search_placeholder(i18n::text("macros.search-placeholder")),
            );
        let mut list = v_flex()
            .id("macros-names-list")
            .flex_1()
            .min_h_0()
            .mt_1()
            .rounded_md()
            .border_1()
            .border_color(theme.border)
            .bg(theme.background)
            .overflow_y_scroll();
        let entry_rows: Vec<(String, bool, SharedString)> = self
            .entries
            .iter()
            .map(|e| (e.name.clone(), e.is_form, e.title.clone()))
            .collect();
        for (clicked, is_form_row, shown) in entry_rows {
            let selected = self.chosen.as_ref().is_some_and(|(n, _)| n == &clicked);

            list = list.child(
                div()
                    .id(SharedString::from(format!("entry-{}", clicked)))
                    .cursor_pointer()
                    .px_2()
                    .py_1()
                    .text_size(px(12.))
                    .overflow_hidden()
                    .whitespace_nowrap()
                    .text_ellipsis()
                    .when(selected, |row| row.bg(theme.accent).text_color(theme.accent_foreground))
                    .hover(|d| d.bg(theme.secondary))
                    .child(shown)
                    .on_click(cx.listener(move |this, _, _, cx| {
                        this.chosen = Some((clicked.clone(), is_form_row));
                        this.show_source = false;
                        cx.notify();
                    })),
            );
        }
        names_col = names_col.child(list);

        // The right column: 実行, and the host-dialog's other actions
        // (the ones the session does not carry stay disabled).
        let mut actions = v_flex().w(px(120.)).flex_shrink_0().gap_1();
        let can_run = self.load_error.is_none() && self.chosen.is_some() && !self.running;
        actions = actions.child(
            Button::new("macros-run")
                .primary()
                .small()
                .label(i18n::text("macros.run"))
                .disabled(!can_run)
                .on_click(cx.listener(|this, _, window, cx| this.run_clicked(window, cx))),
        );
        actions = actions
            .child(Button::new("macros-step").small().ghost().label(i18n::text("macros.step-into")).disabled(true))
            .child(
                Button::new("macros-edit")
                    .small()
                    .ghost()
                    .label(i18n::text("macros.edit-code"))
                    .disabled(iis_nothing_for_edit(self.chosen.as_ref()) || self.running)
                    .on_click(cx.listener(|this, _, _, cx| {
                        this.show_source = !this.show_source;
                        cx.notify();
                    })),
            )
            .child(Button::new("macros-create").small().ghost().label(i18n::text("macros.create-new")).disabled(true))
            .child(Button::new("macros-delete").small().ghost().label(i18n::text("macros.delete-m")).disabled(true))
            .child(Button::new("macros-options").small().ghost().label(i18n::text("macros.options")).disabled(true));

        let mut top = h_flex().flex_1().min_h_0().gap_3().child(names_col).child(actions);

        // 編集: the module's source (the entry's module) folds out.
        if self.show_source {
            let source = self.source_of();
            top = top.child(
                v_flex()
                    .id("macros-source")
                    .w(px(280.))
                    .flex_shrink_0()
                    .min_h_0()
                    .overflow_y_scroll()
                    .rounded_md()
                    .p_2()
                    .bg(theme.secondary)
                    .text_color(theme.secondary_foreground)
                    .text_size(px(11.))
                    .font_family("monospace")
                    .child(source),
            );
        }
        panel = panel.child(top).child(
            v_flex().gap_1()
                // The run's folder — the copy is saved beside the file.
                .child(
                    h_flex().w_full().gap_2().items_center()
                        .child(div().text_size(px(12.)).flex_shrink_0().child(i18n::text("macros.r-location")))
                        .child(
                            div()
                                .flex_1()
                                .min_w_0()
                                .rounded_md()
                                .border_1()
                                .border_color(theme.border)
                                .px_2()
                                .py_1()
                                .text_size(px(11.))
                                .text_color(theme.muted_foreground)
                                .overflow_hidden()
                                .whitespace_nowrap()
                                .text_ellipsis()
                                .child(
                                    self.path
                                        .as_ref()
                                        .map(|p| p.display().to_string())
                                        .unwrap_or_else(|| i18n::text("macros.no-file").to_string()),
                                ),
                        ),
                )
                .child(div().text_size(px(11.)).text_color(theme.muted_foreground).child(i18n::text("macros.description")))
                .child(div().text_size(px(11.)).text_color(theme.muted_foreground).child(i18n::text("macros.runs-on-disk"))),
        );

        // The status and the window's bottom row.
        let mut bottom = h_flex().w_full().gap_2().items_center();
        if let Some((error, text)) = &self.status {
            let color = match error {
                false => theme.muted_foreground,
                true => rgb(0xb3261e).into(),
            };
            bottom = bottom.child(div().text_size(px(11.)).text_color(color).flex_1().min_w_0().child(text.clone()));
        } else if let Some(load_error) = &self.load_error {
            bottom = bottom.child(div().text_size(px(11.)).text_color(rgb(0xb3261e)).flex_1().child(load_error.clone()));
        } else {
            bottom = bottom.child(
                div()
                    .text_size(px(11.))
                    .text_color(rgb(0x52360c))
                    .flex_1()
                    .child(i18n::text("macros.warning")),
            );
        }
        bottom = bottom.child(
            Button::new("macros-cancel")
                .small()
                .ghost()
                .label(i18n::text("_shell.cancel"))
                .on_click(cx.listener(|this, _, window, _| window.remove_window())),
        );
        panel = panel.child(bottom);
        body = body.child(panel);
        body
    }
}

/// 編集's enablement: a module entry (a macro's `Module.Sub`) has
/// source; a form's 実行 opens its own window.
fn iis_nothing_for_edit(chosen: Option<&(String, bool)>) -> bool {
    matches!(chosen, None | Some((_, true)))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[::core::prelude::v1::test]
    fn subs_of_finds_parameterless_subs() {
        let source = "Attribute VB_Name = \"Module1\"\r\nOption Explicit\r\n\r\nSub BuildReport()\r\n    Cells(1, 1).Value = 1\r\nEnd Sub\r\n\r\nPrivate Sub Hidden()\r\nEnd Sub\r\n";
        assert_eq!(subs_of("Module1", source), vec!["Module1.BuildReport", "Module1.Hidden"]);
    }

    #[::core::prelude::v1::test]
    fn subs_of_skips_parametered_subs() {
        // The events: the caller supplies their arguments.
        let source = "Sub Worksheet_Change(ByVal Target As Range)\r\nEnd Sub\r\n";
        assert!(subs_of("Sheet1", source).is_empty());
    }

    #[::core::prelude::v1::test]
    fn subs_of_indented_and_odd_signatures() {
        // No-parenthesis Subs are legal VBA; elixcee will say if it cannot
        // run one. The odd spacing (`Y ( )`) is not a parameter list.
        assert_eq!(subs_of("M", "    Sub X()\nEnd Sub\nSub Y ( )\nEnd Sub\nSub NoParens\nEnd Sub"), vec!["M.X", "M.Y", "M.NoParens"]);
    }
}
