// SPDX-License-Identifier: AGPL-3.0-only
// Copyright (c) ibis inc.
// Additional terms under AGPL-3.0 section 7(b): see NOTICE.
//! ツール > マクロ… (the macros window): the VBA modules of the open
//! file, their source, and a Run action for data-processing macros,
//! headlessly executed on a copy of the workbook by `ibis-vba`
//! (elatables' elixcee VM; ovba reads the project).
//!
//! Safety: nothing runs until the user has pressed 実行 and then
//! confirmed the warning banner. Every run is bounded by
//! ibis-vba's default instruction budget, and the VM refuses external
//! effects (CreateObject, shell and file access, Save/Close).
use std::path::PathBuf;
use std::rc::Rc;
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

/// One macro the dropdown offers (a `Module.Sub` entry point found in
/// the file); value = name.
#[derive(Clone)]
struct MacroEntryItem {
    name: String,
}

impl component::searchable_list::SearchableListItem for MacroEntryItem {
    type Value = String;
    fn title(&self) -> SharedString {
        self.name.clone().into()
    }
    fn value(&self) -> &String {
        &self.name
    }
}

/// The dropdown's data: the runnable macro names found in the file.
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
    /// The file's UserForms (designer data), for the フォーム button.
    forms: Arc<Vec<ibis_vba::forms::UserFormDef>>,
    /// What the forms read said when it failed.
    forms_error: Option<SharedString>,
    /// The open file the modules were read from (None or unreadable: a note).
    path: Option<PathBuf>,
    /// The editor that opened the file: the 実行結果を開く button reloads
    /// the output workbook in it (None: the app-menu path; the button hides).
    editor: Option<WeakEntity<SlideEditor>>,
    /// The window whose file the user came from (for opening the result in it).
    parent: Option<AnyWindowHandle>,
    /// The modules read at open (`Arc` so the loaded list can move into
    /// the run task even if the window closes meanwhile).
    modules: Arc<Vec<ibis_vba::VbaModule>>,
    /// What reading the project said when it failed.
    load_error: Option<SharedString>,
    /// The module shown (an index into `modules`).
    selected: usize,
    /// The macro picker (the found names of the file): the type-to-search
    /// dropdown integrated in the entry box.
    select: Entity<SelectState<MacroEntriesDelegate>>,
    /// The picker's Confirm event, while the window lives.
    _select_subscription: gpui::Subscription,
    /// The chosen name to run (`Module.Sub`), kept to hand to the run.
    chosen: Option<String>,
    /// A run is going on (the buttons while disabled).
    running: bool,
    /// The result of the last run: whether it was an error, and its text.
    status: Option<(bool, SharedString)>,
    /// The workbook of the last run, for 実行結果を開く.
    result: Option<PathBuf>,
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
    let (forms, forms_error) = match forms_loaded {
        Some(Ok(forms)) => (forms, None),
        Some(Err(e)) => (Vec::new(), Some(e.into())),
        None => (Vec::new(), None),
    };
    let (modules, load_error) = match modules {
        Some(Ok(modules)) => (modules, None),
        Some(Err(e)) => (Vec::new(), Some(e.into())),
        None => (Vec::new(), Some(i18n::text("macros.open-a-file"))),
    };
    // The runnable names, found in the file: every `Sub` of the standard
    // and document modules (`Module.Sub`), in the modules' order.
    let entries: Vec<String> = modules
        .iter()
        .filter(|m| m.kind != "class")
        .flat_map(|m| subs_of(&m.name, &m.source))
        .collect();
    let (width, height) = (px(640.), px(520.));
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
        let items = entries.iter().map(|name| MacroEntryItem { name: name.clone() }).collect::<Vec<_>>();
        let delegate = MacroEntriesDelegate { items: Arc::new(items) };
        let selected_index = (!entries.is_empty()).then(|| IndexPath { section: 0, row: 0, column: 0 });
        let select = cx.new(|cx| SelectState::new(delegate, selected_index, window, cx).searchable(true));
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
                        this.chosen = Some(name.clone());
                    }
                },
            );
            MacrosWindow {
                focus: cx.focus_handle(),
                path,
                editor,
                parent: parent_window,
                forms: Arc::new(forms),
                forms_error: forms_error,
                modules: Arc::new(modules),
                load_error: load_error.map(Into::into),
                selected: 0,
                select: select.clone(),
                _select_subscription: subscription,
                chosen: entries.first().cloned(),
                running: false,
                status: None,
                result: None,
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

    /// The file' the run reads (the workbook copy is ibis-vba's own).
    fn output_path(&self) -> Option<PathBuf> {
        let path = self.path.as_ref()?;
        let dir = path.parent().map(|p| p.to_path_buf()).unwrap_or_default();
        let stem = path.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_else(|| "out".into());
        Some(dir.join(format!("{stem}-macro-run.xlsx")))
    }

    /// 実行: one click — run on the disk file, then open the result in
    /// the editor (the banner's warning and the runs-on-disk note stand).
    fn run_clicked(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        if self.load_error.is_some() || self.modules.is_empty() || self.running {
            return;
        }
        let Some(path) = self.path.clone() else { return };
        let Some(out) = self.output_path() else { return };
        let entry = self.chosen.clone().unwrap_or_default();
        if entry.is_empty() {
            self.status = Some((true, i18n::text("macros.enter-an-entry")));
            return cx.notify();
        }
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
}

impl Render for MacrosWindow {
    fn render(&mut self, _window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let theme = cx.theme().clone();
        let mut body = v_flex().size_full().gap_2().p_4().bg(theme.background).text_color(theme.foreground).track_focus(&self.focus);
        // The file the macros are read from.
        let file: SharedString = self.path.as_ref().map(|p| p.display().to_string().into()).unwrap_or_else(|| i18n::text("macros.no-file"));
        body = body.child(
            v_flex()
                .gap_1()
                .child(div().text_size(px(16.)).font_weight(FontWeight::SEMIBOLD).child(i18n::text("tools.macros")))
                .child(div().text_size(px(12.)).text_color(theme.muted_foreground).child(file.clone()))
                .child(div().text_size(px(11.)).text_color(theme.muted_foreground).child(i18n::text("macros.runs-on-disk"))),
        );
        // The forms read (when the file has any) and the button to show one.
        if let Some(first) = self.forms.first() {
            let package = self.path.clone();
            let editor = self.editor.clone();
            let form = first.clone();
            body = body.child(
                h_flex().w_full().gap_2().items_center()
                    .child(div().text_size(px(12.)).flex_1().child(
                        format!("{} ({})", i18n::text("macros.forms"), self.forms.len()),
                    ))
                    .when(self.forms_error.is_none(), |row| {
                        row.child(
                            Button::new("macros-open-form")
                                .small()
                                .label(i18n::text("macros.form-open"))
                                .on_click(move |_, window, cx| {
                                    let form = form.clone();
                                    crate::shell::forms_window::open_form(
                                        editor.clone(),
                                        Some(window.bounds()),
                                        window.display(cx).map(|d| d.id()),
                                        package.clone().unwrap_or_default(),
                                        form,
                                        cx,
                                    );
                                }),
                        )
                    }),
            );
        }
        // Reading the project failed (no file, or it is one elixcee cannot read).
        if let Some(error) = &self.load_error {
            body = body.child(div().text_size(px(12.)).text_color(rgb(0xb3261e)).child(error.clone()));
        }
        // The warning banner: informational, one click runs and opens.
        let mut banner = h_flex().w_full().gap_2().p_2().rounded_md().items_center().bg(rgb(0xfde68a)).text_color(rgb(0x52360c));
        banner = banner.child(div().text_size(px(12.)).flex_1().child(i18n::text("macros.warning")));
        body = body.child(banner);

        if !self.modules.is_empty() {
            // The module list (name and kind) beside the shown source.
            let mut list = v_flex()
                .id("macros-module-list")
                .w(px(200.))
                .flex_shrink_0()
                .gap_1()
                .overflow_y_scroll()
                .child(div().text_size(px(12.)).text_color(theme.muted_foreground).child(i18n::text("macros.modules")));
            for (index, module) in self.modules.iter().enumerate() {
                let selected = index == self.selected;
                list = list.child(
                    div()
                        .id(index)
                        .cursor_pointer()
                        .px_2()
                        .py_1()
                        .rounded_md()
                        .text_size(px(12.))
                        .overflow_hidden()
                        .whitespace_nowrap()
                        .text_ellipsis()
                        .when(selected, |row| row.bg(rgb(0xe7ddff)))
                        .text_color(if selected { theme.foreground } else { theme.muted_foreground })
                        .child(format!("{} ({})", module.name, module.kind))
                        .on_click(cx.listener(move |this, _, _, cx| {
                            this.selected = index;
                            // The picker follows: its first Sub of this module.
                            if let Some(module) = this.modules.get(index) {
                                if let Some(entry) = subs_of(&module.name, &module.source).into_iter().next() {
                                    this.chosen = Some(entry);
                                }
                            }
                            cx.notify();
                        })),
                );
            }
            let module = self.modules.get(self.selected);
            let source = module.map(|m| m.source.clone()).unwrap_or_default();
            let preview = v_flex().flex_1().min_w_0().gap_1()
                .child(div().text_size(px(12.)).text_color(theme.muted_foreground).child(i18n::text("macros.source")))
                .child(
                    v_flex()
                        .id("macros-source")
                        .flex_1()
                        .min_h_0()
                        .overflow_y_scroll()
                        .rounded_md()
                        .p_2()
                        .bg(theme.secondary).text_color(theme.secondary_foreground)
                        .text_size(px(11.))
                        .font_family("monospace")
                        .child(source),
                );
            body = body.child(h_flex().flex_1().min_h_0().gap_3().child(list).child(preview));
        }

        // The entry point: one box, the macro picker (type-to-search
        // dropdown of the found names), and one Run button.
        let mut row = h_flex().w_full().gap_2().items_center();
        row = row.child(div().text_size(px(12.)).flex_shrink_0().child(i18n::text("macros.run-label")));
        row = row.child(
            Select::new(&self.select)
                .id("macros-entry")
                .small()
                .flex_1()
                .disabled(self.running)
                .placeholder(i18n::text("macros.entry-placeholder"))
                .search_placeholder(i18n::text("macros.search-placeholder")),
        );
        if self.running {
            body = body.child(h_flex().w_full().gap_2().items_center().child(div().flex_1()).child(div().text_size(px(12.)).text_color(theme.muted_foreground).child(i18n::text("macros.running"))));
        } else {
            let cancel_run = || Button::new("macros-window-cancel").small().ghost().label(i18n::text("_shell.cancel")).on_click(cx.listener(|this, _, window, _| window.remove_window()));
            let run = Button::new("macros-run").primary().small().label(i18n::text("macros.run-and-open")).disabled(self.modules.is_empty() || self.load_error.is_some()).on_click(cx.listener(|this, _, window, cx| this.run_clicked(window, cx)));
            body = body.child(h_flex().w_full().gap_2().items_center().child(row).child(run).child(cancel_run()));
        }
        if let Some((error, text)) = &self.status {
            let color = match error {
                false => theme.muted_foreground,
                true => rgb(0xb3261e).into(),
            };
            body = body.child(div().text_size(px(12.)).text_color(color).child(text.clone()));
        }
        body
    }
}

#[allow(unused)]
fn unused(_: Rc<()>) {}

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
