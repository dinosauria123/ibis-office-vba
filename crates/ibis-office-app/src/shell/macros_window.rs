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
use gpui_kit::component::menu::{DropdownMenu as _, PopupMenuItem};
use gpui_kit::component::{ActiveTheme as _, Disableable as _, Root, Sizable as _, h_flex, v_flex};
use gpui_kit::prelude::FluentBuilder as _;
use gpui_kit::*;

use crate::shell::editor::SlideEditor;
use crate::shell::i18n;

thread_local! {
    /// The macros window, while it is open (one at a time).
    static OPEN: std::cell::RefCell<Option<WindowHandle<Root>>> = const { std::cell::RefCell::new(None) };
}

pub(crate) struct MacrosWindow {
    focus: FocusHandle,
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
    /// The runnable macro names found in the file (standard and document
    /// modules' `Sub`s), for the マクロ名 dropdown.
    entries: Arc<Vec<String>>,
    /// What reading the project said when it failed.
    load_error: Option<SharedString>,
    /// The module shown (an index into `modules`).
    selected: usize,
    /// The name to run: `Module.Sub` (the entry input).
    entry: Entity<InputState>,
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
        let view = cx.new(|cx| MacrosWindow {
            focus: cx.focus_handle(),
            path,
            editor,
            parent: parent_window,
            modules: Arc::new(modules),
            entries: Arc::new(entries),
            load_error: load_error.map(Into::into),
            selected: 0,
            entry: cx.new(|cx| InputState::new(window, cx).placeholder(i18n::text("macros.entry-placeholder"))),
            running: false,
            status: None,
            result: None,
        });
        // The found macro names go into the entry at once: the first one.
        let first = view.read(cx).entries.first().cloned();
        if let Some(first) = first {
            view.update(cx, |this, cx| {
                this.entry.update(cx, |input, cx| input.set_value(first, window, cx));
            });
        }
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

/// The name of the first `Sub` in a module's source (the entry the
/// input fills with when the module is selected), as `Module.Sub`.
fn first_sub(module_name: &str, source: &str) -> Option<String> {
    subs_of(module_name, source).into_iter().next()
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
                let parametered = rest
                    .find('(')
                    .is_some_and(|open| rest[open..].split_once(')').map_or(true, |(args, _)| !args.trim().ends_with("ByRef") && !args.trim().is_empty()));
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
        let entry = self.entry.read(cx).value().trim().to_string();
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
                        .on_click(cx.listener(move |this, _, window, cx| {
                            this.selected = index;
                            if let Some(module) = this.modules.get(index)
                                && let Some(entry) = first_sub(&module.name, &module.source)
                            {
                                this.entry.update(cx, |input, cx| {
                                    input.set_value(entry, window, cx);
                                });
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

        // The entry point input (the found names in a dropdown), and one
        // Run button (runs, then opens).
        let mut row = h_flex().w_full().gap_2().items_center();
        row = row.child(div().text_size(px(12.)).flex_shrink_0().child(i18n::text("macros.run-label")));
        if self.entries.is_empty() {
            row = row.child(Input::new(&self.entry).small().flex_1().disabled(self.running));
        } else {
            row = row.child(Input::new(&self.entry).small().flex_1().disabled(self.running));
            // The found macro names; picking one fills the entry.
            let entries = self.entries.clone();
            let entry = self.entry.clone();
            row = row.child(Button::new("macros-entries").small().label(format!("{} ({}{}", i18n::text("macros.found"), entries.len(), ")")).disabled(self.running).dropdown_menu(move |menu, _, _| {
                entries.iter().fold(menu, |menu, name| {
                    let (name, entry) = (name.clone(), entry.clone());
                    menu.item(PopupMenuItem::new(name.clone()).on_click(move |_, window, cx| {
                        entry.update(cx, |input, cx| input.set_value(name.clone(), window, cx));
                        cx.refresh_windows();
                    }))
                })
            }));
        }
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
