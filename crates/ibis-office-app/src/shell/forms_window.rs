// SPDX-License-Identifier: AGPL-3.0-only
// Copyright (c) ibis inc.
// Additional terms under AGPL-3.0 section 7(b): see NOTICE.
//! マクロ・フォームのウィンドウ (the form window): one UserForm of the
//! open file, its controls laid out where the designer stored them,
//! running the form module's event procedures (`ibis-vba`'s form
//! session; the value plane is the VM's, see `forms/run.rs`).
//!
//! `UserForm_Initialize` runs when the window opens (if the module
//! carries one), and a CommandButton click runs `<Control>_Click`.
//! The run is bounded like every run here, and nothing runs when the
//! file merely opens.
use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::Arc;

use gpui_kit::component::button::{Button, ButtonVariants as _};
use gpui_kit::component::checkbox::Checkbox;
use gpui_kit::component::input::{Input, InputState};
use gpui_kit::component::{ActiveTheme as _, Disableable as _, Root, Sizable as _, h_flex, v_flex};
use gpui_kit::prelude::FluentBuilder as _;
use gpui_kit::*;

use crate::shell::editor::SlideEditor;
use crate::shell::i18n;

thread_local! {
    /// The form window, while it is open (one at a time).
    static OPEN: std::cell::RefCell<Option<WindowHandle<Root>>> = const { std::cell::RefCell::new(None) };
}

pub(crate) struct FormWindow {
    focus: FocusHandle,
    /// The form (its designer data) and its file (events, workbook).
    form: Arc<ibis_vba::forms::UserFormDef>,
    package: PathBuf,
    /// The text inputs (a TextBox on the form → its edit state); the
    /// click reads them into the run's state.
    inputs: HashMap<String, Entity<InputState>>,
    /// The run state (the plane's last read), carried between runs.
    state: HashMap<String, ibis_vba::forms::FormValue>,
    /// The workbook the last run wrote (the plane across runs).
    workbook: Option<PathBuf>,
    /// Excel-style save: the run's result replaces the original file
    /// (its VBA project survives); off, the `-form-run.xlsx` copy takes
    /// it, the original untouched.
    save_back: bool,
    running: bool,
    status: Option<(bool, SharedString)>,
    /// The editor that opened the form (unused for the run itself; kept
    /// for a future reload of the sheet-based plane in the editor).
    _editor: Option<WeakEntity<SlideEditor>>,
}

/// Opens the form window for one form (the macros window's フォーム row).
pub(crate) fn open_form(
    editor: Option<WeakEntity<SlideEditor>>,
    parent: Option<Bounds<Pixels>>,
    display: Option<DisplayId>,
    package: PathBuf,
    form: ibis_vba::forms::UserFormDef,
    cx: &mut App,
) {
    if let Some(handle) = OPEN.with(|o| *o.borrow())
        && handle.update(cx, |_, window, _| window.activate_window()).is_ok()
    {
        return;
    }
    let width = px((form.width + 32.).max(320.));
    let height = px((form.height + 88.).max(240.));
    let bounds = match parent {
        Some(p) => Bounds::new(
            point(p.origin.x + (p.size.width - width) / 2.0, p.origin.y + (p.size.height - height) / 2.0),
            size(width, height),
        ),
        None => Bounds::centered(None, size(width, height), cx),
    };
    let title: SharedString = if form.caption.is_empty() {
        form.name.clone().into()
    } else {
        form.caption.clone().into()
    };
    let options = WindowOptions {
        window_bounds: Some(WindowBounds::Windowed(bounds)),
        display_id: display,
        titlebar: Some(TitlebarOptions { title: Some(title.clone()), ..Default::default() }),
        window_decorations: crate::server_decorations(),
        ..Default::default()
    };
    let _ = cx.open_window(options, |window, cx| {
        let view = cx.new(|cx| {
            let mut inputs = HashMap::new();
            for control in &form.controls {
                if control.kind == "TextBox" {
                    inputs.insert(
                        control.name.clone(),
                        cx.new(|cx| {
                            let mut state = InputState::new(window, cx).placeholder(String::new());
                            state.set_value(control.value.clone().unwrap_or_default(), window, cx);
                            state
                        }),
                    );
                }
            }
            FormWindow {
                focus: cx.focus_handle(),
                form: Arc::new(form.clone()),
                package: package.clone(),
                inputs,
                state: form
                    .controls
                    .iter()
                    .map(|c| {
                        (
                            c.name.clone(),
                            ibis_vba::forms::FormValue {
                                text: c.value.clone().unwrap_or_default(),
                                checked: false,
                                list: Vec::new(),
                                selected: None,
                            },
                        )
                    })
                    .collect(),
                workbook: None,
                save_back: false,
                running: false,
                status: None,
                _editor: editor.clone(),
            }
        });
        // UserForm_Initialize, when the module carries one.
        view.update(cx, |this, cx| this.perform_run("UserForm_Initialize", window, cx));
        let focus = view.read(cx).focus.clone();
        focus.focus(window, cx);
        cx.new(|cx| Root::new(view, window, cx))
    });
}

impl FormWindow {
    fn close(&self, window: &mut Window) {
        window.remove_window();
    }

    fn output_path(&self) -> PathBuf {
        self.package
            .parent()
            .map(|p| p.to_path_buf())
            .unwrap_or_default()
            .join(format!(
                "{}-form-run.xlsx",
                self.package
                    .file_stem()
                    .map(|s| s.to_string_lossy().into_owned())
                    .unwrap_or_else(|| "out".into())
            ))
    }

    /// Reads the TextBox inputs into the state and runs `event` in the
    /// background; the plane's read-back updates the state and the
    /// labels. An event the module does not have is quietly skipped
    /// (`UserForm_Initialize` on a form without one).
    fn perform_run(&mut self, event: &'static str, window: &mut Window, cx: &mut Context<Self>) {
        if self.running || self.form.controls.is_empty() {
            return;
        }
        let mut state = self.state.clone();
        for (control, input) in &self.inputs {
            if let Some(value) = state.get_mut(control) {
                value.text = input.read(cx).value().to_string();
            }
        }
        self.running = true;
        self.status = Some((false, i18n::text("macros.running")));
        cx.notify();
        let form = self.form.clone();
        let package = self.package.clone();
        let output = self.output_path();
        let workbook = self.workbook.clone();
        let event_marker = format!("'{event}' not found");
        let bounded = ibis_vba::forms::run::run_form_event(
            &ibis_vba::forms::run::FormSession { package, workbook, save_back: self.save_back },
            &form,
            event,
            &state,
            &output,
        );
        let finished = cx.background_executor().spawn(async move { bounded });
        cx.spawn_in(window, async move |this, cx| {
            let result = finished.await;
            let _ = this.update(cx, |this, cx| {
                this.running = false;
                match result {
                    Ok(run) => {
                        this.state = run.state.clone();
                        this.workbook = Some(run.file.clone());
                        this.status = Some((false, i18n::text("macros.done-simple").into()));
                    }
                    Err(err) if err.contains(&event_marker) => {
                        this.status = None;
                    }
                    Err(e) => this.status = Some((true, format!("{e}").into())),
                }
                cx.notify();
            });
        })
        .detach();
    }
}

impl Render for FormWindow {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let theme = cx.theme().clone();
        // The decorations: when the WM drew no bar (its fallback), the
        // window's own titlebar carries the drag and the buttons.
        let mut body = v_flex().size_full().gap_0().bg(theme.background).text_color(theme.foreground).track_focus(&self.focus);
        let mut panel = v_flex().flex_1().min_h_0().gap_2().p_4();
        if let Some(bar) = crate::shell::chrome::floating_titlebar(
            if self.form.caption.is_empty() {
                self.form.name.clone().into()
            } else {
                self.form.caption.clone().into()
            },
            window,
        ) {
            panel = panel.child(bar);
        }
        panel = panel.child(
            h_flex()
                .w_full()
                .gap_2()
                .items_center()
                .child(div().text_size(px(13.)).font_weight(FontWeight::SEMIBOLD).child(self.form.name.clone()))
                .child(div().text_size(px(11.)).text_color(theme.muted_foreground).child(self.form.caption.clone())),
        );
        // The controls, where the designer stored them (px at 96 DPI).
        let mut plane = div()
            .relative()
            .rounded_md()
            .border_1()
            .border_color(theme.border)
            .bg(theme.background)
            .w(px(self.form.width))
            .h(px(self.form.height))
            .overflow_hidden();
        for control in &self.form.controls {
            let base = div()
                .absolute()
                .left(px(control.left))
                .top(px(control.top))
                .w(px(control.width))
                .h(px(control.height));
            let kind = control.kind.as_str();
            let shown = || {
                self.state
                    .get(&control.name)
                    .filter(|v| !v.text.is_empty())
                    .map(|v| v.text.clone())
                    .or_else(|| control.caption.clone())
                    .unwrap_or_else(|| control.name.clone())
            };
            let element: AnyElement = match kind {
                "Label" => base.child(div().text_size(px(11.)).child(shown())).into_any_element(),
                "TextBox" => match self.inputs.get(&control.name) {
                    Some(input) => base.child(Input::new(input).small()).into_any_element(),
                    None => base.child(div().child(control.name.clone())).into_any_element(),
                },
                "CommandButton" => {
                    let name = control.name.clone();
                    base.child(
                        Button::new(name.clone())
                            .small()
                            .label(shown())
                            .disabled(self.running)
                            .on_click(cx.listener(move |this, _, window, cx| {
                                // One event name per click; leaked once —
                                // the same leak pattern's static.
                                let event: &'static str =
                                    Box::leak(format!("{name}_Click").into_boxed_str());
                                this.perform_run(event, window, cx);
                            })),
                    )
                        .into_any_element()
                }
                "ListBox" | "ComboBox" => {
                    // The plane's rows (the last run's list) and a
                    // double-click that runs the control's own
                    // `<Control>_DblClick`, when the module has one.
                    let name = control.name.clone();
                    let rows: Vec<SharedString> = self
                        .state
                        .get(&control.name)
                        .map(|v| {
                            v.list
                                .iter()
                                .map(|line| line.replace('\t', "  ").into())
                                .collect()
                        })
                        .unwrap_or_default();
                    let selected = self
                        .state
                        .get(&control.name)
                        .and_then(|v| v.selected)
                        .unwrap_or(usize::MAX);
                    let mut list = div()
                        .size_full()
                        .rounded_md()
                        .border_1()
                        .border_color(theme.border)
                        .bg(theme.background)
                        .text_size(px(11.))
                        .overflow_hidden();
                    for (ix, line) in rows.iter().enumerate() {
                        let row_name = control.name.clone();
                        let dbl_name = name.clone();
                        list = list.child(
                            div()
                                .id(SharedString::from(format!("{row_name}-row-{ix}")))
                                .px_1()
                                .when(ix == selected, |d| d.bg(theme.accent).text_color(theme.accent_foreground))
                                .hover(|d| d.bg(theme.secondary))
                                .on_click(cx.listener(
                                    move |this, event: &gpui_kit::ClickEvent, window, cx| {
                                        if event.click_count() != 2 {
                                            return;
                                        }
                                        let dbl: &'static str =
                                            Box::leak(format!("{dbl_name}_DblClick").into_boxed_str());
                                        this.perform_run(dbl, window, cx);
                                    },
                                ))
                                .child(line.clone()),
                        );
                    }
                    base.child(list).into_any_element()
                }
                other => base
                    .child(
                        div()
                            .size_full()
                            .rounded_md()
                            .bg(theme.secondary)
                            .text_color(theme.secondary_foreground)
                            .text_size(px(10.))
                            .child(format!("{} ({other})", control.name)),
                    )
                    .into_any_element(),
            };
            plane = plane.child(element);
        }
        panel = panel.child(
            v_flex().id("form-plane-scroll").flex_1().min_h_0().items_start().overflow_y_scroll().child(plane),
        );
        // The status (what ran, what failed) and the window's close.
        let (message, color) = match &self.status {
            Some((true, text)) => (text.clone(), rgb(0xb3261e).into()),
            Some((false, text)) => (text.clone(), rgb(0x52360c).into()),
            None => (
                SharedString::from(i18n::format("macros.note", &[(
                    "what",
                    &format!("{} {}", self.form.controls.len(), i18n::text("macros.controls")),
                )])),
                theme.muted_foreground,
            ),
        };
        panel = panel.child(
            h_flex()
                .w_full()
                .gap_2()
                .items_center()
                .child(div().text_size(px(12.)).text_color(color).child(message))
                .child(div().flex_1())
                .child(
                    Checkbox::new("form-save-back")
                        .label(i18n::text("forms.save-back"))
                        .checked(self.save_back)
                        .on_click(cx.listener(|this, checked: &bool, _, cx| {
                            this.save_back = *checked;
                            cx.notify();
                        })),
                )
                .child(
                    Button::new("form-close")
                        .small()
                        .ghost()
                        .label(i18n::text("_shell.cancel"))
                        .on_click(cx.listener(|this, _, window, _| this.close(window))),
                ),
        );
        body = body.child(panel);
    body
    }
}
