// SPDX-License-Identifier: AGPL-3.0-only
use std::collections::HashMap;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let package = std::path::PathBuf::from(&args[1]);
    match ibis_vba::forms::read_userforms(&package) {
        Ok(forms) => {
            for form in &forms {
                println!("FORM {} ({}x{}) controls={} notes={:?}", form.name, form.width, form.height, form.controls.len(), form.notes);
                for c in &form.controls {
                    let extra = c.caption.clone().unwrap_or_default();
                    println!("  {} [{}] {:?}", c.name, c.kind, extra);
                }
            }
        }
        Err(e) => println!("forms err: {e}"),
    }
    for m in ibis_vba::read_modules(&package).unwrap_or_default() {
        println!("module kind={} name={}", m.kind, m.name);
    }
    if args.len() > 2 {
        // Run the entry event against the first form, like the window does.
        let forms = ibis_vba::forms::read_userforms(&package).unwrap_or_default();
        let Some(form) = forms.first().cloned() else { return };
        let event = args[2].clone();
        let state: HashMap<String, ibis_vba::forms::run::FormValue> = form
            .controls
            .iter()
            .map(|c| {
                (
                    c.name.clone(),
                    ibis_vba::forms::run::FormValue {
                        text: c.value.clone().unwrap_or_default(),
                        checked: false,
                        list: Vec::new(),
                        selected: None,
                    },
                )
            })
            .collect();
        let out = std::env::temp_dir().join("ibis-probe-run.xlsx");
        let session = ibis_vba::forms::run::FormSession { package: package.clone(), workbook: None, save_back: false };
        match ibis_vba::forms::run::run_form_event(&session, &form, &event, &state, &out) {
            Ok(run) => {
                println!("ran {event}");
                for (name, value) in &run.state {
                    if !value.list.is_empty() || !value.text.is_empty() || value.checked {
                        println!("  {name} text={:?} checked={} list={:?} sel={:?}",
                            value.text, value.checked, value.list, value.selected);
                    }
                }
                for msg in &run.messages {
                    println!("  msg: {msg}");
                }
            }
            Err(e) => println!("run {event} err: {e}"),
        }
    }
}
