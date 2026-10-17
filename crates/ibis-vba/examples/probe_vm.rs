// SPDX-License-Identifier: AGPL-3.0-only
fn main() {
    let package = std::path::PathBuf::from("/tmp/opencode/demo-gr2024.xlsm");
    let mut vm = elixcee::vm::Vm::new();
    vm.load_workbook_file(&package.display().to_string()).unwrap();
    vm.ensure_sheet("__probe");
    vm.set_active_sheet("__probe").ok();
    let args: Vec<String> = std::env::args().collect();
    let src = std::fs::read_to_string(&args[1]).unwrap();
    let program = match elixcee::parser::parse(&src) {
        Ok(p) => p,
        Err(e) => {
            println!("parse: {e}");
            return;
        }
    };
    match vm.run_sub(&program, "T") {
        Ok(()) => println!("ran ok"),
        Err(e) => println!("err: {e}"),
    }
    vm.set_active_sheet("__probe").ok();
    let cells = vm.cells();
    for r in 1u32..4 {
        println!("cell(r{r}) = {:?}", cells.get(&(r, 1)).map(|cc| format!("{:?}", cc.value)));
    }
}
