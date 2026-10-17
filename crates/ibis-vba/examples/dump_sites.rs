// SPDX-License-Identifier: AGPL-3.0-only
// A look at a form's site table: name, depth, class, object-stream size.
use std::io::Read;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let file = std::fs::File::open(&args[1]).unwrap();
    let mut archive = zip::ZipArchive::new(file).unwrap();
    let mut raw = Vec::new();
    let mut found = String::new();
    for name in archive.file_names() {
        if name.ends_with("vbaProject.bin") {
            found = name.to_string();
        }
    }
    archive.by_name(&found).unwrap().read_to_end(&mut raw).unwrap();
    let project = ovba::open_project(raw).unwrap();
    let form = args.get(2).cloned().unwrap_or_default();
    for (name, path) in project.list().unwrap() {
        let short = path.split('/').next_back().unwrap_or("");
        let owner = path.split('/').nth(1).unwrap_or("");
        if (short == "f" || short == "o") && owner == form {
            let bytes = project.read_stream(&path).unwrap();
            println!("== {name:?} len {}", bytes.len());
            for (ix, b) in bytes.iter().take(96).enumerate() {
                print!("{:02x} ", b);
                if ix % 16 == 15 { println!(); }
            }
            println!();
        }
    }
}
