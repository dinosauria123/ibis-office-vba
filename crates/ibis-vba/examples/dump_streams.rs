// SPDX-License-Identifier: AGPL-3.0-only
use std::io::Read;

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let file = std::fs::File::open(&args[1]).unwrap();
    let mut archive = zip::ZipArchive::new(file).unwrap();
    let mut raw = Vec::new();
    let vba_entry = {
        let mut found = String::new();
        for name in archive.file_names() {
            if name.ends_with("vbaProject.bin") {
                found = name.to_string();
            }
        }
        found
    };
    archive.by_name(&vba_entry).unwrap().read_to_end(&mut raw).unwrap();
    let project = ovba::open_project(raw).unwrap();
    for (name, p) in project.list().unwrap() {
        println!("LIST {name:?} => {p:?}");
    }
}
