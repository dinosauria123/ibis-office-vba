// SPDX-License-Identifier: AGPL-3.0-only
use std::io::Read;

fn main() {
    let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test/benchmark/vba-userform-demo.xlsm");
    let file = std::fs::File::open(&path).unwrap();
    let mut archive = zip::ZipArchive::new(file).unwrap();
    let mut entry = archive.by_name("xl/vbaProject.bin").unwrap();
    let mut raw = Vec::new();
    entry.read_to_end(&mut raw).unwrap();
    let project = ovba::open_project(raw).unwrap();
    for (name, p) in project.list().unwrap() {
        println!("LIST {name:?} => {p:?}");
    }
    for candidate in [
        "UserForm1\\\u{3}VBFrame",
        "UserForm1/\u{3}VBFrame",
        "UserForm1\u{3}VBFrame",
        "\\UserForm1\\\u{3}VBFrame",
        "/UserForm1/\u{3}VBFrame",
        "UserForm1\\f",
        "UserForm1/f",
        "f",
        "UserForm1\\o",
        "UserForm1/o",
        "o",
    ] {
        match project.read_stream(candidate) {
            Ok(bytes) => println!("STREAM {candidate:?} len {}", bytes.len()),
            Err(e) => println!("STREAM {candidate:?} ERR {e}"),
        }
    }
}
