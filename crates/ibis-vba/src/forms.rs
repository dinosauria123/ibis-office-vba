// SPDX-License-Identifier: AGPL-3.0-only
// Copyright (c) ibis inc.
// Additional terms under AGPL-3.0 section 7(b): see NOTICE.
//! UserForm support: the designer streams inside `vbaProject.bin`
//! ([MS-OFORMS]) give a form's controls, and a form session runs the
//! form module's events against those controls.
//!
//! Table shapes follow [MS-OFORMS] the way a reference parser carries
//! them (pyopenvba's `forms.py` was the working example). Geometry is
//! HIMETRIC, converted at 96 DPI. Parsing is deliberately narrow —
//! Label, TextBox and CommandButton — and a control whose class is not
//! modelled is reported in the form's `notes` instead of failing it.
use std::io::Read;

pub mod run;
pub use run::FormValue;

// Streams inside a form's storage ([MS-OFORMS] 2.4).
const STREAM_VBFRAME: &str = "\u{0003}VBFrame";

// ClsidCacheIndex values: which class a control is ([MS-OFORMS] 2.4.5).
const CLS_COMMAND_BUTTON: u16 = 17;
const CLS_LABEL: u16 = 21;
// MorphData covers the text-like classes; TextBox is index 23 and the
// others (ListBox/ComboBox/…) follow their own DisplayStyle.
const CLS_MORPH_DATA: u16 = 15;
const CLS_TEXTBOX: u16 = 23;
const CLS_LISTBOX: u16 = 24;
const CLS_COMBOBOX: u16 = 25;
const CLS_CHECKBOX: u16 = 26;
const CLS_OPTIONBUTTON: u16 = 27;
const CLS_TOGGLEBUTTON: u16 = 28;
// Site records whose class is at or above this name an ActiveX class
// from the form's own class table, which the reader does not model.
const CLASS_TABLE_BASE: u16 = 0x8000;

fn is_container_class(index: u16) -> bool {
    matches!(index, 7 | 14 | 57) // Form, Frame, MultiPage
}

fn kind_of(index: u16) -> String {
    match index {
        7 => "Form".into(),
        12 => "Image".into(),
        14 => "Frame".into(),
        16 => "SpinButton".into(),
        CLS_COMMAND_BUTTON => "CommandButton".into(),
        18 => "TabStrip".into(),
        CLS_LABEL => "Label".into(),
        CLS_MORPH_DATA => "Morph".into(),
        CLS_TEXTBOX => "TextBox".into(),
        CLS_LISTBOX => "ListBox".into(),
        CLS_COMBOBOX => "ComboBox".into(),
        CLS_CHECKBOX => "CheckBox".into(),
        CLS_OPTIONBUTTON => "OptionButton".into(),
        CLS_TOGGLEBUTTON => "ToggleButton".into(),
        47 => "ScrollBar".into(),
        57 => "MultiPage".into(),
        CLASS_TABLE_BASE..=u16::MAX => "ActiveX".into(),
        _ => "Control".into(),
    }
}

// MARK: model

/// One control as the designer streams define it, in UI pixels
/// (HIMETRIC converted at 96 DPI).
#[derive(Debug, Clone)]
pub struct ControlDef {
    pub name: String,
    pub kind: String,
    pub left: f32,
    pub top: f32,
    pub width: f32,
    pub height: f32,
    /// The `Caption` the designer stored, if any.
    pub caption: Option<String>,
    /// The `Value`/`Text` (a TextBox's content), if stored.
    pub value: Option<String>,
}

/// One form as its designer streams define it.
#[derive(Debug, Clone)]
pub struct UserFormDef {
    pub name: String,
    /// The form's `Caption`, from the `\x03VBFrame` designer text.
    pub caption: String,
    pub width: f32,
    pub height: f32,
    pub controls: Vec<ControlDef>,
    /// What the reader could not model, reported rather than fatal.
    pub notes: Vec<String>,
}

impl UserFormDef {
    pub fn control(&self, name: &str) -> Option<&ControlDef> {
        self.controls.iter().find(|c| c.name == name)
    }
}

// MARK: reader

/// Align `pos` forward to `size` relative to `base` ([MS-OFORMS] 2.1.1.2.4).
fn align(pos: usize, base: usize, size: usize) -> usize {
    let over = (pos - base) % size;
    if over == 0 { pos } else { pos + size - over }
}

struct Reader<'a> {
    data: &'a [u8],
    pos: usize,
}

impl<'a> Reader<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, pos: 0 }
    }

    fn need(&self, count: usize) -> Result<(), String> {
        if self.pos + count > self.data.len() {
            return Err("designer stream walked out of bounds".into());
        }
        Ok(())
    }

    fn take(&mut self, count: usize) -> Result<&'a [u8], String> {
        self.need(count)?;
        let out = &self.data[self.pos..self.pos + count];
        self.pos += count;
        Ok(out)
    }

    fn u8(&mut self) -> Result<u8, String> {
        Ok(self.take(1)?[0])
    }

    fn u16(&mut self) -> Result<u16, String> {
        let b = self.take(2)?;
        Ok(u16::from_le_bytes([b[0], b[1]]))
    }

    fn u32(&mut self) -> Result<u32, String> {
        let b = self.take(4)?;
        Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    }

    fn i32(&mut self) -> Result<i32, String> {
        Ok(self.u32()? as i32)
    }

    fn skip(&mut self, count: usize) -> Result<(), String> {
        self.need(count)?;
        self.pos += count;
        Ok(())
    }

    fn align(&mut self, base: usize, size: usize) {
        self.pos = align(self.pos, base, size);
    }
}

fn utf16le(raw: &[u8]) -> String {
    let pairs = raw.chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]]));
    String::from_utf16_lossy(&pairs.collect::<Vec<_>>())
}

/// A stored string: mask bit set means compressed (cp1252, one byte per
/// char) in [MS-OFORMS] 2.2.10.12.3's convention, utf-16-le otherwise.
fn stored_string(raw: &[u8], compressed: bool) -> String {
    if compressed {
        raw.iter().map(|b| *b as char).collect()
    } else {
        utf16le(raw)
    }
}

fn himetric_px(h: i32) -> f32 {
    h as f32 / 2540.0 * 96.0
}

// MARK: OleSiteConcreteControl ([MS-OFORMS] 2.2.10.12.3)

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum SiteField {
    NameData,
    TagData,
    Id,
    HelpContextId,
    BitFlags,
    ObjectStreamSize,
    TabIndex,
    ClsidCacheIndex,
    GroupId,
    ControlTipTextData,
    RuntimeLicKeyData,
    ControlSourceData,
    RowSourceData,
}

/// (mask bit, field, DataBlock width). Bit 8 is the position: no
/// DataBlock field, the (left, top) pair lives in the extra block.
const SITE_FIELDS: &[(u32, SiteField, usize)] = &[
    (0, SiteField::NameData, 4),
    (1, SiteField::TagData, 4),
    (2, SiteField::Id, 4),
    (3, SiteField::HelpContextId, 4),
    (4, SiteField::BitFlags, 4),
    (5, SiteField::ObjectStreamSize, 4),
    (6, SiteField::TabIndex, 2),
    (7, SiteField::ClsidCacheIndex, 2),
    (9, SiteField::GroupId, 2),
    (11, SiteField::ControlTipTextData, 4),
    (12, SiteField::RuntimeLicKeyData, 4),
    (13, SiteField::ControlSourceData, 4),
    (14, SiteField::RowSourceData, 4),
];

#[derive(Debug, Default)]
struct Site {
    name: String,
    id: u32,
    clsid_cache_index: u16,
    object_stream_size: usize,
    /// fmPosition (left, top), HIMETRIC, when mask bit 8 is set.
    position: Option<(i32, i32)>,
}

/// One site record; the next begins at `start + 4 + cbSite`, the
/// record's length counting from its mask. All strings (Name and Tag,
/// then the rest) read per the record's own DataBlock length fields,
/// each aligned to 4 after it.
fn read_site(data: &[u8], r: &mut Reader) -> Result<Site, String> {
    let start = r.pos;
    if r.u16()? != 0 {
        return Err("unsupported OleSiteConcreteControl version".into());
    }
    let cb_site = r.u16()? as usize;
    let mask = r.u32()?;
    let mut name_len = 0usize;
    let mut name_compressed = false;
    let mut tag_len = 0usize;
    let mut other_lens: [(usize, usize); 4] = [(0, 0); 4];
    let mut lens_n = 0usize;
    let mut site = Site::default();
    for (bit, field, size) in SITE_FIELDS {
        if mask & (1 << bit) == 0 {
            continue;
        }
        r.align(start, *size);
        let value = match size {
            2 => r.u16()? as u32,
            _ => r.u32()?,
        };
        match field {
            SiteField::NameData => {
                name_len = (value & 0x7FFF_FFFF) as usize;
                name_compressed = value & 0x8000_0000 != 0;
            }
            SiteField::TagData => tag_len = (value & 0x7FFF_FFFF) as usize,
            SiteField::Id => site.id = value,
            SiteField::ObjectStreamSize => site.object_stream_size = value as usize,
            SiteField::ClsidCacheIndex => site.clsid_cache_index = value as u16,
            SiteField::ControlTipTextData
            | SiteField::RuntimeLicKeyData
            | SiteField::ControlSourceData
            | SiteField::RowSourceData => {
                if lens_n < other_lens.len() {
                    other_lens[lens_n] = (*field as usize, (value & 0x7FFF_FFFF) as usize);
                    lens_n += 1;
                }
            }
            _ => {}
        }
    }
    r.align(start, 4);
    // The extra block opens with the Name, unpadded before it.
    if mask & (1 << 0) != 0 {
        let raw = r.take(name_len)?;
        site.name = stored_string(raw, name_compressed);
        r.align(start, 4);
    }
    if mask & (1 << 1) != 0 {
        let raw = r.take(tag_len)?;
        let _tag = raw;
        r.align(start, 4);
    }
    // The position, when mask bit 8: (left, top) in HIMETRIC.
    if mask & (1 << 8) != 0 {
        site.position = Some((r.i32()?, r.i32()?));
    }
    for (_field, len) in &other_lens[..lens_n] {
        r.skip(*len)?;
        r.align(start, 4);
    }
    let end = start + 4 + cb_site;
    if r.pos != end {
        return Err(format!("site at {start} has {} unmodelled bytes", end - r.pos));
    }
    Ok(site)
}

// MARK: FormSiteData ([MS-OFORMS] 2.2.10.12)

/// The form's site table: the class table (its count depends on a flag
/// inside the DataBlock, so both layouts are tried and the one whose
/// counts prove out wins), the count, the depths, then the sites.
fn read_site_data(data: &[u8], start: usize) -> Result<Vec<Site>, String> {
    for with_table in [true, false] {
        match try_site_data(data, start, with_table) {
            Ok(sites) => return Ok(sites),
            Err(_) => continue,
        }
    }
    Err("FormSiteData did not reconcile".into())
}

/// True when the bytes left after the sites are one trailing record
/// (a version-stamped, length-prefixed one — a MultiPage's).
fn is_trailing_record(data: &[u8], offset: usize) -> bool {
    if offset + 4 > data.len() {
        return false;
    }
    let version = (data[offset], data[offset + 1]);
    let length = u16::from_le_bytes([data[offset + 2], data[offset + 3]]) as usize;
    version == (0, 2) && offset + 4 + length == data.len()
}

fn try_site_data(data: &[u8], start: usize, with_table: bool) -> Result<Vec<Site>, String> {
    let mut r = Reader { data, pos: start };
    if with_table {
        let count = r.u16()?;
        for _ in 0..count {
            if r.u16()? != 0 {
                return Err("class table entry version".into());
            }
            let len = r.u16()? as usize;
            r.skip(len)?;
        }
    }
    let count_of_sites = r.u32()? as usize;
    let count_of_bytes = r.u32()? as usize;
    if count_of_sites > 10_000 {
        return Err("too many sites".into());
    }
    let end = r.pos + count_of_bytes;
    // SiteData closes the stream — or one trailing record follows (a
    // MultiPage's own record after the form's, version-stamped and
    // length-prefixed), which this checks rather than merely tolerating.
    if end != data.len() && !is_trailing_record(data, end) {
        return Err("FormSiteData does not close the stream".into());
    }
    // SiteDepthsAndTypes: one entry per site, or one counted entry per
    // run of sites sharing a depth and type.
    let depths_start = r.pos;
    let mut accounted = 0;
    while accounted < count_of_sites {
        r.u8()?;
        let type_or_count = r.u8()?;
        if type_or_count & 0x80 != 0 {
            accounted += (type_or_count & 0x7F) as usize;
            r.u8()?;
        } else {
            accounted += 1;
        }
    }
    if accounted != count_of_sites {
        return Err("depths do not account for every site".into());
    }
    r.align(depths_start, 4);
    let mut sites = Vec::with_capacity(count_of_sites);
    for _ in 0..count_of_sites {
        sites.push(read_site(data, &mut r)?);
    }
    if r.pos != end {
        return Err("sites did not end where FormSiteData says".into());
    }
    Ok(sites)
}

// MARK: FormStreamData (the blobs between the record and the sites)

fn skip_form_stream_data(data: &[u8], pos: &mut usize, mask: u32) -> Result<(), String> {
    if mask & (1 << 15) != 0 {
        skip_guid_and_picture(data, pos)?;
    }
    if mask & (1 << 20) != 0 {
        skip_font_blob(data, pos)?;
    }
    if mask & (1 << 21) != 0 {
        skip_guid_and_picture(data, pos)?;
    }
    Ok(())
}

fn skip_guid_and_picture(data: &[u8], pos: &mut usize) -> Result<(), String> {
    let mut r = Reader { data, pos: *pos };
    r.skip(16)?; // GUID
    r.skip(4)?; // the preamble word
    let len = r.u32()? as usize;
    r.skip(len)?; // the bytes the preamble's own length says
    *pos = r.pos;
    Ok(())
}

// The font GUIDs' first DWORDs ([MS-OFORMS] 2.4.6).
const GUID_STDFONT: u32 = 0x0BE3_5203;
const GUID_TEXTPROPS: u32 = 0xAFC2_0920;

fn skip_font_blob(data: &[u8], pos: &mut usize) -> Result<(), String> {
    let mut r = Reader { data, pos: *pos };
    // The full 16-byte GUID's first DWORD tells StdFont from TextProps
    // ([MS-OFORMS] 2.4.6; the blob the fixture carries is a StdFont).
    let guid = r.take(16)?;
    let tag = u32::from_le_bytes([guid[0], guid[1], guid[2], guid[3]]);
    match tag {
        GUID_STDFONT => {
            if r.u8()? != 0x01 {
                return Err("unknown StdFont version in FormStreamData".into());
            }
            r.skip(2 + 1 + 2 + 4)?; // charset, flags, weight, height
            let face = r.u8()? as usize;
            r.skip(face)?;
        }
        GUID_TEXTPROPS => {
            r.skip(2)?; // minor, major
            let cb = r.u16()? as usize;
            r.skip(cb)?;
        }
        _ => return Err(format!("unknown font GUID {tag:#010x}")),
    }
    *pos = r.pos;
    Ok(())
}

// MARK: control property records (each control's slice of `o`)
// [MS-OFORMS] 2.5.x: version(2) cb(2) mask(4, mask64 classes have two)
// then DataBlock fields aligned, then ExtraDataBlock (Size/strings),
// then TextProps for the classes that carry one.



#[derive(Debug, Clone, Copy, PartialEq)]
enum ExtraKind {
    None,
    Str,
    Size8,
}

/// A record's field: DataBlock width, whether the member is a string in
/// the extra block, with `FLEN` as its marker.
#[derive(Debug, Clone, Copy)]
struct DataField {
    bit: u32,
    size: usize,
    len_bytes: bool,
}

const FLEN: bool = true;

struct ClassSpec {
    data: &'static [DataField],
    /// The ExtraDataBlock members, in the class's own order
    /// ([MS-OFORMS] per class): (mask bit, kind).
    extra: &'static [(u32, ExtraKind)],
    /// MorphData-like classes carry the wider mask.
    mask64: bool,
}

const COMMAND_BUTTON: ClassSpec = ClassSpec {
    data: &[
        DataField { bit: 0, size: 4, len_bytes: FLEN }, // ForeColor
        DataField { bit: 1, size: 4, len_bytes: false }, // BackColor
        DataField { bit: 2, size: 4, len_bytes: false }, // VariousPropertyBits
        DataField { bit: 3, size: 4, len_bytes: FLEN },  // Caption (len)
        DataField { bit: 4, size: 4, len_bytes: false }, // PicturePosition
        DataField { bit: 6, size: 1, len_bytes: false }, // MousePointer
        DataField { bit: 8, size: 2, len_bytes: false }, // Accelerator
    ],
    extra: &[(3, ExtraKind::Str), (5, ExtraKind::Size8)],
    mask64: false,
};

const LABEL: ClassSpec = ClassSpec {
    data: &[
        DataField { bit: 0, size: 4, len_bytes: false }, // ForeColor
        DataField { bit: 1, size: 4, len_bytes: false }, // BackColor
        DataField { bit: 2, size: 4, len_bytes: false }, // VariousPropertyBits
        DataField { bit: 3, size: 4, len_bytes: FLEN },  // Caption (len)
        DataField { bit: 4, size: 4, len_bytes: false }, // PicturePosition
        DataField { bit: 6, size: 1, len_bytes: false }, // MousePointer
        DataField { bit: 7, size: 4, len_bytes: false }, // BorderColor
        DataField { bit: 8, size: 2, len_bytes: false }, // BorderStyle
        DataField { bit: 9, size: 2, len_bytes: false }, // SpecialEffect
        DataField { bit: 11, size: 2, len_bytes: false }, // Accelerator
    ],
    extra: &[(3, ExtraKind::Str), (5, ExtraKind::Size8)],
    mask64: false,
};

const MORPH_DATA: ClassSpec = ClassSpec {
    data: &[
        DataField { bit: 0, size: 4, len_bytes: false }, // VariousPropertyBits
        DataField { bit: 1, size: 4, len_bytes: false }, // BackColor
        DataField { bit: 2, size: 4, len_bytes: false }, // ForeColor
        DataField { bit: 3, size: 4, len_bytes: false }, // MaxLength
        DataField { bit: 4, size: 1, len_bytes: false }, // BorderStyle
        DataField { bit: 5, size: 1, len_bytes: false }, // ScrollBars
        DataField { bit: 6, size: 1, len_bytes: false }, // DisplayStyle
        DataField { bit: 7, size: 1, len_bytes: false }, // MousePointer
        DataField { bit: 9, size: 2, len_bytes: false }, // PasswordChar
        DataField { bit: 10, size: 4, len_bytes: false }, // ListWidth
        DataField { bit: 11, size: 2, len_bytes: false }, // BoundColumn
        DataField { bit: 12, size: 2, len_bytes: false }, // TextColumn
        DataField { bit: 13, size: 2, len_bytes: false }, // ColumnCount
        DataField { bit: 14, size: 2, len_bytes: false }, // ListRows
        DataField { bit: 15, size: 2, len_bytes: false }, // cColumnInfo
        DataField { bit: 16, size: 1, len_bytes: false }, // MatchEntry
        DataField { bit: 17, size: 1, len_bytes: false }, // ListStyle
        DataField { bit: 18, size: 1, len_bytes: false }, // ShowDropButtonWhen
        DataField { bit: 20, size: 1, len_bytes: false }, // DropButtonStyle
        DataField { bit: 21, size: 1, len_bytes: false }, // MultiSelect
        DataField { bit: 22, size: 4, len_bytes: FLEN },  // Value (len)
        DataField { bit: 23, size: 4, len_bytes: FLEN },  // Caption (len)
        DataField { bit: 24, size: 4, len_bytes: false }, // PicturePosition
        DataField { bit: 25, size: 4, len_bytes: false }, // BorderColor
        DataField { bit: 26, size: 4, len_bytes: false }, // SpecialEffect
        DataField { bit: 29, size: 2, len_bytes: false }, // Accelerator
        DataField { bit: 32, size: 4, len_bytes: FLEN },  // GroupName (len)
    ],
    extra: &[(8, ExtraKind::Size8), (22, ExtraKind::Str), (23, ExtraKind::Str), (32, ExtraKind::Str)],
    mask64: true,
};

fn spec_of(index: u16) -> Option<&'static ClassSpec> {
    match index {
        CLS_COMMAND_BUTTON => Some(&COMMAND_BUTTON),
        CLS_LABEL => Some(&LABEL),
        CLS_MORPH_DATA | CLS_TEXTBOX | CLS_LISTBOX | CLS_COMBOBOX | CLS_CHECKBOX | CLS_OPTIONBUTTON
        | CLS_TOGGLEBUTTON => Some(&MORPH_DATA),
        _ => None,
    }
}

#[derive(Debug, Default)]
struct ControlRecord {
    caption: Option<String>,
    value: Option<String>,
    /// The record's Size extra member (fmSize, HIMETRIC).
    width: Option<i32>,
    height: Option<i32>,
}

/// One control's property record; parse only what the reader needs (the
/// DataBlock's length fields and the extra members) and leave the rest
/// (TextProps and any tail) unmodelled. pyopenvba writes an unedited
/// control identically; reading stops before those bytes.
fn parse_record(slice: &[u8], spec: &ClassSpec) -> Result<ControlRecord, String> {
    let mut r = Reader::new(slice);
    let minor = r.u8()?;
    let major = r.u8()?;
    if minor != 0 || major != 0x02 {
        return Err(format!("not a control record (version {major}.{minor})"));
    }
    let cb = r.u16()? as usize;
    let mask = r.u32()?;
    let mut mask_hi = 0u32;
    if spec.mask64 {
        mask_hi = r.u32()?;
    }
    // The reference aligns on the record's base (after the versions).
    let base = 4usize;
    let has = |bit: u32| -> bool {
        if bit < 32 {
            mask & (1 << bit) != 0
        } else {
            mask_hi & (1 << (bit - 32)) != 0
        }
    };
    // (DataField.len_bytes) values: the packed length with its
    // compressed flag on the top bit.
    let mut lens: std::collections::HashMap<u32, (usize, bool)> = Default::default();
    for field in spec.data {
        if !has(field.bit) {
            continue;
        }
        r.align(base, field.size);
        let value = match field.size {
            1 => r.u8()? as u32,
            2 => r.u16()? as u32,
            _ => r.u32()?,
        };
        if field.len_bytes {
            lens.insert(field.bit, ((value & 0x7FFF_FFFF) as usize, value & 0x8000_0000 != 0));
        }
    }
    r.align(base, 4);
    // ExtraDataBlock in the class's order; a string's length rides its
    // DataBlock field, and the string pads to 4 after its bytes.
    let mut out = ControlRecord::default();
    for (bit, kind) in spec.extra {
        if !has(*bit) {
            continue;
        }
        match kind {
            ExtraKind::Size8 => {
                out.width = Some(r.i32()?);
                out.height = Some(r.i32()?);
            }
            ExtraKind::Str => {
                let (len, compressed) = lens.get(bit).copied().unwrap_or((0, false));
                let raw = r.take(len)?;
                let text = stored_string(raw, compressed);
                match *bit {
                    3 => out.caption = Some(text),
                    22 => out.value = Some(text),
                    _ => {}
                }
                r.align(base, 4);
            }
            ExtraKind::None => {}
        }
    }
    let _ = cb;
    Ok(out)
}

// MARK: through ovba

fn form_storages(project: &ovba::Project) -> Vec<String> {
    let mut found = Vec::new();
    if let Ok(list) = project.list() {
        for (_name, path) in list {
            let normalized = path.replace('\\', "/");
            // A form storage is one that owns a `\x03VBFrame` stream.
            if normalized.rsplit('/').next() == Some(STREAM_VBFRAME) {
                let storage = normalized
                    .rsplit('/')
                    .nth(1)
                    .and_then(|s| s.split('/').next_back())
                    .unwrap_or("")
                    .to_string();
                if !storage.is_empty() && !found.contains(&storage) {
                    found.push(storage);
                }
            }
        }
    }
    found
}

/// The `\x03VBFrame` designer text is plain VB6 designer source (the
/// reference files write it this way).
fn designer_text(raw: &[u8]) -> String {
    String::from_utf8_lossy(raw).into_owned()
}

/// The `\x03VBFrame` designer text's fields: Caption, ClientWidth/Height
/// (twips; 15 twips = 1 px at 96 DPI).
fn designer_fields(frame: &[u8]) -> (String, f32, f32) {
    let text = designer_text(frame);
    let field = |name: &str| -> Option<String> {
        for line in text.lines() {
            let line = line.trim();
            if let Some((key, value)) = line.split_once('=') {
                if key.trim() == name {
                    return Some(value.trim().replace('"', ""));
                }
            }
        }
        None
    };
    let caption = field("Caption").unwrap_or_default();
    let width = field("ClientWidth").and_then(|v| v.trim().parse::<f32>().ok()).unwrap_or(4500.0) / 15.0;
    let height = field("ClientHeight").and_then(|v| v.trim().parse::<f32>().ok()).unwrap_or(3000.0) / 15.0;
    (caption, width, height)
}

/// One form's designer streams: the site table on `f` and each
/// non-container control's record from its slice of `o`.
fn parse_form(
    project: &ovba::Project,
    storage: &str,
    f_stream: &[u8],
    o_stream: &[u8],
    notes: &mut Vec<String>,
    depth: usize,
    used_storages: &mut std::collections::HashSet<String>,
) -> Result<Vec<ControlDef>, String> {
    let mut r = Reader::new(f_stream);
    let v_lo = r.u8()?;
    let v_hi = r.u8()?;
    if (v_lo, v_hi) != (0, 4) {
        return Err(format!("not a FormControl stream (version {v_hi}.{v_lo})"));
    }
    let cb_form = r.u16()? as usize;
    let mask = r.u32()?;
    let after_record = 4 + cb_form;
    let mut pos = after_record;
    skip_form_stream_data(f_stream, &mut pos, mask)?;
    let sites = read_site_data(f_stream, pos)?;
    if std::env::var_os("IBIS_DUMP_SITES").is_some() {
        for site in &sites {
            eprintln!("SITE {:?} id={} class={} olen={} pos={:?}", site.name, site.id, site.clsid_cache_index, site.object_stream_size, site.position);
        }
        eprintln!("O-LEN {} (sites {}) storage={}", o_stream.len(), sites.len(), storage);
    }

    let total: usize = sites.iter().map(|s| s.object_stream_size).sum();
    if total != o_stream.len() {
        return Err(format!(
            "sites account for {total} bytes of `o` but it holds {}",
            o_stream.len()
        ));
    }

    let mut controls = Vec::new();
    let mut offset = 0usize;
    for site in &sites {
        let index = site.clsid_cache_index;
        if is_container_class(index) {
            // A Frame's children live in their own storage beside the
            // form's (`<form>/iNN/{f,o}`), as a FormControl of their
            // own. Parse it recursively and shift the children into the
            // frame; the frame itself shows as a box around them.
            let (left, top) = site.position.unwrap_or((0, 0));
            if depth < 12 {
                let base = format!("/{storage}/");
                // One container, one storage: the site's own ID names
                // its storage (`i20` for id 20); a site without that
                // match takes the next storage nobody claimed.
                let subs = child_storages(project, &base, site.id);
                let own = format!("i{site_id}", site_id = site.id);
                let picked = subs
                    .iter()
                    .find(|s| **s == own && !used_storages.contains(*s))
                    .or_else(|| subs.iter().find(|s| !used_storages.contains(*s)))
                    .cloned();
                let mut built = false;
                if let Some(sub) = &picked {
                    used_storages.insert(sub.clone());
                }
                for sub in picked.iter() {
                    let child_f = project.read_stream(&format!("{base}{sub}/f"));
                    let child_o = project.read_stream(&format!("{base}{sub}/o"));
                    if let (Ok(child_f), Ok(child_o)) = (child_f, child_o) {
                        match parse_form(project, &format!("{storage}/{sub}"), &child_f, &child_o, notes, depth + 1, used_storages) {
                            Ok(children) => {
                                let mut max_right = 0.0f32;
                                let mut max_bottom = 0.0f32;
                                for mut child in children {
                                    child.left += himetric_px(left);
                                    child.top += himetric_px(top);
                                    max_right = max_right.max(child.left + child.width);
                                    max_bottom = max_bottom.max(child.top + child.height);
                                    controls.push(child);
                                }
                                controls.push(ControlDef {
                                    name: site.name.clone(),
                                    kind: kind_of(index),
                                    left: himetric_px(left),
                                    top: himetric_px(top),
                                    width: (max_right - himetric_px(left) + 6.0).max(40.0),
                                    height: (max_bottom - himetric_px(top) + 6.0).max(24.0),
                                    caption: Some(site.name.clone()),
                                    value: None,
                                });
                                built = true;
                            }
                            Err(e) => notes.push(format!("{}: its storage {}: {e}", site.name, sub)),
                        }
                    } else {
                        notes.push(format!("{}: no child storage found", site.name));
                    }
                }
                if !built {
                    notes.push(format!("{} ({}): its children's storage did not parse", site.name, kind_of(index)));
                }
            } else {
                notes.push(format!("{} ({}): nested too deep", site.name, kind_of(index)));
            }
            // The container's own record slice: a Frame carries none in
            // the parent's `o` (its bytes are 0), but an id-carrying one
            // may still hold a few.
            offset += site.object_stream_size;
            continue;
        }
        if index >= CLASS_TABLE_BASE {
            notes.push(format!("{} (ActiveX): not modelled", site.name));
            offset += site.object_stream_size;
            continue;
        }
        let slice = o_stream.get(offset..offset + site.object_stream_size).ok_or_else(|| {
            if std::env::var_os("IBIS_DUMP_SITES").is_some() {
                eprintln!("SLICE-FAIL {} offset={} size={} olen={}", site.name, offset, site.object_stream_size, o_stream.len());
            }
            format!("{}: record walks out of `o`", site.name)
        })?;
        offset += site.object_stream_size;

        let spec = match spec_of(index) {
            Some(spec) => spec,
            None => {
                // A class without a spec (an Image, a TabStrip...):
                // the site's own position still places a placeholder,
                // so the form's layout stays legible.
                let (left, top) = site.position.unwrap_or((0, 0));
                controls.push(ControlDef {
                    name: site.name.clone(),
                    kind: kind_of(index),
                    left: himetric_px(left),
                    top: himetric_px(top),
                    width: 48.0,
                    height: 24.0,
                    caption: Some(site.name.clone()),
                    value: None,
                });
                // The offset already advanced past this record.
                continue;
            }
        };
        let record = parse_record(slice, spec).unwrap_or_else(|e| {
            notes.push(format!("{}: {e}", site.name));
            ControlRecord::default()
        });
        let (left, top) = site.position.unwrap_or((0, 0));
        let width = record.width.unwrap_or(2540);
        let height = record.height.unwrap_or(635);
        controls.push(ControlDef {
            name: site.name.clone(),
            kind: kind_of(index),
            left: himetric_px(left),
            top: himetric_px(top),
            width: himetric_px(width),
            height: himetric_px(height),
            caption: record.caption,
            value: record.value,
        });
    }
    Ok(controls)
}


/// The child storages of a container site (`iNN` sub-storages of the
/// form's storage, each carrying the container's `f`/`o`). The site's
/// own ID names its storage when it matches one; the rest go in numeric
/// order, so two frames read their own children.
fn child_storages(project: &ovba::Project, base: &str, site_id: u32) -> Vec<String> {
    let mut found: Vec<(u32, String)> = Vec::new();
    if let Ok(list) = project.list() {
        for (_name, path) in list {
            let normalized = path.replace('\\', "/");
            let Some(rest) = normalized.strip_prefix(base) else { continue };
            // `iNN/f`: the storage `iNN` under the form's own.
            let mut parts = rest.splitn(2, '/');
            let sub = parts.next().unwrap_or("");
            let leaf = parts.next().unwrap_or("");
            if leaf == "f"
                && sub.len() > 1
                && sub.starts_with('i')
                && sub[1..].chars().all(|c| c.is_ascii_digit())
            {
                let number: u32 = sub[1..].parse().unwrap_or(0);
                if !found.iter().any(|(_, name)| name == sub) {
                    found.push((number, sub.to_string()));
                }
            }
        }
    }
    found.sort_by_key(|(number, _)| *number);
    let _ = site_id;
    found.into_iter().map(|(_, name)| name).collect()
}

/// Reads the UserForms of an OOXML package. A file with no forms answers
/// `Ok(vec![])`.
pub fn read_userforms(package_path: &std::path::Path) -> Result<Vec<UserFormDef>, String> {
    let file = std::fs::File::open(package_path)
        .map_err(|e| format!("cannot open {}: {e}", package_path.display()))?;
    let mut archive = zip::ZipArchive::new(file)
        .map_err(|e| format!("{} is not an Open XML package: {e}", package_path.display()))?;
    let Some(mut entry) = archive.by_name("xl/vbaProject.bin").ok() else {
        return Ok(Vec::new());
    };
    let mut raw = Vec::new();
    entry
        .read_to_end(&mut raw)
        .map_err(|e| format!("cannot read the VBA project: {e}"))?;
    let project = ovba::open_project(raw)
        .map_err(|e| format!("cannot parse the VBA project: {e}"))?;
    let storages = form_storages(&project);

    let mut forms = Vec::with_capacity(storages.len());
    for storage in storages {
        let mut notes = Vec::new();
        let frame_path = format!("/{storage}/{STREAM_VBFRAME}");
        let (caption, width, height) = match project.read_stream(&frame_path) {
            Ok(frame) => {
                // The streams inside a form's storage may be stored
                // plain (pyopenvba) or per-[MS-OVBA] compression (VBE).
                let (c, w, h) = designer_fields(&frame);
                (c, w, h)
            }
            Err(e) => {
                notes.push(format!("no designer text ({e})"));
                (storage.clone(), 300.0, 200.0)
            }
        };
        let f_path = format!("/{storage}/f");
        let o_path = format!("/{storage}/o");
        let f_stream = project.read_stream(&f_path).unwrap_or_default();
        let o_stream = project.read_stream(&o_path).unwrap_or_default();
        let mut controls = Vec::new();
        let mut used_storages = std::collections::HashSet::new();
        match parse_form(&project, &storage, &f_stream, &o_stream, &mut notes, 0, &mut used_storages) {
            Ok(found) => controls = found,
            Err(err) => notes.push(err),
        }
        forms.push(UserFormDef {
            name: storage,
            caption,
            width,
            height,
            controls,
            notes,
        });
    }
    Ok(forms)
}


#[cfg(test)]
mod form_tests {
    use super::*;
    use std::path::PathBuf;

    fn fixture() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test/benchmark/vba-userform-demo.xlsm")
    }

    #[test]
    fn parses_the_fixture() {
        let forms = read_userforms(&fixture()).expect("read userforms");
        assert_eq!(forms.len(), 1, "one form: {forms:#?}");
        let form = &forms[0];

        assert_eq!(form.name, "UserForm1");
        assert_eq!(form.caption, "Demo");


        let names: Vec<&str> = form.controls.iter().map(|c| c.name.as_str()).collect();
        assert_eq!(names, vec!["Label1", "TextBox1", "CommandButton1"], "form: {form:#?}");
        assert_eq!(form.control("Label1").unwrap().caption.as_deref(), Some("Label1"));
        assert_eq!(form.control("CommandButton1").unwrap().caption.as_deref(), Some("CommandButton1"));
    }
}


#[cfg(test)]
mod run_tests {
    use super::*;
    use super::run::{FormSession, run_form_event, FormValue};
    use std::collections::HashMap;
    use std::path::PathBuf;

    fn fixture() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test/benchmark/vba-userform-demo.xlsm")
    }

    #[test]
    fn runs_the_click_event() {
        let package = fixture();
        let forms = read_userforms(&package).expect("forms");
        let form = &forms[0];
        let mut state: HashMap<String, FormValue> = form
            .controls
            .iter()
            .map(|c| (c.name.clone(), FormValue::default()))
            .collect();
        state.insert(
            "TextBox1".to_string(),
            FormValue { text: "typed in the box".into(), checked: false, list: Vec::new(), selected: None },
        );
        let session = FormSession { package: package.clone(), workbook: None, save_back: false };
        let output = std::env::temp_dir().join("ibis-vba-form-run.xlsx");
        let run = run_form_event(&session, form, "CommandButton1_Click", &state, &output)
            .expect("the click event runs");
        let label = run.state.get("Label1").expect("Label1");
        assert_eq!(label.text, "typed in the box", "the run: {run:#?}");
    }

    #[test]
    fn chains_the_events_through_the_plane() {
        // Initialize seeds the plane; the follow-up run reads its file.
        let package = fixture();
        let forms = read_userforms(&package).expect("forms");
        let form = &forms[0];
        let empty: HashMap<String, FormValue> = form
            .controls
            .iter()
            .map(|c| (c.name.clone(), FormValue::default()))
            .collect();
        let session = FormSession { package: package.clone(), workbook: None, save_back: false };
        let first = std::env::temp_dir().join("ibis-vba-form-init.xlsx");
        let run1 = run_form_event(&session, form, "UserForm_Initialize", &empty, &first)
            .expect("initialize");
        assert_eq!(
            run1.state.get("TextBox1").unwrap().text,
            "hello from vba",
            "{run1:#?}"
        );
        // The click on seeded state: workbook = the previous run's file.
        let session2 = FormSession { package: package.clone(), workbook: Some(first), save_back: false };
        let second = std::env::temp_dir().join("ibis-vba-form-click.xlsx");
        let run2 = run_form_event(&session2, form, "CommandButton1_Click", &run1.state, &second)
            .expect("click");
        assert_eq!(run2.state.get("Label1").unwrap().text, "hello from vba", "{run2:#?}");
    }
}

#[cfg(test)]
mod crud_tests {
    use super::*;
    use super::run::{FormSession, FormValue, run_form_event};
    use std::collections::HashMap;
    use std::path::PathBuf;

    fn fixture() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test/benchmark/vba-userform-demo-crud.xlsm")
    }

    #[test]
    fn runs_the_crud_form() {
        let package = fixture();
        let forms = read_userforms(&package).expect("forms");
        let form = &forms[0];
        assert_eq!(form.name, "oUserForm");
        let mut state: HashMap<String, FormValue> = form
            .controls
            .iter()
            .map(|c| (c.name.clone(), FormValue::default()))
            .collect();
        // TableData at A4:C7: Mr/Louis, Mrs/Dion, Dr/Charles (3 rows).
        let session = FormSession { package: package.clone(), workbook: None, save_back: false };
        let out = std::env::temp_dir().join("ibis-vba-crud-init.xlsx");
        let run = run_form_event(&session, form, "UserForm_Initialize", &state, &out).expect("initialize");
        // Row 50+ probes (the list plane's cells):

        let list = &run.state.get("lbList").unwrap().list;
        assert_eq!(list.len(), 3, "three table rows: {list:?}");
        assert!(list[0].contains("Louis"), "row one: {list:?}");
        // The Create: the texts in and a click.
        state = run.state.clone();
        state.insert("txtTitle".to_string(), FormValue { text: "Ms".into(), checked: false, list: list.clone(), selected: None });
        state.insert("txtName".into(), FormValue { text: "Ada".into(), checked: false, list: list.clone(), selected: None });
        state.insert("txtAge".into(), FormValue { text: "39".into(), checked: false, list: list.clone(), selected: None });
        let out2 = std::env::temp_dir().join("ibis-vba-crud-create.xlsx");
        let session = FormSession { package: package.clone(), workbook: Some(out.clone()), save_back: false };
        let run2 = run_form_event(&session, form, "cmdCreate_Click", &state, &out2).expect("create");
        let list2 = &run2.state.get("lbList").unwrap().list;
        assert_eq!(list2.len(), 4, "a row added: {list2:?}");
        assert!(list2[3].contains("Ada"), "the new row's name: {list2:?}");
        // The Delete: one row off.
        let run3 = run_form_event(
            &FormSession { package: package.clone(), workbook: Some(out2.clone()), save_back: false },
            form,
            "cmdDelete_Click",
            &run2.state,
            &out,
        )
        .expect("delete");
        let list3 = &run3.state.get("lbList").unwrap().list;
        assert_eq!(list3.len(), 3, "a row deleted: {list3:?}");
    }
}


#[cfg(test)]
mod combobox_tests {
    use super::*;
    use super::run::{FormSession, FormValue, run_form_event};
    use std::collections::HashMap;
    use std::path::PathBuf;

    fn fixture() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test/benchmark/vba-demo-combobox-sheetnames.xlsm")
    }

    #[test]
    fn fills_the_combobox_with_the_sheet_names() {
        let package = fixture();
        let forms = read_userforms(&package).expect("forms");
        let form = &forms[0];
        assert_eq!(form.name, "UserForm1");
        let combo = form.control("ComboBox1").expect("the designer carries a ComboBox");
        assert_eq!(combo.kind, "ComboBox");
        let state: HashMap<String, FormValue> = form
            .controls
            .iter()
            .map(|c| (c.name.clone(), FormValue::default()))
            .collect();
        let session = FormSession { package: package.clone(), workbook: None, save_back: false };
        let out = std::env::temp_dir().join("ibis-vba-combobox-init.xlsx");
        let run = run_form_event(&session, form, "UserForm_Initialize", &state, &out).expect("initialize");
        let list = &run.state.get("ComboBox1").unwrap().list;
        // The workbook carries 31 worksheets; the enumerated galleries
        // fill the list (`For i = Count To 1 Step -1`).
        assert_eq!(list.len(), 31, "the sheets fill the box: {list:?}");
        assert!(list[0].contains("bmhp") || list[0].contains("obat"), "a real sheet name: {list:?}");
    }
}

#[cfg(test)]
mod save_back_tests {
    use super::*;
    use super::run::{FormSession, FormValue, run_form_event};
    use std::collections::HashMap;
    use std::path::PathBuf;

    /// The run with `save_back` replaces the original in place: the
    /// cells carry the run's writes, and the file keeps its VBA
    /// project (a macro-enabled workbook stays one).
    #[test]
    fn save_back_overwrites_the_original_keeping_vba() {
        let source = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test/benchmark/vba-userform-demo-crud.xlsm");
        let work = std::env::temp_dir().join("ibis-vba-save-back-test.xlsm");
        std::fs::copy(&source, &work).expect("copy the fixture");
        let forms = read_userforms(&work).expect("forms");
        let form = &forms[0];
        let state: HashMap<String, FormValue> = form
            .controls
            .iter()
            .map(|c| (c.name.clone(), FormValue::default()))
            .collect();
        let session = FormSession { package: work.clone(), workbook: None, save_back: true };
        let out = std::env::temp_dir().join("ibis-vba-save-back-unused.xlsx");
        let run = run_form_event(&session, form, "UserForm_Initialize", &state, &out).expect("initialize");
        // The run reports the original as its file (the next run
        // chains on the original, not a copy).
        assert_eq!(run.file, work);
        // The original was rewritten: it still opens as a macro
        // workbook and keeps its VBA project and form.
        let forms2 = read_userforms(&work).expect("the saved file reopens with its forms");
        assert_eq!(forms2[0].name, "oUserForm");
        let vba_present = {
            let file = std::fs::File::open(&work).expect("open");
            let mut archive = zip::ZipArchive::new(file).expect("zip");
            archive.by_name("xl/vbaProject.bin").is_ok()
        };
        assert!(vba_present, "the VBA project survives the save-back");
        // And the file is a macro-enabled workbook by its content type.
        let file = std::fs::File::open(&work).expect("open");
        let mut archive = zip::ZipArchive::new(file).expect("zip");
        let mut types = String::new();
        {
            let mut entry = archive.by_name("[Content_Types].xml").expect("content types");
            use std::io::Read as _;
            entry.read_to_string(&mut types).expect("read");
        }
        assert!(
            types.contains("sheet.macroEnabled.main+xml") && types.contains("vbaProject"),
            "the package declares its VBA: {types}"
        );
        let _ = std::fs::remove_file(&work);
    }
}

#[cfg(test)]
mod frame_tests {
    use super::*;
    use std::path::PathBuf;

    fn fixture() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test/benchmark/vba-demo-frames.xlsm")
    }

    /// A Frame's children live in their own `iNN` storages; the parser
    /// walks them and shifts the children into the frame, so a
    /// frame-heavy form keeps its whole layout.
    #[test]
    fn reads_the_children_of_frames() {
        let forms = read_userforms(&fixture()).expect("forms");
        let form = forms.iter().find(|f| f.name == "a_t_form").expect("a_t_form");
        assert!(
            form.notes.is_empty(),
            "the frame-heavy form parses whole: {:?}",
            form.notes
        );
        // The frame carries a box control named after it.
        let frame = form.control("frame_transction_info").expect("the frame");
        assert_eq!(frame.kind, "Frame");
        // The frame's children ride inside it (their positions shifted
        // by the frame's own).
        let combo = form.control("cbo_how_long_data").expect("a frame child");
        let frame_top = frame.top;
        assert!(
            combo.top >= frame_top,
            "the child sits inside the frame: child top {} vs frame top {}",
            combo.top,
            frame_top
        );
        // Two frames: each read its own storage (no children twice).
        let count = form
            .controls
            .iter()
            .filter(|c| c.name == "cbo_how_long_data")
            .count();
        assert_eq!(count, 1, "no duplicated children");
    }
}
