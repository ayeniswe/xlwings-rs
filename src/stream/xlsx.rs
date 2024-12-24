//! The module holds all logic to fully deserialize a .xlsx file and its contents
use super::utils::xml_reader;
use crate::errors::XcelmateError;
use quick_xml::{
    events::{attributes::Attribute, Event},
    name::QName,
    Reader,
};
use std::{
    borrow::Cow,
    collections::HashMap,
    io::{BufReader, Read, Seek},
};
use zip::{read::ZipFile, ZipArchive};

/// The `Rgb` promotes better api usage with hexadecimal coloring
#[derive(Debug, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
enum Rgb {
    Custom((u8, u8, u8)),
}

/// The `Color` denotes the type of coloring system to
/// use since excel has builtin coloring to choose that will map to `theme` but
/// for custom specfic coloring `rgb` is used
///
/// Default is `Theme((1, None))` = black
#[derive(Debug, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
enum Color {
    /// Builtin theme from excel color palette selector which includes theme id and tint value
    Theme { id: u32, tint: Option<String> },
    /// RGB color model
    Rgb(Rgb),
}
impl Default for Color {
    fn default() -> Self {
        Color::Theme { id: 1, tint: None }
    }
}

/// The `RichTextProperty` denotes all styling options
/// that can be added to text
#[derive(Debug, Default, Clone, PartialEq, Eq, PartialOrd, Hash, Ord)]
struct RichTextProperty {
    bold: bool,
    underline: bool,
    /// Double underline
    double: bool,
    italic: bool,
    size: String,
    color: Color,
    /// Font type
    font: String,
    /// Font family
    family: u32,
    /// Font scheme
    scheme: String,
}

/// The `StringType` allows to differentiate between some strings that could
/// have leading and trailing spaces
#[derive(Debug, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
enum StringType {
    // String that have leading or trailing spaces
    Preserve(String),
    // Normal string with no leading or trailing spaces
    NoPreserve(String),
}

/// The `TextType` represents either a rich text or plaintext string and needs
/// to track ref counts to handle removal.
#[derive(Debug, PartialEq, Eq, Clone, PartialOrd, Hash, Ord)]
enum TextType {
    RichText(Vec<StringPiece>),
    PlainText(StringType),
}

/// The `StringPiece` represents a string that is contained in a richtext denoted by having a `TextType::RichText`.
/// The pieces of text can be with styling or no styling
#[derive(Debug, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
struct StringPiece {
    /// The styling applied to text
    props: Option<RichTextProperty>,
    // The actual raw value of text
    value: StringType,
}

/// The `SharedStringTable` provides an efficient way to map strings
/// to their corresponding integer references used in the spreadsheet.
/// Excel uses shared string tables to store unique strings once and reference them by index.
///
/// **Note**: rich text will always have atleast greater than one `StringItem` and plain text will only have a single `StringItem`
#[derive(Default)]
struct SharedStringTable {
    table: HashMap<TextType, usize>,
    /// Allows lookups with integer reference and tracks reference counter
    reverse_table: HashMap<usize, u32>,
    count: u32,
}

/// The `Xlsx` struct represents an Excel workbook stored in an OpenXML format (XLSX).
/// It encapsulates foundational pieces of a workbook
struct Xlsx<RS> {
    /// The zip archive containing all files of the XLSX workbook.
    zip: ZipArchive<RS>,
    /// The shared string table for efficient mapping of shared strings.
    shared_string_table: SharedStringTable,
}

// Ported from calamine https://github.com/tafia/calamine/tree/master
impl<RS: Read + Seek> Xlsx<RS> {
    fn read_shared_strings(&mut self) -> Result<(), XcelmateError> {
        let mut xml = match xml_reader(&mut self.zip, "xl/sharedStrings.xml") {
            None => return Ok(()),
            Some(x) => x?,
        };
        let mut buf = Vec::with_capacity(1024);
        let mut idx = 0; // Track index for referencing updates, etc
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sst" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"count") => {
                                    self.shared_string_table.count =
                                        a.unescape_value()?.to_string().parse::<u32>()?
                                }
                                // We dont care about unique count since that will be the len() of the table in SharedStringTable
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"si" => {
                    if let Some(s) = Xlsx::<RS>::read_string(&mut xml, e.name())? {
                        self.shared_string_table.table.insert(s.clone(), idx);
                        self.shared_string_table.reverse_table.insert(idx, 1);
                        idx += 1;
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sst" => break,
                Ok(Event::Eof) => return Err(XcelmateError::XmlEof("sst".into())),
                Err(e) => return Err(XcelmateError::Xml(e)),
                _ => (),
            }
        }
        Ok(())
    }

    /// Read either a simple or richtext string
    fn read_string(
        xml: &mut Reader<BufReader<ZipFile>>,
        QName(closing): QName,
    ) -> Result<Option<TextType>, XcelmateError> {
        let mut buf = Vec::with_capacity(1024);
        let mut val_buf = Vec::with_capacity(1024);
        let mut rich_buffer: Option<TextType> = None;
        let mut is_phonetic_text = false;
        let mut props: Option<RichTextProperty> = None;
        let mut preserve = false;
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"r" => {
                    if rich_buffer.is_none() {
                        // Use a buffer since richtext has multiples <r> and <t> for the same cell
                        rich_buffer = Some(TextType::RichText(Vec::new()));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPr" => {
                    if props.is_none() {
                        props = Some(RichTextProperty::default());
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPh" => {
                    is_phonetic_text = true;
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"sz" => {
                    if let Some(p) = props.as_mut() {
                        for attr in e.attributes() {
                            if let Ok(a) = attr {
                                match a.key {
                                    QName(b"val") => p.size = a.unescape_value()?.to_string(),
                                    _ => (),
                                }
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"color" => {
                    if let Some(p) = props.as_mut() {
                        for attr in e.attributes() {
                            if let Ok(a) = attr {
                                match a.key {
                                    QName(b"rgb") => {
                                        let val = a.unescape_value()?.to_string();
                                        // The first two letter are ignored since they response to alpha
                                        let base16 = 16u32;
                                        let red = u8::from_str_radix(&val[2..4], base16)?;
                                        let green = u8::from_str_radix(&val[4..6], base16)?;
                                        let blue = u8::from_str_radix(&val[6..8], base16)?;
                                        p.color = Color::Rgb(Rgb::Custom((red, green, blue)));
                                    }
                                    QName(b"theme") => {
                                        p.color = Color::Theme {
                                            id: a.unescape_value()?.to_string().parse::<u32>()?,
                                            tint: None,
                                        };
                                    }
                                    QName(b"tint") => match p.color {
                                        Color::Theme { id, .. } => {
                                            p.color = Color::Theme {
                                                id,
                                                tint: Some(a.unescape_value()?.to_string()),
                                            }
                                        }
                                        _ => (),
                                    },
                                    _ => (),
                                }
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"rFont" => {
                    if let Some(p) = props.as_mut() {
                        for attr in e.attributes() {
                            if let Ok(a) = attr {
                                match a.key {
                                    QName(b"val") => p.font = a.unescape_value()?.to_string(),
                                    _ => (),
                                }
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"family" => {
                    if let Some(p) = props.as_mut() {
                        for attr in e.attributes() {
                            if let Ok(a) = attr {
                                match a.key {
                                    QName(b"val") => {
                                        p.family = a.unescape_value()?.parse::<u32>()?
                                    }
                                    _ => (),
                                }
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"scheme" => {
                    if let Some(p) = props.as_mut() {
                        for attr in e.attributes() {
                            if let Ok(a) = attr {
                                match a.key {
                                    QName(b"val") => p.scheme = a.unescape_value()?.to_string(),
                                    _ => (),
                                }
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"b" => {
                    if let Some(p) = props.as_mut() {
                        p.bold = true
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"i" => {
                    if let Some(p) = props.as_mut() {
                        p.italic = true
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"u" => {
                    if let Some(p) = props.as_mut() {
                        p.underline = true;
                        for attr in e.attributes() {
                            if let Ok(a) = attr {
                                match a.key {
                                    QName(b"val") => {
                                        p.double = true;
                                        // No longer can be true if doubled
                                        p.underline = false;
                                    }
                                    _ => (),
                                }
                            }
                        }
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == closing => {
                    return Ok(rich_buffer);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"rPh" => {
                    is_phonetic_text = false;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"t" && !is_phonetic_text => {
                    val_buf.clear();
                    let mut value = String::new();
                    loop {
                        match xml.read_event_into(&mut val_buf)? {
                            Event::Text(t) => value.push_str(&t.unescape()?),
                            Event::End(end) if end.name() == e.name() => break,
                            Event::Eof => return Err(XcelmateError::XmlEof("t".to_string())),
                            _ => (),
                        }
                    }

                    // Determine if string leading and trailing spaces should be preserved
                    for attr in e.attributes() {
                        match attr? {
                            Attribute {
                                key: QName(b"xml:space"),
                                value: Cow::Borrowed(b"preserve"),
                            } => preserve = true,
                            _ => (),
                        }
                    }

                    if let Some(ref mut s) = rich_buffer {
                        let value = if preserve {
                            StringPiece {
                                props: props.clone(),
                                value: StringType::Preserve(value),
                            }
                        } else {
                            StringPiece {
                                props: props.clone(),
                                value: StringType::NoPreserve(value),
                            }
                        };
                        match s {
                            TextType::RichText(pieces) => pieces.push(value),
                            _ => (),
                        }

                        // Reset since other <t> tags may not be preserved or have properties
                        preserve = false;
                        props = None;
                    } else {
                        // Consume any remaining events up to expected closing tag
                        xml.read_to_end_into(QName(closing), &mut val_buf)?;

                        if preserve {
                            return Ok(Some(TextType::PlainText(StringType::Preserve(value))));
                        } else {
                            return Ok(Some(TextType::PlainText(StringType::NoPreserve(value))));
                        }
                    }
                }
                Ok(Event::Eof) => return Err(XcelmateError::XmlEof("".into())),
                Err(e) => return Err(XcelmateError::Xml(e)),
                _ => (),
            }
        }
    }

    /// Decrement the total count of all strings creation
    fn decrement_count(&mut self) {
        if self.shared_string_table.count > 0 {
            self.shared_string_table.count -= 1;
        }
    }

    /// Increment the total count of all strings creation
    fn increment_count(&mut self) {
        self.shared_string_table.count += 1;
    }

    /// Get the unique count
    fn unique_count(&self) -> usize {
        self.shared_string_table.table.len()
    }

    /// Get the shared string integer ref
    pub(crate) fn shared_string_ref(&mut self, item: TextType) -> usize {
        self.add_to_table(item)
    }

    /// Increment the shared string reference counter
    fn increment_shared_string_ref(&mut self, int_ref: usize) {
        if let Some(ref_counter) = self.shared_string_table.reverse_table.get_mut(&int_ref) {
            *ref_counter += 1;
        }
    }

    /// As every string is added, the shared table must reflect the changes in count
    fn add_to_table(&mut self, item: TextType) -> usize {
        let mut int_ref = 0;
        if let Some(i) = self.shared_string_table.table.get(&item) {
            int_ref = *i;
            // We have to keep track of update to references
            *self
                .shared_string_table
                .reverse_table
                .get_mut(&int_ref)
                .unwrap() += 1
        } else {
            int_ref = self.unique_count();
            self.shared_string_table.table.insert(item, int_ref);
            // References should only be 1 for a new item
            self.shared_string_table.reverse_table.insert(int_ref, 1);
        }
        self.shared_string_table.count += 1;
        int_ref
    }

    /// As every string is removed, the shared table must reflect the changes in count.
    /// Shared strings can ONLY be removed when `ref_count` is `0`
    ///
    /// # Returns
    /// The new `ref_count` or `None` if already has been removed.
    /// A `ref_count` of `Some(0)` denotes the entry has just been removed
    pub(crate) fn remove_from_table(&mut self, item: TextType) -> Option<u32> {
        let mut remove = false;
        if let Some(int_ref) = self.shared_string_table.table.get(&item) {
            let ref_counter = self.shared_string_table.reverse_table.get(int_ref).unwrap();
            if *ref_counter == 1 {
                remove = true
            }
        }

        self.decrement_count();

        if remove {
            if let Some(int_ref) = self.shared_string_table.table.remove(&item) {
                self.shared_string_table.reverse_table.remove(&int_ref);
                Some(0)
            } else {
                None
            }
        } else {
            if let Some(int_ref) = self.shared_string_table.table.get(&item) {
                let ref_counter = self
                    .shared_string_table
                    .reverse_table
                    .get_mut(&int_ref)
                    .unwrap();
                *ref_counter -= 1;
                Some(ref_counter.clone())
            } else {
                None
            }
        }
    }
}

#[cfg(test)]
mod shared_string_api {
    use crate::stream::xlsx::{
        Color, Rgb, RichTextProperty, SharedStringTable, StringPiece, StringType, TextType, Xlsx,
    };
    use std::fs::File;
    use zip::ZipArchive;

    #[test]
    fn query_richtext_match() {
        let file = File::open("tests/workbook01.xlsx").unwrap();
        let mut xlsx = Xlsx {
            zip: ZipArchive::new(file).unwrap(),
            shared_string_table: SharedStringTable::default(),
        };
        let pieces = vec![
            StringPiece {
                props: None,
                value: StringType::Preserve("The ".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: true,
                    underline: true,
                    double: false,
                    italic: true,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::NoPreserve("big".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: true,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::Preserve(" example".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::Preserve(" ".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: true,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::Preserve("of ".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: true,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::NoPreserve("some".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::Preserve(" ".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme {
                        id: 5,
                        tint: Some("0.39997558519241921".into()),
                    },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::NoPreserve("rich".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::Preserve(" ".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: true,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::NoPreserve("text".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::Preserve(" ".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Rgb(Rgb::Custom((186, 155, 203))),
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::NoPreserve("here".into()),
            },
        ];
        xlsx.read_shared_strings().unwrap();

        // Item should exist
        let actual = xlsx
            .shared_string_table
            .table
            .get(&TextType::RichText(pieces.clone()));
        assert_eq!(actual, Some(&0));

        // Item should be of type rich text
        let actual = xlsx
            .shared_string_table
            .table
            .keys()
            .into_iter()
            .collect::<Vec<&TextType>>();
        assert_eq!(*actual[0], TextType::RichText(pieces))
    }

    #[test]
    fn query_richtext_no_match() {
        let file = File::open("tests/workbook01.xlsx").unwrap();
        let mut xlsx = Xlsx {
            zip: ZipArchive::new(file).unwrap(),
            shared_string_table: SharedStringTable::default(),
        };
        let pieces = vec![
            StringPiece {
                props: None,
                value: StringType::Preserve("The ".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: true,
                    underline: true,
                    double: false,
                    italic: true,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::NoPreserve("big".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: true,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::Preserve(" example".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::Preserve(" ".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: true,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::Preserve("of ".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: true,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::NoPreserve("some".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::Preserve(" ".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme {
                        id: 5,
                        tint: Some("0.39997558519241921".into()),
                    },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::NoPreserve("rich".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::Preserve(" ".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: true,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::NoPreserve("text".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Theme { id: 1, tint: None },
                    font: "Calibri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::Preserve(" ".into()),
            },
            StringPiece {
                props: Some(RichTextProperty {
                    bold: false,
                    underline: false,
                    double: false,
                    italic: false,
                    size: "11".into(),
                    color: Color::Rgb(Rgb::Custom((186, 155, 203))),
                    font: "Calibrri".into(),
                    family: 2,
                    scheme: "minor".into(),
                }),
                value: StringType::NoPreserve("here".into()),
            },
        ];
        xlsx.read_shared_strings().unwrap();

        // Item should not exist
        let actual = xlsx
            .shared_string_table
            .table
            .get(&TextType::RichText(pieces.clone()));
        assert_eq!(actual, None);
    }

    #[test]
    fn query_plaintext_no_match() {
        let file = File::open("tests/workbook02.xlsx").unwrap();
        let mut xlsx = Xlsx {
            zip: ZipArchive::new(file).unwrap(),
            shared_string_table: SharedStringTable::default(),
        };
        let key = TextType::PlainText(StringType::Preserve("The".into()));
        xlsx.read_shared_strings().unwrap();

        // Item should not exist
        let actual = xlsx.shared_string_table.table.get(&key);
        assert_eq!(actual, None);
    }

    #[test]
    fn query_plaintext_match() {
        let file = File::open("tests/workbook02.xlsx").unwrap();
        let mut xlsx = Xlsx {
            zip: ZipArchive::new(file).unwrap(),
            shared_string_table: SharedStringTable::default(),
        };
        let key = TextType::PlainText(StringType::Preserve("The ".into()));
        xlsx.read_shared_strings().unwrap();

        // Item should exist
        let actual = xlsx.shared_string_table.table.get(&key);
        assert_eq!(actual, Some(&0));

        // Item should be of type plain text
        let actual = xlsx
            .shared_string_table
            .table
            .keys()
            .into_iter()
            .collect::<Vec<&TextType>>();
        assert_eq!(
            *actual[0],
            TextType::PlainText(StringType::Preserve("The ".into()))
        )
    }

    #[test]
    fn table_count_is_deserialized() {
        let file = File::open("tests/workbook02.xlsx").unwrap();
        let mut xlsx = Xlsx {
            zip: ZipArchive::new(file).unwrap(),
            shared_string_table: SharedStringTable::default(),
        };
        xlsx.read_shared_strings().unwrap();

        // Count should increment
        let actual = xlsx.shared_string_table.count;
        assert_eq!(actual, 1);
    }

    #[test]
    fn add_new_shared_string_to_table() {
        let file = File::open("tests/workbook02.xlsx").unwrap();
        let mut xlsx = Xlsx {
            zip: ZipArchive::new(file).unwrap(),
            shared_string_table: SharedStringTable::default(),
        };
        xlsx.read_shared_strings().unwrap();
        let item = TextType::PlainText(StringType::NoPreserve("The".into()));

        // Should add a new item
        let actual = xlsx.shared_string_ref(item);
        assert_eq!(actual, 1);

        // Total count should be incremented
        let actual = xlsx.shared_string_table.count;
        assert_eq!(actual, 2);
    }

    #[test]
    fn add_shared_string_to_table_no_duplicates() {
        let file = File::open("tests/workbook02.xlsx").unwrap();
        let mut xlsx = Xlsx {
            zip: ZipArchive::new(file).unwrap(),
            shared_string_table: SharedStringTable::default(),
        };
        xlsx.read_shared_strings().unwrap();
        let item = TextType::PlainText(StringType::Preserve("The ".into()));

        // Should not readd item that exists already
        let actual = xlsx.shared_string_ref(item);
        assert_eq!(actual, 0);

        // Total count should be incremented
        let actual = xlsx.shared_string_table.count;
        assert_eq!(actual, 2);

        // Reference count should be incremented
        let actual = xlsx.shared_string_table.reverse_table.get(&0).unwrap();
        assert_eq!(actual, &2);
    }

    #[test]
    fn remove_shared_string_from_table() {
        let file = File::open("tests/workbook02.xlsx").unwrap();
        let mut xlsx = Xlsx {
            zip: ZipArchive::new(file).unwrap(),
            shared_string_table: SharedStringTable::default(),
        };
        xlsx.read_shared_strings().unwrap();
        let item = TextType::PlainText(StringType::Preserve("The ".into()));

        // Should give zero ref count since removed
        let actual = xlsx.remove_from_table(item.clone());
        assert_eq!(actual, Some(0));

        // Table should be all empty
        let ref_counter_table = xlsx.shared_string_table.reverse_table.keys().len();
        let shared_string_table = xlsx.shared_string_table.table.keys().len();
        assert_eq!(ref_counter_table, 0);
        assert_eq!(shared_string_table, 0);

        // Total count should be decremented
        let actual = xlsx.shared_string_table.count;
        assert_eq!(actual, 0);
        
        // Should no longer exist
        let actual = xlsx.remove_from_table(item);
        assert_eq!(actual, None);
    }
}
