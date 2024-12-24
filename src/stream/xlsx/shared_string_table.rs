//! The module holds all logic to fully deserialize the sharedStrings.xml in the .xlsx file
use crate::{errors::XcelmateError, stream::utils::{xml_reader, Key}};
use quick_xml::{
    events::{attributes::Attribute, Event},
    name::QName,
    Reader,
};
use std::{
    borrow::Cow,
    collections::HashMap,
    io::{BufReader, Read, Seek},
    sync::Arc,
};
use zip::{read::ZipFile, ZipArchive};

use super::stylesheet::{Color, FontProperty, Rgb};

type SharedStringRef = Arc<SharedString>;

/// The `StringType` allows to differentiate between some strings that could
/// have leading and trailing spaces
#[derive(Debug, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
enum StringType {
    // String that have leading or trailing spaces
    Preserve(String),
    // Normal string with no leading or trailing spaces
    NoPreserve(String),
}

/// The `SharedString` represents either a rich text or plaintext string and needs
/// to track ref counts to handle removal.
#[derive(Debug, PartialEq, Eq, Clone, PartialOrd, Hash, Ord)]
enum SharedString {
    RichText(Vec<StringPiece>),
    PlainText(StringType),
}

/// The `StringPiece` represents a string that is contained in a richtext denoted by having a `SharedString::RichText`.
/// The pieces of text can be with styling or no styling
#[derive(Debug, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
struct StringPiece {
    /// The styling applied to text
    props: Option<FontProperty>,
    // The actual raw value of text
    value: StringType,
}

/// The `SharedStringTable` provides an efficient way to map strings
/// to their corresponding integer references used in the spreadsheet.
/// Excel uses shared string tables to store unique strings once and reference them by index.
///
/// **Note**: rich text will always have atleast greater than one `StringItem` and plain text will only have a single `StringItem`
#[derive(Default)]
pub(crate) struct SharedStringTable {
    table: HashMap<SharedStringRef, Key>,
    reverse_table: HashMap<Key, SharedStringRef>,
    count: u32,
}

impl SharedStringTable {
    // Ported from calamine https://github.com/tafia/calamine/tree/master
    pub(crate) fn read_shared_strings<'a, RS: Read + Seek>(
        &mut self,
        zip: &'a mut ZipArchive<RS>,
    ) -> Result<(), XcelmateError> {
        let mut xml = match xml_reader(zip, "xl/sharedStrings.xml") {
            None => return Ok(()),
            Some(x) => x?,
        };
        let mut buf = Vec::with_capacity(1024);
        let mut idx: usize = 0; // Track index for referencing updates, etc
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sst" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"count") => {
                                    self.count = a.unescape_value()?.to_string().parse::<u32>()?
                                }
                                // We dont care about unique count since that will be the len() of the table in SharedStringTable
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"si" => {
                    if let Some(s) = SharedStringTable::read_string(&mut xml, e.name())? {
                        let text = Arc::new(s);
                        self.reverse_table.insert(idx, text.clone());
                        self.table.insert(text, idx);
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

    // Ported from calamine https://github.com/tafia/calamine/tree/master
    /// Read either a simple or richtext string
    fn read_string(
        xml: &mut Reader<BufReader<ZipFile>>,
        QName(closing): QName,
    ) -> Result<Option<SharedString>, XcelmateError> {
        let mut buf = Vec::with_capacity(1024);
        let mut val_buf = Vec::with_capacity(1024);
        let mut rich_buffer: Option<SharedString> = None;
        let mut is_phonetic_text = false;
        let mut props: Option<FontProperty> = None;
        let mut preserve = false;
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"r" => {
                    if rich_buffer.is_none() {
                        // Use a buffer since richtext has multiples <r> and <t> for the same cell
                        rich_buffer = Some(SharedString::RichText(Vec::new()));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPr" => {
                    if props.is_none() {
                        props = Some(FontProperty::default());
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
                            SharedString::RichText(pieces) => pieces.push(value),
                            _ => (),
                        }

                        // Reset since other <t> tags may not be preserved or have properties
                        preserve = false;
                        props = None;
                    } else {
                        // Consume any remaining events up to expected closing tag
                        xml.read_to_end_into(QName(closing), &mut val_buf)?;

                        if preserve {
                            return Ok(Some(SharedString::PlainText(StringType::Preserve(value))));
                        } else {
                            return Ok(Some(SharedString::PlainText(StringType::NoPreserve(
                                value,
                            ))));
                        }
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == closing => {
                    return Ok(rich_buffer);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"rPh" => {
                    is_phonetic_text = false;
                }
                Ok(Event::Eof) => return Err(XcelmateError::XmlEof("".into())),
                Err(e) => return Err(XcelmateError::Xml(e)),
                _ => (),
            }
        }
    }

    /// Decrement the total count of all strings creation
    fn decrement_count(&mut self) {
        if self.count > 0 {
            self.count -= 1;
        }
    }

    /// Increment the total count of all strings creation
    fn increment_count(&mut self) {
        self.count += 1;
    }

    /// Get the total count of all strings creation
    fn count(&self) -> u32 {
        self.count
    }

    /// Get the unique count
    fn unique_count(&self) -> usize {
        self.table.len()
    }

    /// Get the shared string ref
    pub(crate) fn shared_string_ref(&mut self, item: SharedString) -> Option<SharedStringRef> {
        self.increment_count();
        if let Some(i) = self.table.get(&item) {
            Some(self.reverse_table.get(i).unwrap().clone())
        } else {
            None
        }
    }

    /// Get the shared string ref from key
    pub(crate) fn get_shared_string_ref_from_key(&mut self, key: Key) -> Option<SharedStringRef> {
        self.count += 1;
        if let Some(i) = self.reverse_table.get(&key) {
            Some(i.clone())
        } else {
            None
        }
    }

    /// As every string is added, the shared table must reflect the changes in count
    pub(crate) fn add_to_table(&mut self, item: SharedString) -> SharedStringRef {
        self.increment_count();
        let int_ref = self.unique_count();
        let item = Arc::new(item);
        self.reverse_table.insert(int_ref, item.clone());
        self.table.insert(item.clone(), int_ref);
        item
    }

    /// As every string is removed, the shared table must reflect the changes in count.
    /// Shared strings can ONLY be removed when smart pointer count is `2`. The pointer will always have atleast `2` counts since
    /// at load time we keep a pointer in two hashmaps to support bidirectional access
    ///
    /// # Returns
    /// The new smart pointer count or `None` if already has been removed.
    /// A return value of `Some(0)` indicates the entry has been removed from both tables
    pub(crate) fn remove_from_table(&mut self, item: SharedString) -> Option<usize> {
        self.decrement_count();
        if let Some(key) = self.table.get(&item) {
            let shared_string_ref = self.reverse_table.get(key).unwrap();
            let count = Arc::strong_count(shared_string_ref);

            // 2 count since both hashmaps hold ref to shared string
            const TABLE_COUNT: usize = 2;
            if count == TABLE_COUNT {
                let int_ref = self.table.remove(&item).unwrap();
                self.reverse_table.remove(&int_ref);
                Some(0)
            } else {
                Some(count - TABLE_COUNT)
            }
        } else {
            None
        }
    }
}

#[cfg(test)]
mod shared_string_api {
    use crate::stream::xlsx::{
        shared_string_table::{
            Color, Rgb, FontProperty, SharedString, SharedStringTable, StringPiece, StringType,
        },
        Xlsx,
    };
    use std::{fs::File, sync::Arc};
    use zip::ZipArchive;

    fn init(path: &str) -> SharedStringTable {
        let file = File::open(path).unwrap();
        let mut zip = ZipArchive::new(file).unwrap();
        let mut sst = SharedStringTable::default();
        sst.read_shared_strings(&mut zip).unwrap();
        sst
    }

    #[test]
    fn query_richtext_match() {
        let mut sst = init("tests/workbook01.xlsx");
        let pieces = vec![
            StringPiece {
                props: None,
                value: StringType::Preserve("The ".into()),
            },
            StringPiece {
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
        // Item should already exist
        let actual = sst
            .shared_string_ref(SharedString::RichText(pieces.clone()))
            .unwrap();
        assert_eq!(Arc::strong_count(&actual), 3);
    }

    #[test]
    fn query_richtext_no_match() {
        let mut sst = init("tests/workbook01.xlsx");
        let pieces = vec![
            StringPiece {
                props: None,
                value: StringType::Preserve("The ".into()),
            },
            StringPiece {
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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
                props: Some(FontProperty {
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

        // Item should not exist
        let actual = sst.shared_string_ref(SharedString::RichText(pieces.clone()));
        assert_eq!(actual, None);
    }

    #[test]
    fn query_plaintext_no_match() {
        let mut sst = init("tests/workbook01.xlsx");

        // Item should not exist
        let actual =
            sst.shared_string_ref(SharedString::PlainText(StringType::Preserve("The".into())));
        assert_eq!(actual, None);
    }

    #[test]
    fn query_plaintext_match() {
        let mut sst = init("tests/workbook02.xlsx");

        // Item should exist
        let actual = sst
            .shared_string_ref(SharedString::PlainText(StringType::Preserve("The ".into())))
            .unwrap();
        assert_eq!(Arc::strong_count(&actual), 3);
    }

    #[test]
    fn table_count_is_deserialized() {
        let sst = init("tests/workbook02.xlsx");

        // Count should increment
        let actual = sst.count();
        assert_eq!(actual, 1);
    }

    #[test]
    fn add_new_shared_string_to_table() {
        let mut sst = init("tests/workbook02.xlsx");
        let item = SharedString::PlainText(StringType::NoPreserve("The".into()));

        // Should add a new item
        let actual = sst.add_to_table(item);
        assert_eq!(Arc::strong_count(&actual), 3);

        // Total count should be incremented
        let actual = sst.count();
        assert_eq!(actual, 2);
    }

    #[test]
    fn remove_shared_string_from_table() {
        let mut sst = init("tests/workbook02.xlsx");
        let item = SharedString::PlainText(StringType::Preserve("The ".into()));

        // Should give zero ref count since removed
        let actual = sst.remove_from_table(item.clone());
        assert_eq!(actual, Some(0));

        // Total count should be empty
        let actual = sst.count();
        assert_eq!(actual, 0);

        // Should no longer exist
        let actual = sst.remove_from_table(item);
        assert_eq!(actual, None);
    }

    #[test]
    fn do_not_remove_shared_string_from_table_when_ref_exists() {
        let mut sst = init("tests/workbook02.xlsx");
        let item = SharedString::PlainText(StringType::Preserve("The ".into()));

        {
            // Should show one ref exists still
            let shared_string_ref = sst
                .get_shared_string_ref_from_key(0)
                .unwrap();
            let actual = sst.remove_from_table(item.clone());
            assert_eq!(actual, Some(1));

            // Total count should not empty
            let actual = sst.count();
            assert_eq!(actual, 1);
        }

        // Should no longer exist after drop
        let actual = sst.remove_from_table(item);
        assert_eq!(actual, Some(0));
    }
}
