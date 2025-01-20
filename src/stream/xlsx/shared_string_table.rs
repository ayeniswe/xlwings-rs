//! The module holds all logic to fully deserialize the sharedStrings.xml in the .xlsx file
use super::{errors::XlsxError, stylesheet::FontProperty, Stylesheet};
use crate::stream::utils::{xml_reader, Key, Save, XmlWriter};
use bimap::BiBTreeMap;
use quick_xml::{
    events::{attributes::Attribute, BytesDecl, BytesText, Event},
    name::QName,
    Reader, Writer,
};
use std::{
    borrow::Cow,
    io::{BufRead, Read, Seek, Write},
    sync::Arc,
};
use zip::{
    write::{FileOptionExtension, FileOptions},
    ZipArchive,
};

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
impl<W: Write> XmlWriter<W> for StringType {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &str,
    ) -> Result<&'a mut Writer<W>, XlsxError> {
        match self {
            StringType::Preserve(s) => writer
                .create_element(tag_name)
                .with_attribute(("xml:space", "preserve"))
                .write_text_content(BytesText::new(s))?,
            StringType::NoPreserve(s) => writer
                .create_element(tag_name)
                .write_text_content(BytesText::new(s))?,
        };
        Ok(writer)
    }
}

/// The `SharedString` represents either a rich text or plaintext string and needs
/// to track ref counts to handle removal.
#[derive(Debug, PartialEq, Eq, Clone, PartialOrd, Hash, Ord)]
pub(crate) enum SharedString {
    RichText(Vec<StringPiece>),
    PlainText(StringType),
}
impl<W: Write> XmlWriter<W> for SharedString {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &'a str,
    ) -> Result<&'a mut Writer<W>, XlsxError> {
        let writer = writer.create_element(tag_name);
        match self {
            SharedString::RichText(pieces) => {
                Ok(writer.write_inner_content::<_, XlsxError>(|writer| {
                    // <r>
                    for piece in pieces {
                        writer
                            .create_element("r")
                            .write_inner_content::<_, XlsxError>(|writer| {
                                piece.write_xml(writer, "rPr")?;
                                Ok(())
                            })?;
                    }
                    Ok(())
                })?)
            }
            SharedString::PlainText(st) => {
                // <t>
                Ok(writer.write_inner_content::<_, XlsxError>(|writer| {
                    st.write_xml(writer, "t")?;
                    Ok(())
                })?)
            }
        }
    }
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
impl<W: Write> XmlWriter<W> for StringPiece {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &str,
    ) -> Result<&'a mut Writer<W>, XlsxError> {
        if let Some(props) = &self.props {
            props.write_xml(writer, tag_name)?;
        }
        self.value.write_xml(writer, "t")?;
        Ok(writer)
    }
}

/// The `SharedStringTable` provides an efficient way to map strings
/// to their corresponding integer references used in the spreadsheet.
/// Excel uses shared string tables to store unique strings once and reference them by index.
///
/// **Note**: rich text will always have atleast greater than one `StringItem` and plain text will only have a single `StringItem`
#[derive(Default)]
pub(crate) struct SharedStringTable {
    table: BiBTreeMap<SharedStringRef, Key>,
    count: u32,
}
impl<W: Write> XmlWriter<W> for SharedStringTable {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &str,
    ) -> Result<&'a mut Writer<W>, XlsxError> {
        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;
        writer
            .create_element(tag_name)
            .with_attributes(vec![
                (
                    "xmlns",
                    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                ),
                ("count", self.count.to_string().as_str()),
                ("uniqueCount", self.table.len().to_string().as_str()),
            ])
            .write_inner_content::<_, XlsxError>(|writer| {
                // <si>
                for (si, _) in self.table.right_range(0..self.table.len()) {
                    si.write_xml(writer, "si")?;
                }
                Ok(())
            })?;
        Ok(writer)
    }
}
impl SharedStringTable {
    // Ported from calamine https://github.com/tafia/calamine/tree/master
    pub(crate) fn read_shared_strings<'a, RS: Read + Seek>(
        &mut self,
        zip: &'a mut ZipArchive<RS>,
    ) -> Result<(), XlsxError> {
        let mut xml = match xml_reader(zip, "xl/sharedStrings.xml", None) {
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
                                    self.count = a.unescape_value()?.parse::<u32>()?
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
                        self.table.insert(text, idx);
                        idx += 1;
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sst" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sst".into())),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
        Ok(())
    }

    // Ported from calamine https://github.com/tafia/calamine/tree/master
    /// Read either a simple or richtext string
    fn read_string<B: BufRead>(
        xml: &mut Reader<B>,
        QName(closing): QName,
    ) -> Result<Option<SharedString>, XlsxError> {
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
                    props = Some(Stylesheet::read_font(xml, e.name())?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rPh" => {
                    is_phonetic_text = true;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"t" && !is_phonetic_text => {
                    val_buf.clear();
                    let mut value = String::new();
                    loop {
                        match xml.read_event_into(&mut val_buf)? {
                            Event::Text(t) => value.push_str(&t.unescape()?),
                            Event::End(end) if end.name() == e.name() => break,
                            Event::Eof => return Err(XlsxError::XmlEof("t".to_string())),
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
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("".into())),
                Err(e) => return Err(XlsxError::Xml(e)),
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
        if let Some(i) = self.table.get_by_left(&item) {
            Some(self.table.get_by_right(i).unwrap().clone())
        } else {
            None
        }
    }

    /// Get the shared string ref from key
    pub(crate) fn get_shared_string_ref_from_key(&mut self, key: Key) -> Option<SharedStringRef> {
        self.count += 1;
        if let Some(i) = self.table.get_by_right(&key) {
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
        if let Some(key) = self.table.get_by_left(&item) {
            let shared_string_ref = self.table.get_by_right(key).unwrap();
            let count = Arc::strong_count(shared_string_ref);

            // We should only remove when no one reference it no longer
            // the 1 means the table only has a reference
            if count == 1 {
                self.table.remove_by_left(&item).unwrap();
                Some(0)
            } else {
                Some(count - 1)
            }
        } else {
            None
        }
    }
}

impl<W: Write + Seek, EX: FileOptionExtension> Save<W, EX> for SharedStringTable {
    fn save(
        &mut self,
        writer: &mut zip::ZipWriter<W>,
        options: FileOptions<EX>,
    ) -> Result<(), XlsxError> {
        writer.start_file("xl/sharedStrings.xml", options)?;
        self.write_xml(&mut Writer::new(writer), "sst")?;
        Ok(())
    }
}

#[cfg(test)]
mod shared_string_unittests {

    mod shared_string_api {
        use crate::stream::{
            utils::Save,
            xlsx::{
                shared_string_table::{
                    FontProperty, SharedString, SharedStringTable, StringPiece, StringType,
                },
                stylesheet::{Color, FormatState, Rgb},
            },
        };
        use std::{fs::File, io::Cursor, sync::Arc};
        use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

        fn init(path: &str) -> SharedStringTable {
            let file = File::open(path).unwrap();
            let mut zip = ZipArchive::new(file).unwrap();
            let mut sst = SharedStringTable::default();
            sst.read_shared_strings(&mut zip).unwrap();
            sst
        }

        #[test]
        fn get_richtext_match() {
            let mut sst = init("tests/workbook01.xlsx");
            let pieces = vec![
                StringPiece {
                    props: None,
                    value: StringType::Preserve("The ".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        bold: FormatState::Enabled,
                        underline: FormatState::Enabled,
                        italic: FormatState::Enabled,
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::NoPreserve("big".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        bold: FormatState::Enabled,
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::Preserve(" example".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::Preserve(" ".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        italic: FormatState::Enabled,
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::Preserve("of ".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        underline: FormatState::Enabled,
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::NoPreserve("some".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::Preserve(" ".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        size: "11".into(),
                        color: Color::Theme {
                            id: 5,
                            tint: Some("0.39997558519241921".into()),
                        },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::NoPreserve("rich".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::Preserve(" ".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        double: FormatState::Enabled,
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::NoPreserve("text".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::Preserve(" ".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        size: "11".into(),
                        color: Color::Rgb {
                            value: Rgb::Custom(186, 155, 203),
                            tint: None,
                        },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::NoPreserve("here".into()),
                },
            ];
            // Item should already exist
            let actual = sst
                .shared_string_ref(SharedString::RichText(pieces.clone()))
                .unwrap();
            assert_eq!(Arc::strong_count(&actual), 2);
        }

        #[test]
        fn get_richtext_no_match() {
            let mut sst = init("tests/workbook01.xlsx");
            let pieces = vec![
                StringPiece {
                    props: None,
                    value: StringType::Preserve("The ".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        bold: FormatState::Enabled,
                        underline: FormatState::Enabled,

                        italic: FormatState::Enabled,
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::NoPreserve("big".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        bold: FormatState::Enabled,
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::Preserve(" example".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::Preserve(" ".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        italic: FormatState::Enabled,
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::Preserve("of ".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        underline: FormatState::Enabled,
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::NoPreserve("some".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::Preserve(" ".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        size: "11".into(),
                        color: Color::Theme {
                            id: 5,
                            tint: Some("0.39997558519241921".into()),
                        },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::NoPreserve("rich".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::Preserve(" ".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        double: FormatState::Enabled,
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::NoPreserve("text".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        size: "11".into(),
                        color: Color::Theme { id: 1, tint: None },
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::Preserve(" ".into()),
                },
                StringPiece {
                    props: Some(FontProperty {
                        size: "11".into(),
                        color: Color::Rgb {
                            value: Rgb::Custom(186, 155, 203),
                            tint: None,
                        },
                        font: "Calibrri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    value: StringType::NoPreserve("here".into()),
                },
            ];

            // Item should not exist
            let actual = sst.shared_string_ref(SharedString::RichText(pieces.clone()));
            assert_eq!(actual, None);
        }

        #[test]
        fn get_plaintext_no_match() {
            let mut sst = init("tests/workbook01.xlsx");

            // Item should not exist
            let actual =
                sst.shared_string_ref(SharedString::PlainText(StringType::Preserve("The".into())));
            assert_eq!(actual, None);
        }

        #[test]
        fn get_plaintext_match() {
            let mut sst = init("tests/workbook02.xlsx");

            // Item should exist
            let actual = sst
                .shared_string_ref(SharedString::PlainText(StringType::Preserve("The ".into())))
                .unwrap();
            assert_eq!(Arc::strong_count(&actual), 2);
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
            assert_eq!(Arc::strong_count(&actual), 2);

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
                let _shared_string_ref = sst.get_shared_string_ref_from_key(0).unwrap();
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

        #[test]
        fn save_file() {
            let mut sst = init("tests/workbook01.xlsx");
            let mut zip = ZipWriter::new(Cursor::new(Vec::<u8>::new()));
            sst.save(
                &mut zip,
                SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
            )
            .unwrap();

            // Verify all data is written
            assert_eq!(zip.finish().unwrap().into_inner().len(), 479);
        }
    }
}
