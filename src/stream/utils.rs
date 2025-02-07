//! The module includes extra utility tooling to help glue logic together
use super::xlsx::errors::XlsxError;
use quick_xml::{events::Event, parser::Parser, Error, Reader, Writer};
use std::io::{BufRead, BufReader, Cursor, Read, Seek, SeekFrom, Write};
use zip::{
    read::{ZipFile, ZipFileSeek},
    result::ZipError,
    write::{FileOptionExtension, FileOptions},
    ZipArchive, ZipWriter,
};

pub(crate) type Key = usize;

// ported from calamine https://github.com/tafia/calamine/tree/master
pub(crate) fn xml_reader<'a, RS: Read + Seek>(
    zip: &'a mut ZipArchive<RS>,
    path: &str,
    offset: Option<SeekFrom>,
) -> Option<Result<Reader<BufReader<ZipFile<'a>>>, XlsxError>> {
    let actual_path = zip
        .file_names()
        .find(|n| n.eq_ignore_ascii_case(path))?
        .to_owned();
    match zip.by_name(&actual_path) {
        Ok(mut f) => {
            let mut r = Reader::from_reader(BufReader::new(f));
            let config = r.config_mut();
            config.check_end_names = false;
            config.trim_text(false);
            config.check_comments = false;
            config.expand_empty_elements = false;
            Some(Ok(r))
        }
        Err(ZipError::FileNotFound) => None,
        Err(e) => Some(Err(XlsxError::Zip(e))),
    }
}

/// A trait for saving an XML-based file into a ZIP archive `.xlsx`.
///
/// This trait extends [`XmlWriter<W>`] and provides functionality for serializing and saving the file
/// within a ZIP archive for `.xlsx`.
///
/// # Arguments
/// - `writer`: A mutable reference to a [`ZipWriter<W>`] used for writing the file inside a ZIP archive.
/// - `options`: A [`FileOptions<EX>`] struct specifying configuration settings for file saving.
pub(crate) trait Save<W: Write + Seek, EX: FileOptionExtension>: XmlWriter<W> {
    /// Save file in a zip folder aka .xlsx
    fn save(
        &mut self,
        writer: &mut ZipWriter<W>,
        options: FileOptions<EX>,
    ) -> Result<(), XlsxError>;
}
/// A trait for writing objects as XML elements.
///
/// This trait defines a method for serializing an object into XML format using a [`Writer<W>`].
///
/// # Arguments
/// - `writer`: A mutable reference to a [`Writer<W>`], which writes XML data.
/// - `tag_name`: The name of the XML element to be written (e.g., `"name"` for `<name>...</name>`).
pub trait XmlWriter<W: Write> {
    /// Allows us to serialize an object into XML format
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &'a str,
    ) -> Result<&'a mut Writer<W>, XlsxError>;
}
/// A trait for reading XML data into a custom object.
///
/// This trait provides a method for parsing XML elements while allowing recursive parsing of nested elements.
/// It ensures that the first read event is not skipped when delegating parsing to a nested structure.
///
/// # Arguments
/// - `tag_name`: The name of the XML element being read (e.g., `"name"` for `<name>...</name>`).
/// - `xml`: A mutable reference to a [`Reader<B>`], used for reading the XML stream.
/// - `closing`: The expected closing tag for start elements. This has no effect on self-closing elements like `<empty/>`.
/// - `propagated_event`: A mutable reference to an [`Option<Result<Event<'static>, quick_xml::Error>>`] that allows events
///   to be passed down to nested elements. This prevents the first event from being consumed before it can be processed
///   in a recursive parsing structure.
pub trait XmlReader<B: BufRead> {
    /// Allows us to deserialize xml into a custom object
    fn read_xml<'a>(
        &mut self,
        tag_name: &'a str,
        xml: &'a mut Reader<B>,
        closing: &'a str,
        propagated_event: &'a mut Option<Result<Event<'static>, quick_xml::Error>>,
    ) -> Result<(), XlsxError>;
}
