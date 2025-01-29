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

pub(crate) trait Save<W: Write + Seek, EX: FileOptionExtension>: XmlWriter<W> {
    /// Save file in a zip folder aka .xlsx
    fn save(
        &mut self,
        writer: &mut ZipWriter<W>,
        options: FileOptions<EX>,
    ) -> Result<(), XlsxError>;
}

pub trait XmlWriter<W: Write> {
    /// Allows us to write objects to xml
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &'a str,
    ) -> Result<&'a mut Writer<W>, XlsxError>;
}

pub trait XmlReader<B: BufRead> {
    /// Allows us to read xml into a custom object
    fn read_xml<'a>(
        &mut self,
        tag_name: &'a str,
        xml: &'a mut Reader<B>,
        closing: &'a str,
        propagated_event: Option<Result<Event<'_>, quick_xml::Error>>,
    ) -> Result<(), XlsxError>;
}
