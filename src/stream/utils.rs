//! The module includes extra utility tooling to help glue logic together
use quick_xml::{parser::Parser, Reader, Writer};
use std::io::{BufReader, Read, Seek, SeekFrom, Write};
use zip::{read::{ZipFile, ZipFileSeek}, result::ZipError, write::{FileOptionExtension, FileOptions}, ZipArchive, ZipWriter};
use super::xlsx::errors::XlsxError;

pub(crate) type Key = usize;

// ported from calamine https://github.com/tafia/calamine/tree/master
pub(crate) fn xml_reader<'a, RS: Read + Seek>(
    zip: &'a mut ZipArchive<RS>,
    path: &str,
    offset: Option<SeekFrom>
) -> Option<Result<Reader<BufReader<ZipFileSeek<'a, RS>>>, XlsxError>> {
    let actual_path = zip
        .file_names()
        .find(|n| n.eq_ignore_ascii_case(path))?
        .to_owned();
    match zip.by_name_seek(&actual_path) {
        Ok(mut f) => {
            if let Some(offset) = offset {
                f.seek(offset);
            };
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

pub(crate) trait XmlWriter<W: Write> {
    /// Allows us to piece up how we will write from objects to xml
    fn write_xml<'a>(&self, writer: &'a mut Writer<W>, tag_name: &'a str) -> Result<&'a mut Writer<W>, XlsxError>;
}

pub(crate) trait Save<W: Write + Seek, EX: FileOptionExtension>: XmlWriter<W> {
    /// Save file in a zip folder aka .xlsx
    fn save(&mut self, writer: &mut ZipWriter<W>, options: FileOptions<EX>) -> Result<(), XlsxError>;
}
