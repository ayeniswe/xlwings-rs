//! The module holds all logic to fully deserialize a .xlsx file and its contents
mod shared_string_table;
mod stylesheet;
mod sheet;

use crate::errors::XcelmateError;
use shared_string_table::SharedStringTable;
use std::{
    fs::File,
    io::{Read, Seek},
};
use stylesheet::Stylesheet;
use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

use super::utils::Save;

/// The `Xlsx` struct represents an Excel workbook stored in an OpenXML format (XLSX).
/// It encapsulates foundational pieces of a workbook
struct Xlsx<RS> {
    /// The zip archive containing all files of the XLSX workbook.
    zip: ZipArchive<RS>,
    /// The shared string table for efficient mapping of shared strings.
    shared_string_table: SharedStringTable,
    /// The stylesheet for formating cells.
    style: Stylesheet,
}

impl<RS: Read + Seek> Xlsx<RS> {
    fn read_shared_strings(&mut self) -> Result<(), XcelmateError> {
        self.shared_string_table.read_shared_strings(&mut self.zip)
    }
    fn read_stylesheet(&mut self) -> Result<(), XcelmateError> {
        self.style.read_stylesheet(&mut self.zip)
    }
    fn save(&mut self, name: &str) -> Result<(), XcelmateError> {
        let mut zip = ZipWriter::new(File::create(name)?);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        let _ = self.shared_string_table.save(&mut zip, options);
        let _ = self.style.save(&mut zip, options);
        Ok(())
    }
}
