//! The module holds all logic to fully deserialize a .xlsx file and its contents
mod shared_string_table;
mod sheet;
mod stylesheet;
pub(crate) mod errors;

use super::utils::Save;
use errors::XlsxError;
use shared_string_table::SharedStringTable;
use sheet::Sheet;
use std::{
    collections::HashMap,
    fs::File,
    io::{Read, Seek},
};
use stylesheet::Stylesheet;
use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

/// The `Xlsx` struct represents an Excel workbook stored in an OpenXML format (XLSX).
/// It encapsulates foundational pieces of a workbook
pub(crate) struct Xlsx<RS> {
    /// The zip archive containing all files of the XLSX workbook.
    zip: ZipArchive<RS>,
    /// The shared string table for efficient mapping of shared strings.
    shared_string_table: SharedStringTable,
    /// The stylesheet for formating cells.
    style: Stylesheet,
    // All sheets in workbook
    sheets: HashMap<String, Sheet>,
}

impl<RS: Read + Seek> Xlsx<RS> {
    fn read_shared_strings(&mut self) -> Result<(), XlsxError> {
        self.shared_string_table.read_shared_strings(&mut self.zip)
    }
    fn read_stylesheet(&mut self) -> Result<(), XlsxError> {
        self.style.read_stylesheet(&mut self.zip)
    }
    fn read_sheet(&mut self, name: &str) -> Result<(), XlsxError> {
        // Plan to use workbook reader to create the sheets that
        // will store the paths location for lazy reading of sheets
        if let Some(sheet) = self.sheets.get_mut(name) {
            Ok(sheet.read_sheet(&mut self.zip)?)
        } else {
            Err(XlsxError::SheetNotFound(name.to_string()))
        }
    }
    fn save(&mut self, name: &str) -> Result<(), XlsxError> {
        let mut zip = ZipWriter::new(File::create(name)?);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        let _ = self.shared_string_table.save(&mut zip, options);
        let _ = self.style.save(&mut zip, options);
        Ok(())
    }
}
