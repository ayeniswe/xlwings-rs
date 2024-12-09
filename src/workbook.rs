use thiserror::Error;

pub fn open_workbook<T: AsRef<Path>>(path: T) -> Result<bool, WorkbookError> {
    if let Ok(file) = File::open(path) {
        if let Ok(mut zip) = ZipArchive::new(file) {
            // Plan to generate all separate data and combine to workbook struct
            let stylsheet = open_xml_file::<Style>(&mut zip, "xl/styles.xml").unwrap();
            let content_types =
                open_xml_file::<ContentType>(&mut zip, "[Content_Types].xml").unwrap();
            let rel = open_xml_file::<Relationship>(&mut zip, "_rels/.rels").unwrap();
            let book = open_xml_file::<Book>(&mut zip, "xl/workbook.xml").unwrap();
            let shared_strings =
                open_xml_file::<SharedString>(&mut zip, "xl/sharedStrings.xml").unwrap();
            let drawings = open_xml_file::<Drawing>(&mut zip, "xl/drawings/drawing1.xml").unwrap();
            let themes = open_xml_file::<Theme>(&mut zip, "xl/theme/theme1.xml").unwrap();
            let book_rel =
                open_xml_file::<Relationship>(&mut zip, "xl/_rels/workbook.xml.rels").unwrap();
            let sheet_rel =
                open_xml_file::<Relationship>(&mut zip, "xl/worksheets/_rels/sheet1.xml.rels")
                    .unwrap();
            let sheet = open_xml_file::<Sheet>(&mut zip, "xl/worksheets/sheet1.xml").unwrap();

            // LOGIC TO SAVE TODO
            //
            // let new_file = File::create("example.xlsx").unwrap();
            // let mut new_zip = ZipWriter::new(new_file);
            // let options =
            //     SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
            // new_zip.start_file("_rels/.rels", options).unwrap();
            // new_zip.write(rel.to_string().as_bytes()).unwrap();
            // new_zip.start_file("[Content_Types].xml", options).unwrap();
            // new_zip.write(content_types.to_string().as_bytes()).unwrap();
            // new_zip.start_file("xl/styles.xml", options).unwrap();
            // new_zip.write(stylsheet.to_string().as_bytes()).unwrap();
            // new_zip.start_file("xl/workbook.xml", options).unwrap();
            // new_zip.write(book.to_string().as_bytes()).unwrap();
            // new_zip
            //     .start_file("xl/_rels/workbook.xml.rels", options)
            //     .unwrap();
            // new_zip.write(book_rel.to_string().as_bytes()).unwrap();
            // new_zip.start_file("xl/sharedStrings.xml", options).unwrap();
            // new_zip
            //     .write(shared_strings.to_string().as_bytes())
            //     .unwrap();
            // // Could be dynamic but not common
            // new_zip.start_file("xl/theme/theme1.xml", options).unwrap();
            // new_zip.write(themes.to_string().as_bytes()).unwrap();
            // new_zip
            //     .start_file("xl/drawings/drawing1.xml", options)
            //     .unwrap();
            // new_zip.write(drawings.to_string().as_bytes()).unwrap();
            // // Dynamic and can grow sheets amounts
            // new_zip
            //     .start_file("xl/worksheets/sheet1.xml", options)
            //     .unwrap();
            // new_zip.write(sheet.to_string().as_bytes()).unwrap();
            // new_zip
            //     .start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
            //     .unwrap();
            // new_zip.write(sheet_rel.to_string().as_bytes()).unwrap();
            Ok(true)
        } else {
            Err(WorkbookError::InvalidFileFormat)
        }
    } else {
        Err(WorkbookError::FileNotFound)
    }
}

#[derive(Error, Debug)]
pub enum WorkbookError {
    #[error("File is not a valid Excel file.")]
    InvalidFileFormat,
    #[error("Excel file not found.")]
    FileNotFound,
}
