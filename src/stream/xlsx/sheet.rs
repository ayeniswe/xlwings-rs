use super::stylesheet::{Color, Stylesheet};
use crate::{
    errors::XcelmateError,
    stream::utils::{xml_reader, Key, Save, XmlWriter},
};
use bimap::{BiBTreeMap, BiHashMap, BiMap};
use quick_xml::{
    events::{BytesDecl, BytesStart, Event},
    name::QName,
    Reader, Writer,
};
use std::{
    collections::HashMap,
    io::{BufRead, Read, Seek, Write},
    ops::RangeInclusive,
    sync::Arc,
};
use zip::{
    write::{FileOptionExtension, FileOptions},
    ZipArchive,
};

type Row = u32;
type Col = u16;

const MAX_COLUMNS: u16 = 16_384;
const MAX_ROWS: u32 = 1_048_576;

struct Worksheet {
    name: String,
    tab_color: Color,
    fit: bool,
    dimensions: ((Col, Row), (Col, Row)),
}

impl Worksheet {
    pub(crate) fn read_sheet<'a, RS: Read + Seek>(
        &mut self,
        zip: &'a mut ZipArchive<RS>,
        sheet: &str,
    ) -> Result<(), XcelmateError> {
        let mut xml = match xml_reader(zip, &format!("xl/worksheets/{}", sheet)) {
            None => return Err(XcelmateError::StylesMissing),
            Some(x) => x?,
        };
        let mut buf = Vec::with_capacity(1024);
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                ////////////////////
                // SHEET PROPERTIES
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheetPr" => {
                    ////////////////////
                    // SHEET PROPERTIES Attrs
                    /////////////
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"codeName") => self.name = a.unescape_value()?.to_string(),
                                _ => (),
                            }
                        }
                    }
                }
                ////////////////////
                // SHEET PROPERTIES nth-1
                /////////////
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"tabColor" => {
                    self.tab_color = Stylesheet::read_color(e.attributes())?;
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"pageSetUpPr" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"fitToPage") => {
                                    if a.unescape_value()?.parse::<u32>()? == 1 {
                                        self.fit = true
                                    }
                                }
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dimension" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"ref") => {
                                    let dimension = a.unescape_value()?.to_string();
                                    let dimensions = dimension.split(":").collect::<Vec<&str>>();
                                    let m = Worksheet::convert_dim_to_row_col(dimensions[0])?;
                                    let x = Worksheet::convert_dim_to_row_col(dimensions[1])?;
                                    self.dimensions = (m, x);
                                }
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"worksheet" => break,
                Ok(Event::Eof) => return Err(XcelmateError::XmlEof("worksheet".into())),
                Err(e) => return Err(XcelmateError::Xml(e)),
                _ => (),
            }
        }
        Ok(())
    }

    fn convert_dim_to_row_col(dimension: &str) -> Result<(Col, Row), XcelmateError> {
        let mut col = String::new();
        let mut row = String::new();

        // Seperate dimensions
        for c in dimension.chars() {
            if c.is_alphabetic() {
                col.push(c)
            } else if c.is_digit(10) {
                row.push(c)
            } else {
                return Err(XcelmateError::ExcelDimensionParseError(
                    dimension.to_string(),
                ));
            }
        }

        // Calculate column to integer
        let mut result = 0;
        for c in col.chars() {
            if c < 'A' || c > 'Z' {
                return Err(XcelmateError::ExcelDimensionParseError(
                    dimension.to_string(),
                ));
            }
            let value = c as i32 - 'A' as i32 + 1;
            result = result * 26 + value;
        }

        let row = row.parse::<u32>()?;
        let col = result as u16;
        if row > MAX_ROWS {
            return Err(XcelmateError::ExcelMaxRowExceeded);
        } else if col > MAX_COLUMNS {
            return Err(XcelmateError::ExcelMaxColumnExceeded);
        } else {
            Ok((col - 1, row - 1))
        }
    }
}

mod sheet_unittests {

    mod sheet_api {
        use crate::stream::xlsx::sheet::Worksheet;

        #[test]
        fn test_convert_dim_to_row_col() {
            let actual = Worksheet::convert_dim_to_row_col("B1").unwrap();
            assert_eq!(actual, (1, 0))
        }

        #[test]
        fn test_convert_dim_to_row_col_max_col() {
            let actual = Worksheet::convert_dim_to_row_col("XFE1").err().unwrap().to_string();
            assert_eq!(actual, "excel columns can not exceed 16,384")
        }

        #[test]
        fn test_convert_dim_to_row_col_max_row() {
            let actual = Worksheet::convert_dim_to_row_col("XFD1048577").err().unwrap().to_string();
            assert_eq!(actual, "excel rows can not exceed 1,048,576")
        }

        #[test]
        fn test_convert_dim_to_row_col_invalid_char() {
            let actual = Worksheet::convert_dim_to_row_col("*1048577").err().unwrap().to_string();
            assert_eq!(actual, "can not parse excel dimension: *1048577")
        }

        #[test]
        fn test_convert_dim_to_row_col_not_within_a_to_z_range() {
            let actual = Worksheet::convert_dim_to_row_col("B京1048577").err().unwrap().to_string();
            assert_eq!(actual, "can not parse excel dimension: B京1048577")
        }
    }
}
