use super::stylesheet::{Color, Stylesheet};
use crate::{
    errors::XcelmateError,
    stream::utils::{xml_reader, Key, Save, XmlWriter},
};
use bimap::{BiBTreeMap, BiHashMap, BiMap};
use num_enum::{FromPrimitive, IntoPrimitive};
use quick_xml::{
    events::{BytesDecl, BytesStart, Event},
    name::QName,
    Reader, Writer,
};
use std::{
    borrow::Cow,
    collections::HashMap,
    default,
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

#[derive(Default, Debug, Clone, FromPrimitive, IntoPrimitive, PartialEq, Eq)]
#[repr(u8)]
pub(crate) enum GridlineColor {
    #[default]
    Automatic = 0,
    Black = 8,
    Turquoise = 15,
    Brown = 60,
    Pink = 14,
    OliveGreen = 59,
    DarkGreen = 58,
    DarkTeal = 56,
    DarkBlue = 18,
    Indigo = 62,
    Gray80 = 63,
    Gray50 = 23,
    Gray40 = 55,
    Gray25 = 22,
    White = 9,
    IceBlue = 31,
    Blue = 12,
    Teal = 21,
    OceanBlue = 30,
    Plum = 25,
    Lavender = 46,
    Violet = 20,
    BlueGray = 54,
    LightBlue = 48,
    SkyBlue = 40,
    PaleBlue = 44,
    Coral = 29,
    DarkRed = 16,
    Aqua = 49,
    LightTurquoise = 27,
    DarkPurple = 28,
    SeaGreen = 57,
    LightGreen = 42,
    BrightGreen = 11,
    Yellow = 13,
    Ivory = 26,
    LightYellow = 43,
    DarkYellow = 19,
    Lime = 50,
    Orange = 53,
    LightOrange = 52,
    Gold = 51,
    Tan = 47,
    Rose = 45,
    Periwinkle = 24,
    Red = 10,
    Green = 17,
}

#[derive(Debug, PartialEq, Clone, Eq)]
struct Worksheet {
    name: String,
    tab_color: Color,
    fit: bool,
    dimensions: ((Col, Row), (Col, Row)),
    show_grid: bool,
    zoom_scale: u16, // make sure to write zoomScale and zoomNormalScale
    view_id: u32,
    show_tab: bool,
    show_ruler: bool,
    show_header: bool,
    show_outline_symbol: bool,
    default_grid_color: GridlineColor,
    show_formula: bool,
    show_zero: bool,
    use_rtl: bool,
}
impl Default for Worksheet {
    fn default() -> Self {
        Self {
            show_grid: true,
            show_zero: true,
            show_outline_symbol: true,
            ..Default::default()
        }
    }
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheetViews" => {
                    let mut view_buf = Vec::with_capacity(1024);
                    loop {
                        view_buf.clear();
                        let event = xml.read_event_into(&mut view_buf);
                        match event {
                            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                                if e.local_name().as_ref() == b"sheetView" =>
                            {
                                for attr in e.attributes() {
                                    if let Ok(a) = attr {
                                        match a.key {
                                            QName(b"showGridLines") => {
                                                if a.value == Cow::Borrowed(&[48]) {
                                                    self.show_grid = false;
                                                }
                                            }
                                            QName(b"zoomScaleNormal") => {
                                                self.zoom_scale =
                                                    a.unescape_value()?.parse::<u16>()?;
                                            }
                                            QName(b"zoomScale") => {
                                                self.zoom_scale =
                                                    a.unescape_value()?.parse::<u16>()?;
                                            }
                                            QName(b"workbookViewId") => {
                                                self.view_id =
                                                    a.unescape_value()?.parse::<u32>()?;
                                            }
                                            QName(b"tabSelected") => {
                                                if a.value == Cow::Borrowed(&[49]) {
                                                    self.show_tab = true;
                                                }
                                            }
                                            QName(b"showRulers") => {
                                                if a.value == Cow::Borrowed(&[49]) {
                                                    self.show_ruler = true;
                                                }
                                            }
                                            QName(b"showRowColHeaders") => {
                                                if a.value == Cow::Borrowed(&[49]) {
                                                    self.show_header = true;
                                                }
                                            }
                                            QName(b"rightToLeft") => {
                                                if a.value == Cow::Borrowed(&[49]) {
                                                    self.use_rtl = true;
                                                }
                                            }
                                            QName(b"showOutlineSymbols") => {
                                                if a.value == Cow::Borrowed(&[48]) {
                                                    self.show_outline_symbol = false;
                                                }
                                            }
                                            QName(b"colorId") => {
                                                self.default_grid_color = GridlineColor::from(
                                                    a.unescape_value()?.parse::<u8>()?,
                                                );
                                            }
                                            QName(b"showZeros") => {
                                                if a.value == Cow::Borrowed(&[48]) {
                                                    self.show_zero = false;
                                                }
                                            }
                                            QName(b"showFormulas") => {
                                                if a.value == Cow::Borrowed(&[49]) {
                                                    self.show_formula = true;
                                                }
                                            }
                                            _ => (),
                                        }
                                    }
                                }

                                if let Ok(Event::Start(_)) = event {
                                    //loop
                                }
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetViews" => {
                                break
                            }
                            Ok(Event::Eof) => {
                                return Err(XcelmateError::XmlEof("sheetViews".into()))
                            }
                            Err(e) => return Err(XcelmateError::Xml(e)),
                            _ => (),
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
    use super::Worksheet;
    use std::fs::File;
    use zip::ZipArchive;

    fn init(path: &str, sheet: &str) -> Worksheet {
        let file = File::open(path).unwrap();
        let mut zip = ZipArchive::new(file).unwrap();
        let mut worksheet = Worksheet::default();
        worksheet.read_sheet(&mut zip, sheet).unwrap();
        worksheet
    }

    mod sheet_api {
        use super::init;
        use crate::stream::xlsx::sheet::{GridlineColor, Worksheet};

        #[test]
        fn test_convert_dim_to_row_col() {
            let actual = Worksheet::convert_dim_to_row_col("B1").unwrap();
            assert_eq!(actual, (1, 0))
        }

        #[test]
        fn test_convert_dim_to_row_col_max_col() {
            let actual = Worksheet::convert_dim_to_row_col("XFE1")
                .err()
                .unwrap()
                .to_string();
            assert_eq!(actual, "excel columns can not exceed 16,384")
        }

        #[test]
        fn test_convert_dim_to_row_col_max_row() {
            let actual = Worksheet::convert_dim_to_row_col("XFD1048577")
                .err()
                .unwrap()
                .to_string();
            assert_eq!(actual, "excel rows can not exceed 1,048,576")
        }

        #[test]
        fn test_convert_dim_to_row_col_invalid_char() {
            let actual = Worksheet::convert_dim_to_row_col("*1048577")
                .err()
                .unwrap()
                .to_string();
            assert_eq!(actual, "can not parse excel dimension: *1048577")
        }

        #[test]
        fn test_convert_dim_to_row_col_not_within_a_to_z_range() {
            let actual = Worksheet::convert_dim_to_row_col("B京1048577")
                .err()
                .unwrap()
                .to_string();
            assert_eq!(actual, "can not parse excel dimension: B京1048577")
        }

        #[test]
        fn test_read_sheet() {
            let sheet = init("tests/workbook04.xlsx", "sheet2.xml");
            assert_eq!(sheet.name, "Sheet2");
        }
    }
}
