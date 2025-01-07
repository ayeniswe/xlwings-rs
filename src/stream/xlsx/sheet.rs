use super::{
    errors::XlsxError,
    stylesheet::{Color, Stylesheet},
    Xlsx,
};
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

/// Cell row
type Row = u32;
/// Cell Column
type Col = u16;

/// Max inclusive of cell columns allowed. Max letter column: `XFD`
const MAX_COLUMNS: u16 = 16_384;
/// Max inclusive of cell rows allowed
const MAX_ROWS: u32 = 1_048_576;

/// Predefined zoom level for a sheet
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) enum Zoom {
    Z25,
    Z50,
    Z75,
    #[default]
    Z100,
    Z200,
}
impl Into<Vec<u8>> for Zoom {
    fn into(self) -> Vec<u8> {
        match self {
            Zoom::Z25 => vec![b'2', b'5'],
            Zoom::Z50 => vec![b'5', b'0'],
            Zoom::Z75 => vec![b'7', b'5'],
            Zoom::Z100 => vec![b'1', b'0', b'0'],
            Zoom::Z200 => vec![b'2', b'0', b'0'],
        }
    }
}

/// The view setting of a sheet
#[derive(Default, Debug, Clone, PartialEq, Eq)]
pub(crate) enum View {
    #[default]
    /// Normal view
    Normal,
    /// Page break preview
    PageBreakPreview,
    /// Page layout view
    PageLayout,
}

#[derive(Default, Debug, Clone, FromPrimitive, IntoPrimitive, PartialEq, Eq)]
#[repr(u8)]
pub(crate) enum GridlineColor {
    #[default]
    Automatic = 0, // will reflect writing defaultGridColor instead of colorId
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

#[derive(Debug, Default, PartialEq, Clone, Eq)]
pub struct Sheet {
    path: String,
    name: Option<String>,
    tab_color: Option<Color>,
    fit_to_page: bool,
    auto_page_break: bool,
    dimensions: ((Col, Row), (Col, Row)),
    show_grid: bool,
    zoom_scale: Option<Vec<u8>>, // writes zoomScale and zoomNormalScale with same value
    zoom_scale_page: Option<Vec<u8>>,
    zoom_scale_sheet: Option<Vec<u8>>,
    view_id: Option<Vec<u8>>,
    show_tab: bool,
    show_ruler: bool,
    show_header: bool,
    show_outline_symbol: bool,
    default_grid_color: GridlineColor,
    show_formula: bool,
    show_zero: bool,
    show_whitespace: bool,
    use_protection: bool,
    use_rtl: bool,
    top_left_cell: Option<(Col, Row)>,
    view: Option<View>,
    enable_cond_format_calc: bool,
    published: bool,
    sync_vertical: bool,
    sync_horizontal: bool,
    sync_ref: Option<String>,
    transition_eval: bool,
    transition_entry: bool,
    filter_mode: bool,
    apply_outline_style: bool,
    show_summary_below: bool, // summary row should be inserted to above when off
    show_summary_right: bool, // sumamry row should be inserted to left when off
}
impl Sheet {
    fn new(path: &str) -> Self {
        Self {
            path: path.into(),
            show_grid: true,
            show_zero: true,
            apply_outline_style: false,
            show_summary_below: true,
            show_summary_right: true,
            show_outline_symbol: true,
            view: None,
            default_grid_color: GridlineColor::Automatic,
            name: None,
            tab_color: None,
            fit_to_page: false,
            auto_page_break: true,
            dimensions: ((0, 0), (0, 0)),
            zoom_scale: None,
            zoom_scale_page: None,
            zoom_scale_sheet: None,
            view_id: None,
            show_tab: false,
            show_ruler: false,
            show_header: false,
            show_formula: false,
            show_whitespace: false,
            use_protection: false,
            use_rtl: false,
            top_left_cell: None,
            enable_cond_format_calc: true,
            published: true,
            sync_vertical: false,
            sync_horizontal: false,
            sync_ref: None,
            transition_eval: false,
            transition_entry: false,
            filter_mode: false,
        }
    }
    pub fn read_properties<'a, RS: Read + Seek>(
        &mut self,
        zip: &'a mut ZipArchive<RS>,
    ) -> Result<(), XlsxError> {
        let mut xml = match xml_reader(zip, &self.path, None) {
            None => return Err(XlsxError::SheetNotFound(self.path.clone())),
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
                            match a.key.as_ref() {
                                b"codeName" => self.name = Some(a.unescape_value()?.to_string()),
                                b"enableFormatConditions" => {
                                    self.enable_cond_format_calc = *a.value == *b"1"
                                }
                                _ => (),
                            }
                        }
                    }
                }
                ////////////////////
                // SHEET PROPERTIES nth-1
                /////////////
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"tabColor" => {
                    self.tab_color = Some(Stylesheet::read_color(e.attributes())?);
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"pageSetUpPr" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key.as_ref() {
                                b"fitToPage" => self.fit_to_page = *a.value != *b"0",
                                b"autoPageBreaks" => self.auto_page_break = *a.value == *b"1",
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"dimension" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key.as_ref() {
                                b"ref" => {
                                    // Split dimension range for top left cell and bottom right cell
                                    let mut dimensions = a.value.split(|b| b == &b':');

                                    let m =
                                        Sheet::convert_dim_to_row_col(dimensions.next().unwrap())?;
                                    if let Some(x) = dimensions.next() {
                                        let x = Sheet::convert_dim_to_row_col(x)?;
                                        self.dimensions = (m, x);
                                    } else {
                                        self.dimensions = (m, m);
                                    };
                                }
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetPr" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetPr".into())),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
        Ok(())
    }
    pub fn read_sheet<'a, RS: Read + Seek>(
        &mut self,
        zip: &'a mut ZipArchive<RS>,
    ) -> Result<(), XlsxError> {
        let mut xml = match xml_reader(zip, &self.path, None) {
            None => return Err(XlsxError::SheetNotFound(self.path.clone())),
            Some(x) => x?,
        };

        let mut buf = Vec::with_capacity(1024);
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                ////////////////////
                // SHEET VIEW
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheetViews" => {
                    let mut view_buf = Vec::with_capacity(1024);
                    loop {
                        view_buf.clear();
                        let event = xml.read_event_into(&mut view_buf);
                        match event {
                            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e))
                                if e.local_name().as_ref() == b"sheetView" =>
                            {
                                ////////////////////
                                // SHEET VIEW Attrs
                                /////////////
                                for attr in e.attributes() {
                                    if let Ok(a) = attr {
                                        match a.key.as_ref() {
                                            b"showGridLines" => {
                                                self.show_grid = *a.value != *b"0";
                                            }
                                            b"zoomScaleNormal" | b"zoomScale" => {
                                                self.zoom_scale = Some(a.value.into_owned());
                                            }
                                            b"zoomScalePageLayoutView" => {
                                                self.zoom_scale_page = Some(a.value.into_owned());
                                            }
                                            b"zoomScaleSheetLayoutView" => {
                                                self.zoom_scale_sheet = Some(a.value.into_owned());
                                            }
                                            b"workbookViewId" => {
                                                self.view_id = Some(a.value.into_owned())
                                            }
                                            b"windowProtection" => {
                                                self.use_protection = *a.value == *b"1";
                                            }
                                            b"tabSelected" => {
                                                self.show_tab = *a.value == *b"1";
                                            }
                                            b"showRuler" => {
                                                self.show_ruler = *a.value == *b"1";
                                            }
                                            b"showRowColHeaders" => {
                                                self.show_header = *a.value == *b"1";
                                            }
                                            b"rightToLeft" => {
                                                self.use_rtl = *a.value == *b"1";
                                            }
                                            b"showOutlineSymbols" => {
                                                self.show_outline_symbol = *a.value != *b"0"
                                            }
                                            b"showWhiteSpace" => {
                                                self.show_whitespace = *a.value == *b"1";
                                            }
                                            b"view" => match a.value.as_ref() {
                                                b"pageBreakPreview" => {
                                                    self.view = Some(View::PageBreakPreview)
                                                }
                                                b"pageLayout" => self.view = Some(View::PageLayout),
                                                b"normal" => self.view = Some(View::Normal),
                                                _ => (),
                                            },
                                            b"topLeftCell" => {
                                                self.top_left_cell =
                                                    Some(Sheet::convert_dim_to_row_col(&a.value)?);
                                            }
                                            b"colorId" => {
                                                match a.value.as_ref() {
                                                    b"0" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Automatic
                                                    }
                                                    b"8" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Black
                                                    }
                                                    b"15" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Turquoise
                                                    }
                                                    b"60" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Brown
                                                    }
                                                    b"14" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Pink
                                                    }
                                                    b"59" => {
                                                        self.default_grid_color =
                                                            GridlineColor::OliveGreen
                                                    }
                                                    b"58" => {
                                                        self.default_grid_color =
                                                            GridlineColor::DarkGreen
                                                    }
                                                    b"56" => {
                                                        self.default_grid_color =
                                                            GridlineColor::DarkTeal
                                                    }
                                                    b"18" => {
                                                        self.default_grid_color =
                                                            GridlineColor::DarkBlue
                                                    }
                                                    b"62" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Indigo
                                                    }
                                                    b"63" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Gray80
                                                    }
                                                    b"23" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Gray50
                                                    }
                                                    b"55" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Gray40
                                                    }
                                                    b"22" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Gray25
                                                    }
                                                    b"9" => {
                                                        self.default_grid_color =
                                                            GridlineColor::White
                                                    }
                                                    b"31" => {
                                                        self.default_grid_color =
                                                            GridlineColor::IceBlue
                                                    }
                                                    b"12" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Blue
                                                    }
                                                    b"21" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Teal
                                                    }
                                                    b"30" => {
                                                        self.default_grid_color =
                                                            GridlineColor::OceanBlue
                                                    }
                                                    b"25" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Plum
                                                    }
                                                    b"46" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Lavender
                                                    }
                                                    b"20" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Violet
                                                    }
                                                    b"54" => {
                                                        self.default_grid_color =
                                                            GridlineColor::BlueGray
                                                    }
                                                    b"48" => {
                                                        self.default_grid_color =
                                                            GridlineColor::LightBlue
                                                    }
                                                    b"40" => {
                                                        self.default_grid_color =
                                                            GridlineColor::SkyBlue
                                                    }
                                                    b"44" => {
                                                        self.default_grid_color =
                                                            GridlineColor::PaleBlue
                                                    }
                                                    b"29" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Coral
                                                    }
                                                    b"16" => {
                                                        self.default_grid_color =
                                                            GridlineColor::DarkRed
                                                    }
                                                    b"49" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Aqua
                                                    }
                                                    b"27" => {
                                                        self.default_grid_color =
                                                            GridlineColor::LightTurquoise
                                                    }
                                                    b"28" => {
                                                        self.default_grid_color =
                                                            GridlineColor::DarkPurple
                                                    }
                                                    b"57" => {
                                                        self.default_grid_color =
                                                            GridlineColor::SeaGreen
                                                    }
                                                    b"42" => {
                                                        self.default_grid_color =
                                                            GridlineColor::LightGreen
                                                    }
                                                    b"11" => {
                                                        self.default_grid_color =
                                                            GridlineColor::BrightGreen
                                                    }
                                                    b"13" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Yellow
                                                    }
                                                    b"26" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Ivory
                                                    }
                                                    b"43" => {
                                                        self.default_grid_color =
                                                            GridlineColor::LightYellow
                                                    }
                                                    b"19" => {
                                                        self.default_grid_color =
                                                            GridlineColor::DarkYellow
                                                    }
                                                    b"50" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Lime
                                                    }
                                                    b"53" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Orange
                                                    }
                                                    b"52" => {
                                                        self.default_grid_color =
                                                            GridlineColor::LightOrange
                                                    }
                                                    b"51" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Gold
                                                    }
                                                    b"47" => {
                                                        self.default_grid_color = GridlineColor::Tan
                                                    }
                                                    b"45" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Rose
                                                    }
                                                    b"24" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Periwinkle
                                                    }
                                                    b"10" => {
                                                        self.default_grid_color = GridlineColor::Red
                                                    }
                                                    b"17" => {
                                                        self.default_grid_color =
                                                            GridlineColor::Green
                                                    }
                                                    _ => (),
                                                };
                                            }

                                            b"showZeros" => self.show_zero = *a.value != *b"0",
                                            b"showFormulas" => {
                                                self.show_formula = *a.value == *b"1"
                                            }
                                            _ => (),
                                        }
                                    }
                                }

                                // if let Ok(Event::Start(_)) = event {
                                //     //loop
                                // }
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetViews" => {
                                break
                            }
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetViews".into())),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"worksheet" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("worksheet".into())),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
        Ok(())
    }

    fn convert_dim_to_row_col(dimension: &[u8]) -> Result<(Col, Row), XlsxError> {
        let mut col: Vec<u8> = Vec::with_capacity(3);
        let mut row: Vec<u8> = Vec::with_capacity(7);

        for c in dimension.iter() {
            if c.is_ascii_alphabetic() {
                col.push(*c)
            } else if c.is_ascii_digit() {
                row.push(*c)
            } else {
                let mut buf = String::with_capacity(11);
                let _ = dimension.as_ref().read_to_string(&mut buf)?;
                return Err(XlsxError::ExcelDimensionParseError(buf));
            }
        }

        // Calculate column to integer
        let mut result = 0;
        for c in col.iter() {
            if c < &b'A' || c > &b'Z' {
                let mut buf = String::with_capacity(11);
                let _ = dimension.as_ref().read_to_string(&mut buf)?;
                return Err(XlsxError::ExcelDimensionParseError(buf));
            }
            let value = *c as i32 - 'A' as i32 + 1;
            result = result * 26 + value;
        }

        let row = row
            .iter()
            .fold(0, |acc, &byte| acc * 10 + (&byte - b'0') as u32);
        let col = result as u16;
        if row > MAX_ROWS {
            return Err(XlsxError::ExcelMaxRowExceeded);
        } else if col > MAX_COLUMNS {
            return Err(XlsxError::ExcelMaxColumnExceeded);
        } else {
            Ok((col - 1, row - 1))
        }
    }
}

mod sheet_unittests {
    use super::Sheet;
    use std::fs::File;
    use zip::ZipArchive;

    fn init(path: &str, sheet_path: &str) -> Sheet {
        let file = File::open(path).unwrap();
        let mut zip = ZipArchive::new(file).unwrap();
        let mut sheet = Sheet::new(sheet_path);
        sheet.read_sheet(&mut zip).unwrap();
        sheet
    }

    mod sheet_api {
        use std::{
            fs::File,
            io::{Seek, SeekFrom},
        };

        use super::init;
        use crate::stream::xlsx::sheet::{Color, GridlineColor, Sheet};

        #[test]
        fn test_convert_dim_to_row_col() {
            let actual = Sheet::convert_dim_to_row_col(&"B1".as_bytes().to_vec()).unwrap();
            assert_eq!(actual, (1, 0))
        }

        #[test]
        fn test_convert_dim_to_row_col_max_col() {
            let actual = Sheet::convert_dim_to_row_col(&"XFE1".as_bytes().to_vec())
                .err()
                .unwrap()
                .to_string();
            assert_eq!(actual, "excel columns can not exceed 16,384")
        }

        #[test]
        fn test_convert_dim_to_row_col_at_max_col() {
            let actual = Sheet::convert_dim_to_row_col(&"XFD1".as_bytes().to_vec()).unwrap();

            assert_eq!(actual, (16_383, 0))
        }

        #[test]
        fn test_convert_dim_to_row_col_at_max_row() {
            let actual = Sheet::convert_dim_to_row_col(&"XFD1048576".as_bytes().to_vec()).unwrap();

            assert_eq!(actual, (16_383, 1_048_575))
        }

        #[test]
        fn test_convert_dim_to_row_col_max_row() {
            let actual = Sheet::convert_dim_to_row_col(&"XFD1048577".as_bytes().to_vec())
                .err()
                .unwrap()
                .to_string();
            assert_eq!(actual, "excel rows can not exceed 1,048,576")
        }

        #[test]
        fn test_convert_dim_to_row_col_invalid_char() {
            let actual = Sheet::convert_dim_to_row_col(&"*1048577".as_bytes().to_vec())
                .err()
                .unwrap()
                .to_string();
            assert_eq!(actual, "can not parse excel dimension: *1048577")
        }

        #[test]
        fn test_convert_dim_to_row_col_not_within_a_to_z_range() {
            let actual = Sheet::convert_dim_to_row_col(&"B京1048577".as_bytes().to_vec())
                .err()
                .unwrap()
                .to_string();
            assert_eq!(actual, "can not parse excel dimension: B京1048577")
        }

        #[test]
        fn test_read_sheet() {
            let sheet = init("tests/workbook04.xlsx", "xl/worksheets/sheet2.xml");
            assert_eq!(
                sheet,
                Sheet {
                    path: "xl/worksheets/sheet2.xml".into(),
                    name: Some("Sheet2".into()),
                    tab_color: Some(Color::Theme { id: 5, tint: None }),
                    fit_to_page: true,
                    dimensions: ((1, 0), (6, 13)),
                    show_grid: false,
                    zoom_scale: Some("100".as_bytes().to_vec()),
                    view_id: Some("0".as_bytes().to_vec()),
                    ..Default::default()
                }
            );
        }
    }
}
