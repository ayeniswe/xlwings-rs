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
    fs::read_to_string,
    io::{BufRead, Cursor, Read, Seek, SeekFrom, Write},
    ops::RangeInclusive,
    sync::Arc,
};
use zip::{
    read::ZipFileSeek,
    write::{FileOptionExtension, FileOptions},
    ZipArchive,
};

/// Cell row
type Row = u32;
/// Cell Column
type Col = u16;
type Cell = (Col, Row);
type CellRange = ((Col, Row), (Col, Row));

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

/// The pane for viewing worksheet
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) struct Pane {
    active_pane: PanePosition,
    state: PaneState,
    top_left_cell: Option<Cell>,
    x_split: Vec<u8>,
    y_split: Vec<u8>,
}
impl Pane {
    pub fn new() -> Self {
        Self {
            x_split: b"0".to_vec(),
            y_split: b"0".to_vec(),
            ..Default::default()
        }
    }
}
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) enum PanePosition {
    BottomRight,
    TopRight,
    BottomLeft,
    #[default]
    TopLeft,
}
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) enum PaneState {
    Frozen,
    #[default]
    Split,
    FrozenSplit,
}

/// A sheet view selection view
#[derive(Debug, Default, PartialEq, Clone, Eq)]
pub(crate) struct Selection {
    cell: Option<Cell>,
    cell_id: Vec<u8>,
    pane: PanePosition,
    sqref: Vec<CellRange>,
}
impl Selection {
    pub(crate) fn new() -> Self {
        Self {
            sqref: vec![((0, 0), (0, 0))],
            cell_id: b"0".to_vec(),
            ..Default::default()
        }
    }
}
/// A default view of a sheet
#[derive(Debug, Default, PartialEq, Clone, Eq)]
pub(crate) struct SheetView {
    zoom_scale: Option<Vec<u8>>, // writes zoomScale and zoomNormalScale with same value
    zoom_scale_page: Option<Vec<u8>>,
    zoom_scale_sheet: Option<Vec<u8>>,
    view_id: Option<Vec<u8>>,
    top_left_cell: Option<Cell>,
    view: Option<View>,
    grid_color: GridlineColor,
    show_formula: bool,
    show_zero: bool,
    show_grid: bool,
    show_whitespace: bool,
    use_protection: bool,
    show_tab: bool,
    show_ruler: bool,
    use_rtl: bool,
    pane: Option<Pane>,
    selection: Option<Selection>,
    show_header: bool,
    show_outline_symbol: bool,
}
impl SheetView {
    fn new() -> Self {
        Self {
            show_grid: true,
            show_zero: true,
            show_outline_symbol: true,
            view: None,
            grid_color: GridlineColor::Automatic,
            pane: None,
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
            ..Default::default()
        }
    }
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
    uid: Vec<u8>,
    path: String,
    code_name: Option<Vec<u8>>,
    fit_to_page: bool,
    auto_page_break: bool,
    dimensions: CellRange,
    enable_cond_format_calc: bool,
    published: bool,
    sync_vertical: bool,
    sync_horizontal: bool,
    sync_ref: Option<(Cell, Cell)>,
    transition_eval: bool,
    transition_entry: bool,
    filter_mode: bool,
    apply_outline_style: bool,
    show_summary_below: bool, // summary row should be inserted to above when off
    show_summary_right: bool, // sumamry row should be inserted to left when off
    sheet_views: Vec<SheetView>,
    tab_color: Option<Color>,
    show_outline_symbol: bool,
}

impl<W: Write> XmlWriter<W> for Sheet {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &'a str,
    ) -> Result<&'a mut Writer<W>, XlsxError> {
        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;

        // worksheet
        writer
            .create_element(tag_name)
            .with_attributes(vec![
                (
                    "xmlns",
                    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                ),
                (
                    "xmlns:r",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                ),
                (
                    "xmlns:mc",
                    "http://schemas.openxmlformats.org/markup-compatibility/2006",
                ),
                ("mc:Ignorable", "x14ac xr xr2 xr3"),
                (
                    "xmlns:x14ac",
                    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
                ),
                (
                    "xmlns:xr2",
                    "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
                ),
                (
                    "xmlns:xr3",
                    "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
                ),
                (
                    "xmlns:xr",
                    "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
                ),
                ("xr:uid", &String::from_utf8(self.uid.clone())?),
            ])
            .write_inner_content::<_, XlsxError>(|writer| {
                // sheetPr
                let mut attrs = Vec::with_capacity(9);
                if let Some(ref code_name) = self.code_name {
                    attrs.push((b"codeName".as_ref(), code_name.as_ref()));
                }
                if !self.enable_cond_format_calc {
                    attrs.push((b"enableFormatConditions".as_ref(), b"0".as_ref()));
                }
                if self.filter_mode {
                    attrs.push((b"filterMode".as_ref(), b"1".as_ref()));
                }
                if !self.published {
                    attrs.push((b"published".as_ref(), b"0".as_ref()));
                }
                if self.transition_entry {
                    attrs.push((b"transitionEntry".as_ref(), b"1".as_ref()));
                }
                if self.transition_eval {
                    attrs.push((b"transitionEvaluation".as_ref(), b"1".as_ref()));
                }
                let sync_ref_bytes;
                if let Some(ref sync_ref) = self.sync_ref {
                    sync_ref_bytes = Sheet::cell_range_to_cell_reference(sync_ref);
                    attrs.push((b"syncRef".as_ref(), sync_ref_bytes.as_ref()));
                }
                if self.sync_vertical {
                    attrs.push((b"syncVertical".as_ref(), b"1".as_ref()));
                }
                if self.sync_horizontal {
                    attrs.push((b"syncHorizontal".as_ref(), b"1".as_ref()));
                }
                writer
                    .create_element("sheetPr")
                    .with_attributes(attrs)
                    .write_inner_content::<_, XlsxError>(|writer| {
                        // tabColor
                        if let Some(color) = &self.tab_color {
                            color.write_xml(writer, "tabColor")?;
                        }
                        // pageSetUpPr
                        if self.fit_to_page || self.auto_page_break {
                            let mut attrs = Vec::with_capacity(2);
                            if self.fit_to_page {
                                attrs.push((b"fitToPage".as_ref(), b"1".as_ref()));
                            }
                            if !self.auto_page_break {
                                attrs.push((b"autoPageBreaks".as_ref(), b"0".as_ref()));
                            }
                            writer
                                .create_element("pageSetUpPr")
                                .with_attributes(attrs)
                                .write_empty()?;
                        }
                        Ok(())
                    })?;

                Ok(())
            })?;
        Ok(writer)
    }
}
impl Sheet {
    fn new(path: &str) -> Self {
        Self {
            uid: vec![],
            tab_color: None,
            path: path.into(),
            apply_outline_style: false,
            show_summary_below: true,
            show_summary_right: true,
            fit_to_page: false,
            auto_page_break: true,
            dimensions: ((0, 0), (0, 0)),
            enable_cond_format_calc: true,
            published: true,
            code_name: None,
            sheet_views: Vec::new(),
            sync_vertical: false,
            sync_horizontal: false,
            sync_ref: None,
            transition_eval: false,
            transition_entry: false,
            filter_mode: false,
            show_outline_symbol: true,
        }
    }

    /// Read all of the `sheetPr`, `dimension` sections
    fn read_properties<B: BufRead>(&mut self, xml: &mut Reader<B>) -> Result<(), XlsxError> {
        let mut buf = Vec::with_capacity(1024);
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"worksheet" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key.as_ref() {
                                b"xr:uid" => self.uid = a.value.into(),
                                _ => (),
                            }
                        }
                    }
                }
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
                                b"codeName" => self.code_name = Some(a.value.into()),
                                b"enableFormatConditions" => {
                                    self.enable_cond_format_calc = *a.value == *b"1"
                                }
                                b"filterMode" => self.filter_mode = *a.value == *b"1",
                                b"published" => self.published = *a.value == *b"1",
                                b"syncHorizontal" => self.sync_horizontal = *a.value == *b"1",
                                b"syncVertical" => self.sync_vertical = *a.value == *b"1",
                                b"syncRef" => {
                                    self.sync_ref =
                                        Some(Sheet::cell_reference_to_cell_range(&a.value)?)
                                }
                                b"transitionEvaluation" => self.transition_eval = *a.value == *b"1",
                                b"transitionEntry" => self.transition_entry = *a.value == *b"1",
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
                                    self.dimensions = Sheet::cell_reference_to_cell_range(&a.value)?
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
                                b"fitToPage" => self.fit_to_page = *a.value == *b"1",
                                b"autoPageBreaks" => self.auto_page_break = *a.value == *b"1",
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"outlinePr" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key.as_ref() {
                                b"applyStyles" => self.apply_outline_style = *a.value == *b"1",
                                b"summaryBelow" => self.show_summary_below = *a.value == *b"1",
                                b"summaryRight" => self.show_summary_right = *a.value == *b"1",
                                b"showOutlineSymbols" => {
                                    self.show_outline_symbol = *a.value == *b"1"
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

    /// Read all of the `sheetViews` section
    pub fn read_sheet_views<B: BufRead>(&mut self, xml: &mut Reader<B>) -> Result<(), XlsxError> {
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
                        let mut sheet_view = SheetView::new();
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
                                                sheet_view.show_grid = *a.value == *b"1";
                                            }
                                            b"zoomScaleNormal" | b"zoomScale" => {
                                                sheet_view.zoom_scale = Some(a.value.into());
                                            }
                                            b"zoomScalePageLayoutView" => {
                                                sheet_view.zoom_scale_page = Some(a.value.into());
                                            }
                                            b"zoomScaleSheetLayoutView" => {
                                                sheet_view.zoom_scale_sheet = Some(a.value.into());
                                            }
                                            b"workbookViewId" => {
                                                sheet_view.view_id = Some(a.value.into())
                                            }
                                            b"windowProtection" => {
                                                sheet_view.use_protection = *a.value == *b"1";
                                            }
                                            b"tabSelected" => {
                                                sheet_view.show_tab = *a.value == *b"1";
                                            }
                                            b"showRuler" => {
                                                sheet_view.show_ruler = *a.value == *b"1";
                                            }
                                            b"showRowColHeaders" => {
                                                sheet_view.show_header = *a.value == *b"1";
                                            }
                                            b"rightToLeft" => {
                                                sheet_view.use_rtl = *a.value == *b"1";
                                            }
                                            b"showOutlineSymbols" => {
                                                sheet_view.show_outline_symbol = *a.value == *b"1"
                                            }
                                            b"showWhiteSpace" => {
                                                sheet_view.show_whitespace = *a.value == *b"1";
                                            }
                                            b"view" => match a.value.as_ref() {
                                                b"pageBreakPreview" => {
                                                    sheet_view.view = Some(View::PageBreakPreview)
                                                }
                                                b"pageLayout" => {
                                                    sheet_view.view = Some(View::PageLayout)
                                                }
                                                b"normal" => sheet_view.view = Some(View::Normal),
                                                _ => (),
                                            },
                                            b"topLeftCell" => {
                                                sheet_view.top_left_cell =
                                                    Some(Sheet::cell_reference_to_cell(&a.value)?);
                                            }
                                            b"colorId" => {
                                                match a.value.as_ref() {
                                                    b"0" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::Automatic
                                                    }
                                                    b"8" => {
                                                        sheet_view.grid_color = GridlineColor::Black
                                                    }
                                                    b"15" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::Turquoise
                                                    }
                                                    b"60" => {
                                                        sheet_view.grid_color = GridlineColor::Brown
                                                    }
                                                    b"14" => {
                                                        sheet_view.grid_color = GridlineColor::Pink
                                                    }
                                                    b"59" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::OliveGreen
                                                    }
                                                    b"58" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::DarkGreen
                                                    }
                                                    b"56" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::DarkTeal
                                                    }
                                                    b"18" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::DarkBlue
                                                    }
                                                    b"62" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::Indigo
                                                    }
                                                    b"63" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::Gray80
                                                    }
                                                    b"23" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::Gray50
                                                    }
                                                    b"55" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::Gray40
                                                    }
                                                    b"22" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::Gray25
                                                    }
                                                    b"9" => {
                                                        sheet_view.grid_color = GridlineColor::White
                                                    }
                                                    b"31" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::IceBlue
                                                    }
                                                    b"12" => {
                                                        sheet_view.grid_color = GridlineColor::Blue
                                                    }
                                                    b"21" => {
                                                        sheet_view.grid_color = GridlineColor::Teal
                                                    }
                                                    b"30" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::OceanBlue
                                                    }
                                                    b"25" => {
                                                        sheet_view.grid_color = GridlineColor::Plum
                                                    }
                                                    b"46" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::Lavender
                                                    }
                                                    b"20" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::Violet
                                                    }
                                                    b"54" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::BlueGray
                                                    }
                                                    b"48" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::LightBlue
                                                    }
                                                    b"40" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::SkyBlue
                                                    }
                                                    b"44" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::PaleBlue
                                                    }
                                                    b"29" => {
                                                        sheet_view.grid_color = GridlineColor::Coral
                                                    }
                                                    b"16" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::DarkRed
                                                    }
                                                    b"49" => {
                                                        sheet_view.grid_color = GridlineColor::Aqua
                                                    }
                                                    b"27" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::LightTurquoise
                                                    }
                                                    b"28" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::DarkPurple
                                                    }
                                                    b"57" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::SeaGreen
                                                    }
                                                    b"42" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::LightGreen
                                                    }
                                                    b"11" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::BrightGreen
                                                    }
                                                    b"13" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::Yellow
                                                    }
                                                    b"26" => {
                                                        sheet_view.grid_color = GridlineColor::Ivory
                                                    }
                                                    b"43" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::LightYellow
                                                    }
                                                    b"19" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::DarkYellow
                                                    }
                                                    b"50" => {
                                                        sheet_view.grid_color = GridlineColor::Lime
                                                    }
                                                    b"53" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::Orange
                                                    }
                                                    b"52" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::LightOrange
                                                    }
                                                    b"51" => {
                                                        sheet_view.grid_color = GridlineColor::Gold
                                                    }
                                                    b"47" => {
                                                        sheet_view.grid_color = GridlineColor::Tan
                                                    }
                                                    b"45" => {
                                                        sheet_view.grid_color = GridlineColor::Rose
                                                    }
                                                    b"24" => {
                                                        sheet_view.grid_color =
                                                            GridlineColor::Periwinkle
                                                    }
                                                    b"10" => {
                                                        sheet_view.grid_color = GridlineColor::Red
                                                    }
                                                    b"17" => {
                                                        sheet_view.grid_color = GridlineColor::Green
                                                    }
                                                    _ => (),
                                                };
                                            }

                                            b"showZeros" => {
                                                sheet_view.show_zero = *a.value == *b"1"
                                            }
                                            b"showFormulas" => {
                                                sheet_view.show_formula = *a.value == *b"1"
                                            }
                                            _ => (),
                                        }
                                    }
                                }
                                ////////////////////
                                // SHEET VIEW Attrs nth-1
                                /////////////
                                if let Ok(Event::Start(_)) = event {
                                    let mut view_value_buf = Vec::with_capacity(1024);
                                    loop {
                                        view_value_buf.clear();
                                        match xml.read_event_into(&mut view_value_buf) {
                                            Ok(Event::Empty(ref e))
                                            ////////////////////
                                            // Pane
                                            /////////////
                                            if e.local_name().as_ref() == b"pane" =>
                                            {
                                                let mut pane = Pane::new();
                                                for attr in e.attributes() {
                                                    if let Ok(a) = attr {
                                                        match a.key.as_ref() {
                                                            b"activePane" => {
                                                                match a.value.as_ref() {
                                                                    b"bottomLeft" => {
                                                                        pane.active_pane =
                                                                            PanePosition::BottomLeft
                                                                    }
                                                                    b"bottomRight" => pane.active_pane =
                                                                        PanePosition::BottomRight,
                                                                    b"topLeft" => {
                                                                        pane.active_pane =
                                                                            PanePosition::TopLeft
                                                                    }
                                                                    b"topRight" => {
                                                                        pane.active_pane =
                                                                            PanePosition::TopRight
                                                                    }
                                                                    _ => (),
                                                                }
                                                            }
                                                            b"state" => match a.value.as_ref() {
                                                                b"frozen" => {
                                                                    pane.state = PaneState::Frozen
                                                                }
                                                                b"split" => {
                                                                    pane.state = PaneState::Split
                                                                }
                                                                b"frozenSplit" => {
                                                                    pane.state =
                                                                        PaneState::FrozenSplit
                                                                }
                                                                _ => (),
                                                            },
                                                            b"topLeftCell" => {
                                                                pane.top_left_cell = Some(
                                                                    Sheet::cell_reference_to_cell(
                                                                        &a.value,
                                                                    )?,
                                                                )
                                                            }
                                                            b"xSplit" => {
                                                                pane.x_split = a.value.to_vec()
                                                            }
                                                            b"ySplit" => {
                                                                pane.y_split = a.value.to_vec()
                                                            }
                                                            _ => (),
                                                        }
                                                    }
                                                }
                                                sheet_view.pane = Some(pane)
                                            }
                                                                                        ////////////////////
                                            // Pane
                                            /////////////
                                            Ok(Event::Start(ref e))
                                                if e.local_name().as_ref() == b"selection" =>
                                            {
                                                let mut selection = Selection::new();
                                                for attr in e.attributes() {
                                                    if let Ok(a) = attr {
                                                        match a.key.as_ref() {
                                                            b"pane" => {
                                                                match a.value.as_ref() {
                                                                    b"bottomLeft" => {
                                                                        selection.pane =
                                                                            PanePosition::BottomLeft
                                                                    }
                                                                    b"bottomRight" => selection.pane =
                                                                        PanePosition::BottomRight,
                                                                    b"topLeft" => {
                                                                        selection.pane =
                                                                            PanePosition::TopLeft
                                                                    }
                                                                    b"topRight" => {
                                                                        selection.pane =
                                                                            PanePosition::TopRight
                                                                    }
                                                                    _ => (),
                                                                }
                                                            }
                                                            b"activeCellId" => {
                                                                selection.cell_id =
                                                                    a.value.into();
                                                            }
                                                            b"activeCell" => {
                                                                selection.cell =
                                                                    Some(Sheet::cell_reference_to_cell(&a.value)?);
                                                            }
                                                            b"sqref" => {
                                                                selection.sqref = Sheet::cell_group_to_list_cell_range(&a.value)?;
                                                            }

                                                            _ => (),
                                                        }
                                                    }
                                                }
                                                sheet_view.selection = Some(selection)
                                            }
                                            Ok(Event::End(ref e))
                                                if e.local_name().as_ref() == b"sheetView" =>
                                            {
                                                break
                                            }
                                            Ok(Event::Eof) => {
                                                return Err(XlsxError::XmlEof("sheetView".into()))
                                            }
                                            Err(e) => {
                                                return Err(XlsxError::Xml(e));
                                            }
                                            _ => (),
                                        }
                                    }
                                }
                                self.sheet_views.push(sheet_view);
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetViews" => {
                                break
                            }
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetViews".into())),
                            Err(e) => {
                                return Err(XlsxError::Xml(e));
                            }
                            _ => (),
                        }
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"worksheet" => break,
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("worksheet".into())),
                Err(e) => {
                    return Err(XlsxError::Xml(e));
                }
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
        let _ = self.read_properties(&mut xml);
        // let _ = self.read_sheet_views(&mut xml);
        Ok(())
    }

    /// Convert a `ST_CellRef` cell dimensions to `CellRange`
    fn cell_reference_to_cell_range(cell_ref: &[u8]) -> Result<CellRange, XlsxError> {
        // Split dimension range for top left cell and bottom right cell
        let mut dimensions = cell_ref.split(|b| b == &b':');
        let m = Sheet::cell_reference_to_cell(dimensions.next().unwrap())?;
        if let Some(x) = dimensions.next() {
            let x = Sheet::cell_reference_to_cell(x)?;
            Ok((m, x))
        } else {
            Ok((m, m))
        }
    }

    /// Convert a `ST_CellRef` cell-column or cell-row to a tuple to easily represent a `Cell`
    fn cell_reference_to_cell(dimension: &[u8]) -> Result<Cell, XlsxError> {
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

    /// Convert a `Cell` to a `ST_Ref` cell-column or cell-row
    fn cell_to_cell_reference(dimension: Cell) -> Vec<u8> {
        let mut row = Vec::new();
        let mut temp = dimension.1 + 1;
        while temp > 0 {
            let digit = (temp % 10) as u8 + b'0';
            row.insert(0, digit);
            temp /= 10;
        }

        let mut col = Vec::new();
        let mut number = dimension.0;
        // Edge case for zero indexed
        if number == 0 {
            col.push(b'A');
        } else {
            while number > 0 {
                let remainder = (number % 26) as u8;
                let value = remainder + b'A';
                col.insert(0, value);
                number /= 26;
            }
        }

        col.extend(row.iter());
        col
    }

    /// Convert a `CellRange` to a `ST_Ref` cell-range
    fn cell_range_to_cell_reference(dimension: &CellRange) -> Vec<u8> {
        let mut top_left = Sheet::cell_to_cell_reference(dimension.0);
        top_left.push(b':');
        let bottom_right = Sheet::cell_to_cell_reference(dimension.1);
        top_left.extend(bottom_right.iter());
        top_left
    }

    /// Convert `Vec<CellRange>` to a group of `ST_CellRef`
    fn list_cell_range_to_cell_group(ranges: &Vec<CellRange>) -> Vec<u8> {
        let mut group = Vec::new();
        for range in ranges {
            group.extend(Sheet::cell_range_to_cell_reference(range));
            group.push(b' ');
        }
        // Remove last added space
        group.pop();
        group
    }

    /// Convert a group of `ST_CellRef` to `Vec<CellRange>`
    fn cell_group_to_list_cell_range(value: &[u8]) -> Result<Vec<CellRange>, XlsxError> {
        // Sepearate by space delimeter to get groups of single ref, or range ref
        // all ref will be auto converted to ranges for ease of one type. Example A1 becomes A1:A1
        let groups = value
            .split(|b| *b == b' ')
            .map(|g| g.to_vec())
            .collect::<Vec<Vec<u8>>>();
        let mut cell_ranges = Vec::new();
        for cell_ref in groups {
            cell_ranges.push(Sheet::cell_reference_to_cell_range(&cell_ref)?);
        }
        Ok(cell_ranges)
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
        fn test_cell_group_to_list_cell_range() {
            let actual = Sheet::cell_group_to_list_cell_range(&"A1:B2 D4:D5 F10".as_bytes().to_vec()).unwrap();
            assert_eq!(actual[0], ((0, 0), (1,1)));
            assert_eq!(actual[1], ((3, 3), (3,4)));
            assert_eq!(actual[2], ((5, 9), (5,9)));
        }

        #[test]
        fn test_list_cell_range_to_cell_group() {
            let actual = Sheet::cell_group_to_list_cell_range(&"A1:B2 D4:D5 F10".as_bytes().to_vec()).unwrap();
            let actual = Sheet::list_cell_range_to_cell_group(&actual);
            assert_eq!(actual, "A1:B2 D4:D5 F10:F10".as_bytes().to_vec());
        }

        #[test]
        fn test_cell_reference_to_cell() {
            let actual = Sheet::cell_reference_to_cell(&"B1".as_bytes().to_vec()).unwrap();
            assert_eq!(actual, (1, 0))
        }

        #[test]
        fn test_cell_to_cell_reference() {
            let actual = Sheet::cell_to_cell_reference((0, 0));
            assert_eq!(actual, "A1".as_bytes())
        }

        #[test]
        fn test_cell_range_to_cell_reference() {
            let actual = Sheet::cell_range_to_cell_reference(&((4, 12), (8, 26)));
            assert_eq!(actual, "E13:I27".as_bytes())
        }

        #[test]
        fn test_cell_reference_to_cell_max_col() {
            let actual = Sheet::cell_reference_to_cell(&"XFE1".as_bytes().to_vec())
                .err()
                .unwrap()
                .to_string();
            assert_eq!(actual, "excel columns can not exceed 16,384")
        }

        #[test]
        fn test_cell_reference_to_cell_at_max_col() {
            let actual = Sheet::cell_reference_to_cell(&"XFD1".as_bytes().to_vec()).unwrap();

            assert_eq!(actual, (16_383, 0))
        }

        #[test]
        fn test_cell_reference_to_cell_at_max_row() {
            let actual = Sheet::cell_reference_to_cell(&"XFD1048576".as_bytes().to_vec()).unwrap();

            assert_eq!(actual, (16_383, 1_048_575))
        }

        #[test]
        fn test_cell_reference_to_cell_max_row() {
            let actual = Sheet::cell_reference_to_cell(&"XFD1048577".as_bytes().to_vec())
                .err()
                .unwrap()
                .to_string();
            assert_eq!(actual, "excel rows can not exceed 1,048,576")
        }

        #[test]
        fn test_cell_reference_to_cell_invalid_char() {
            let actual = Sheet::cell_reference_to_cell(&"*1048577".as_bytes().to_vec())
                .err()
                .unwrap()
                .to_string();
            assert_eq!(actual, "can not parse excel dimension: *1048577")
        }

        #[test]
        fn test_cell_reference_to_cell_not_within_a_to_z_range() {
            let actual = Sheet::cell_reference_to_cell(&"B京1048577".as_bytes().to_vec())
                .err()
                .unwrap()
                .to_string();
            assert_eq!(actual, "can not parse excel dimension: B京1048577")
        }

        #[test]
        fn test_cell_reference_to_cell_range_not_with_different_dimesions() {
            let actual = Sheet::cell_reference_to_cell_range(&"B1:D3".as_bytes().to_vec()).unwrap();
            assert_eq!(actual, ((1, 0), (3, 2)))
        }

        #[test]
        fn test_cell_reference_to_cell_range_with_same_dimensions() {
            let actual = Sheet::cell_reference_to_cell_range(&"B1".as_bytes().to_vec()).unwrap();
            assert_eq!(actual, ((1, 0), (1, 0)))
        }

        #[test]
        fn test_read_sheet() {
            let sheet = init("tests/workbook04.xlsx", "xl/worksheets/sheet2.xml");
            // assert_eq!(
            //     sheet,
            //     Sheet {
            //         path: "xl/worksheets/sheet2.xml".into(),
            //         name: Some("Sheet2".into()),
            //         tab_color: Some(Color::Theme { id: 5, tint: None }),
            //         fit_to_page: true,
            //         dimensions: ((1, 0), (6, 13)),
            //         show_grid: false,
            //         zoom_scale: Some("100".as_bytes().to_vec()),
            //         view_id: Some("0".as_bytes().to_vec()),
            //         ..Default::default()
            //     }
            // );
        }
    }
}
