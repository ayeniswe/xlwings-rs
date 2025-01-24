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
use derive::XmlWriter;
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
    Z0,
    Z25,
    Z50,
    Z75,
    #[default]
    Z100,
    Z200,
    Z400,
}

impl Into<Vec<u8>> for Zoom {
    fn into(self) -> Vec<u8> {
        match self {
            Zoom::Z0 => b"0".to_vec(),
            Zoom::Z25 => b"25".to_vec(),
            Zoom::Z50 => b"50".to_vec(),
            Zoom::Z75 => b"75".to_vec(),
            Zoom::Z100 => b"100".to_vec(),
            Zoom::Z200 => b"200".to_vec(),
            Zoom::Z400 => b"400".to_vec(),
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

// impl<W: Write> XmlWriter<W> for Pane {
//     fn write_xml<'a>(
//         &self,
//         writer: &'a mut Writer<W>,
//         tag_name: &'a str,
//     ) -> Result<&'a mut Writer<W>, XlsxError> {
//         let mut attrs = Vec::with_capacity(5);
//         if self.x_split != b"0" {
//             attrs.push((b"xSplit".as_ref(), self.x_split.as_ref()));
//         }
//         if self.y_split != b"0" {
//             attrs.push((b"ySplit".as_ref(), self.y_split.as_ref()));
//         }
//         let cell_ref;
//         if let Some(ref cell) = self.top_left_cell {
//             cell_ref = Sheet::cell_to_cell_reference(*cell);
//             attrs.push((b"topLeftCell".as_ref(), cell_ref.as_ref()));
//         }
//         if self.active_pane != PanePosition::default() {
//             attrs.push((
//                 b"activePane".as_ref(),
//                 match self.active_pane {
//                     PanePosition::BottomRight => b"bottomRight".as_ref(),
//                     PanePosition::TopRight => b"topRight".as_ref(),
//                     PanePosition::BottomLeft => b"bottomLeft".as_ref(),
//                     PanePosition::TopLeft => b"topLeft".as_ref(),
//                 },
//             ));
//         }
//         if self.state != PaneState::default() {
//             attrs.push((
//                 b"state".as_ref(),
//                 match self.state {
//                     PaneState::Frozen => b"frozen".as_ref(),
//                     PaneState::Split => b"split".as_ref(),
//                     PaneState::FrozenSplit => b"frozenSplit".as_ref(),
//                 },
//             ));
//         }
//         writer
//             .create_element(tag_name)
//             .with_attributes(attrs)
//             .write_empty()?;
//         Ok(writer)
//     }
// }

// Location of pane within sheet
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
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) enum PivotType {
    #[default]
    None,
    Normal,
    Data,
    All,
    Origin,
    Button,
    TopRight,
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
/// Represents a selected field and item within its parent
#[derive(Debug, Default, PartialEq, Clone, Eq)]
pub(crate) struct Reference {
    include_variance_filter: bool,
    include_pop_variance_filter: bool,
    include_sum_aggregate_filter: bool,
    include_prod_aggregate_filter: bool,
    include_min_aggregate_filter: bool,
    include_avg_aggregate_filter: bool,
    include_max_aggregate_filter: bool,
    include_std_deviation_filter: bool,
    include_pop_std_deviation_filter: bool,
    include_default_filter: bool,
    include_count_filter: bool,
    include_counta_filter: bool,
    selected: bool,
    relative: bool,
    by_position: bool,
    field: Option<Vec<u8>>,
    values: Vec<Vec<u8>>, // count attribute need to write and represents the length
}
impl Reference {
    pub(crate) fn new() -> Self {
        Self {
            ..Default::default()
        }
    }
}
/// A rule to describe `PivotTable` selection
#[derive(Debug, Default, PartialEq, Clone, Eq)]
pub(crate) struct PivotArea {
    references: Vec<Reference>, // count attribute need to write and represents the length
    axis: Option<Axis>,
    collapsed_are_subtotal: bool,
    offset: CellRange,
    outline: bool,
    cache_index: bool,
    include_col_total: bool,
    include_row_total: bool,
    use_label_only: bool,
    use_data_only: bool,
    pivot_type: PivotType,
    field: Option<Vec<u8>>,
    field_pos: Option<Vec<u8>>,
}
impl PivotArea {
    pub(crate) fn new() -> Self {
        Self {
            outline: true,
            ..Default::default()
        }
    }
}
/// Axis on a `PivotTable`
#[derive(Debug, Default, PartialEq, Clone, Eq)]
pub(crate) enum Axis {
    #[default]
    AxisRow,
    AxisCol,
    AxisPage,
    AxisValues,
}

/// Selections found within a `PivotTable`
#[derive(Debug, Default, PartialEq, Clone, Eq)]
pub(crate) struct PivotSelection {
    col: Vec<u8>,
    row: Vec<u8>,
    axis: Option<Axis>,
    click: Vec<u8>,
    count: Vec<u8>,
    data: bool,
    extendable: bool,
    label: bool,
    show_header: bool,
    dimension: Vec<u8>,
    max: Vec<u8>,
    min: Vec<u8>,
    pane: PanePosition,
    rid: Option<Vec<u8>>,
    prev_col: Vec<u8>,
    prev_row: Vec<u8>,
    start: Vec<u8>,
    area: PivotArea,
}

impl PivotSelection {
    pub(crate) fn new() -> Self {
        Self {
            col: b"0".to_vec(),
            row: b"0".to_vec(),
            click: b"0".to_vec(),
            count: b"0".to_vec(),
            dimension: b"0".to_vec(),
            max: b"0".to_vec(),
            min: b"0".to_vec(),
            prev_col: b"0".to_vec(),
            prev_row: b"0".to_vec(),
            start: b"0".to_vec(),
            ..Default::default()
        }
    }
}
/// A default view of a sheet
#[derive(Debug, Default, PartialEq, Clone, Eq)]
pub(crate) struct SheetView {
    zoom_scale: Vec<u8>,
    zoom_scale_normal: Vec<u8>,
    zoom_scale_page: Vec<u8>,
    zoom_scale_sheet: Vec<u8>,
    view_id: Vec<u8>,
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
    pivot_selection: Option<PivotSelection>,
    show_header: bool,
    show_outline_symbol: bool,
}
impl SheetView {
    fn new(id: u32) -> Self {
        Self {
            show_grid: true,
            show_zero: true,
            show_outline_symbol: true,
            zoom_scale: Zoom::Z100.into(),
            zoom_scale_normal: Zoom::Z0.into(),
            zoom_scale_page: Zoom::Z0.into(),
            zoom_scale_sheet: Zoom::Z0.into(),
            view_id: id.to_string().as_bytes().to_vec(),
            show_ruler: true,
            show_header: true,
            show_whitespace: true,
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
                // sheetViews
                for view in &self.sheet_views {
                    let mut attrs = Vec::with_capacity(9);
                    if view.use_protection {
                        attrs.push((b"windowProtection".as_ref(), b"1".as_ref()));
                    }
                    if view.show_formula {
                        attrs.push((b"showFormulas".as_ref(), b"1".as_ref()));
                    }
                    if !view.show_grid {
                        attrs.push((b"showGridLines".as_ref(), b"0".as_ref()));
                    }
                    if !view.show_header {
                        attrs.push((b"showRowColHeaders".as_ref(), b"0".as_ref()));
                    }
                    if !view.show_zero {
                        attrs.push((b"showZeros".as_ref(), b"0".as_ref()));
                    }
                    if view.use_rtl {
                        attrs.push((b"rightToLeft".as_ref(), b"1".as_ref()));
                    }
                    if view.show_tab {
                        attrs.push((b"tabSelected".as_ref(), b"1".as_ref()));
                    }
                    if !view.show_ruler {
                        attrs.push((b"showRuler".as_ref(), b"0".as_ref()));
                    }
                    if !view.show_outline_symbol {
                        attrs.push((b"showOutlineSymbols".as_ref(), b"0".as_ref()));
                    }
                    let grid_color;
                    if view.grid_color != GridlineColor::Automatic {
                        grid_color = (view.grid_color.clone() as u8).to_string();
                        attrs.push((b"colorId".as_ref(), grid_color.as_ref()));
                    }
                    if !view.show_whitespace {
                        attrs.push((b"showWhiteSpace".as_ref(), b"0".as_ref()));
                    }
                    if let Some(ref view_type) = view.view {
                        let view = match view_type {
                            View::Normal => b"".as_ref(),
                            View::PageBreakPreview => b"pageBreakPreview".as_ref(),
                            View::PageLayout => b"pageLayout".as_ref(),
                        };
                        if view != b"" {
                            attrs.push((b"view".as_ref(), view));
                        }
                    }
                    let top_left_cell;
                    if let Some(ref cell) = view.top_left_cell {
                        top_left_cell = Sheet::cell_to_cell_reference(*cell);
                        attrs.push((b"topLeftCell".as_ref(), &top_left_cell));
                    }
                    if view.zoom_scale != <Zoom as Into<Vec<u8>>>::into(Zoom::Z100) {
                        attrs.push((b"zoomScale".as_ref(), &view.zoom_scale));
                    }
                    if view.zoom_scale_normal != <Zoom as Into<Vec<u8>>>::into(Zoom::Z0) {
                        attrs.push((b"zoomScaleNormal".as_ref(), &view.zoom_scale_normal));
                    }
                    if view.zoom_scale_sheet != <Zoom as Into<Vec<u8>>>::into(Zoom::Z0) {
                        attrs.push((b"zoomScaleSheetLayoutView".as_ref(), &view.zoom_scale_sheet));
                    }
                    if view.zoom_scale_page != <Zoom as Into<Vec<u8>>>::into(Zoom::Z0) {
                        attrs.push((b"zoomScalePageLayoutView".as_ref(), &view.zoom_scale_page));
                    }
                    attrs.push((b"workbookViewId".as_ref(), &view.view_id));
                    writer
                        .create_element("sheetView")
                        .with_attributes(attrs)
                        .write_inner_content::<_, XlsxError>(|writer| {
                            // pane
                            // if let Some(ref pane) = view.pane {
                            //     pane.write_xml(writer, "pane")?;
                            // }
                            // // selection
                            // if let Some(ref selection) = view.selection {
                            //     selection.write_xml(writer, "selection")?;
                            // }
                            // // pivotSelection
                            // if let Some(ref pivot_selection) = view.pivot_selection {
                            //     pivot_selection.write_xml(writer, "pivotSelection")?;
                            // }

                            Ok(())
                        })?;
                }
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
                // TAB COLOR
                /////////////
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"tabColor" => {
                    self.tab_color = Some(Stylesheet::read_color(e.attributes())?);
                }
                ////////////////////
                // PAGE SETUP PROPERTIES
                /////////////
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
                ////////////////////
                // OUTLINE PROPERTIES
                /////////////
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
                        let mut sheet_view = SheetView::new(0);
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
                                            b"zoomScaleNormal" => {
                                                sheet_view.zoom_scale_normal = a.value.into();
                                            }
                                            b"zoomScale" => {
                                                sheet_view.zoom_scale = a.value.into();
                                            }
                                            b"zoomScalePageLayoutView" => {
                                                sheet_view.zoom_scale_page = a.value.into();
                                            }
                                            b"zoomScaleSheetLayoutView" => {
                                                sheet_view.zoom_scale_sheet = a.value.into();
                                            }
                                            b"workbookViewId" => {
                                                sheet_view.view_id = a.value.into()
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
                                            ////////////////////
                                            // Pane
                                            /////////////
                                            Ok(Event::Empty(ref e))
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
                                                                    b"bottomRight" => pane
                                                                        .active_pane =
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
                                            // Selection
                                            /////////////
                                            Ok(Event::Start(ref e))
                                                if e.local_name().as_ref() == b"selection" =>
                                            {
                                                let mut selection = Selection::new();
                                                for attr in e.attributes() {
                                                    if let Ok(a) = attr {
                                                        match a.key.as_ref() {
                                                            b"pane" => match a.value.as_ref() {
                                                                b"bottomLeft" => {
                                                                    selection.pane =
                                                                        PanePosition::BottomLeft
                                                                }
                                                                b"bottomRight" => {
                                                                    selection.pane =
                                                                        PanePosition::BottomRight
                                                                }
                                                                b"topLeft" => {
                                                                    selection.pane =
                                                                        PanePosition::TopLeft
                                                                }
                                                                b"topRight" => {
                                                                    selection.pane =
                                                                        PanePosition::TopRight
                                                                }
                                                                _ => (),
                                                            },
                                                            b"activeCellId" => {
                                                                selection.cell_id = a.value.into();
                                                            }
                                                            b"activeCell" => {
                                                                selection.cell = Some(
                                                                    Sheet::cell_reference_to_cell(
                                                                        &a.value,
                                                                    )?,
                                                                );
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
                                            ////////////////////
                                            // PIVOT SELECTION
                                            /////////////
                                            Ok(Event::Start(ref e))
                                                if e.local_name().as_ref() == b"pivotSelection" =>
                                            {
                                                let mut pivot = PivotSelection::new();
                                                for attr in e.attributes() {
                                                    if let Ok(a) = attr {
                                                        match a.key.as_ref() {
                                                            b"pane" => {
                                                                match a.value.as_ref() {
                                                                    b"bottomLeft" => {
                                                                        pivot.pane =
                                                                            PanePosition::BottomLeft
                                                                    }
                                                                    b"bottomRight" => pivot.pane =
                                                                        PanePosition::BottomRight,
                                                                    b"topLeft" => {
                                                                        pivot.pane =
                                                                            PanePosition::TopLeft
                                                                    }
                                                                    b"topRight" => {
                                                                        pivot.pane =
                                                                            PanePosition::TopRight
                                                                    }
                                                                    _ => (),
                                                                }
                                                            }
                                                            b"showHeader" => {
                                                                pivot.show_header =
                                                                    *a.value == *b"1";
                                                            }
                                                            b"label" => {
                                                                pivot.label = *a.value == *b"1";
                                                            }
                                                            b"data" => {
                                                                pivot.data = *a.value == *b"1";
                                                            }
                                                            b"extendable" => {
                                                                pivot.extendable =
                                                                    *a.value == *b"1";
                                                            }
                                                            b"dimension" => {
                                                                pivot.dimension = a.value.into();
                                                            }
                                                            b"start" => {
                                                                pivot.start = a.value.into();
                                                            }
                                                            b"min" => {
                                                                pivot.min = a.value.into();
                                                            }
                                                            b"max" => {
                                                                pivot.max = a.value.into();
                                                            }
                                                            b"activeRow" => {
                                                                pivot.row = a.value.into();
                                                            }
                                                            b"activeCol" => {
                                                                pivot.col = a.value.into();
                                                            }
                                                            b"previousRow" => {
                                                                pivot.prev_row = a.value.into();
                                                            }
                                                            b"previousCol" => {
                                                                pivot.prev_col = a.value.into();
                                                            }
                                                            b"click" => {
                                                                pivot.click = a.value.into();
                                                            }
                                                            b"r:id" => {
                                                                pivot.rid = Some(a.value.into());
                                                            }
                                                            b"axis" => match a.value.as_ref() {
                                                                b"axisCol" => {
                                                                    pivot.axis = Some(Axis::AxisCol)
                                                                }
                                                                b"axisPage" => {
                                                                    pivot.axis =
                                                                        Some(Axis::AxisPage)
                                                                }
                                                                b"axisRow" => {
                                                                    pivot.axis = Some(Axis::AxisRow)
                                                                }
                                                                b"axisValues" => {
                                                                    pivot.axis =
                                                                        Some(Axis::AxisValues)
                                                                }
                                                                _ => (),
                                                            },

                                                            _ => (),
                                                        }
                                                    }
                                                }

                                                /////////////////////
                                                //// PIVOT AREA
                                                ///////////////
                                                let mut area_buf = Vec::with_capacity(1024);
                                                loop {
                                                    let mut area = PivotArea::new();
                                                    area_buf.clear();
                                                    match xml.read_event_into(&mut area_buf) {
                                                        Ok(Event::Start(ref e))
                                                            if e.local_name().as_ref()
                                                                == b"pivotArea" =>
                                                        {
                                                            for attr in e.attributes() {
                                                                if let Ok(a) = attr {
                                                                    match a.key.as_ref() {
                                                                        b"field" => {
                                                                            area.field = Some(
                                                                                a.value.to_vec(),
                                                                            )
                                                                        }
                                                                        b"type" => {
                                                                            match a.value.as_ref() {
                                                                                b"all" => area
                                                                                    .pivot_type =
                                                                                    PivotType::All,
                                                                                b"button" => area
                                                                                    .pivot_type =
                                                                                    PivotType::Button,
                                                                                b"data" => area
                                                                                    .pivot_type =
                                                                                    PivotType::Data,
                                                                                b"none" => area
                                                                                    .pivot_type =
                                                                                    PivotType::None,
                                                                                b"normal" => area
                                                                                    .pivot_type =
                                                                                    PivotType::Normal,
                                                                                b"origin" => area
                                                                                    .pivot_type =
                                                                                    PivotType::Origin,
                                                                                b"topRight" => area
                                                                                    .pivot_type =
                                                                                    PivotType::TopRight,
                                                                                _ => (),
                                                                            }
                                                                        }
                                                                        b"dataOnly" => {
                                                                            area.use_data_only =
                                                                                *a.value == *b"1"
                                                                        }
                                                                        b"labelOnly" => {
                                                                            area.use_label_only =
                                                                                *a.value == *b"1"
                                                                        }
                                                                        b"grandRow" => {
                                                                            area.include_row_total =
                                                                                *a.value == *b"1"
                                                                        }
                                                                        b"grandCol" => {
                                                                            area.include_col_total =
                                                                                *a.value == *b"1"
                                                                        }
                                                                        b"cacheIndex" => {
                                                                            area.cache_index =
                                                                                *a.value == *b"1"
                                                                        }
                                                                        b"outline" => {
                                                                            area.outline =
                                                                                *a.value == *b"1"
                                                                        }
                                                                        b"offset" => {
                                                                            let offset = Sheet::cell_reference_to_cell_range(&a.value)?;
                                                                            area.offset = offset;
                                                                        }
                                                                        // This binary string breaks rustfmt with line length
                                                                        b"collapsedLevelsAreSubtotals" => {
                                                                            area.collapsed_are_subtotal = *a.value == *b"1";
                                                                        }
                                                                        b"axis" => {
                                                                            match a.value.as_ref() {
                                                                                b"axisCol" =>  area.axis = Some(Axis::AxisCol),
                                                                                b"axisPage" => area.axis = Some(Axis::AxisPage),
                                                                                b"axisRow" => area.axis = Some(Axis::AxisRow),
                                                                                b"axisValues" =>  area.axis = Some(Axis::AxisValues),
                                                                                _ => ()
                                                                            }
                                                                        }
                                                                        b"fieldPosition" => {
                                                                            area.field_pos = Some(
                                                                                a.value.to_vec(),
                                                                            )
                                                                        }
                                                                        _ => (),
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        //////////////////
                                                        //// REFERENCE
                                                        //////////////
                                                        Ok(Event::Start(ref e))
                                                            if e.local_name().as_ref()
                                                                == b"reference" =>
                                                        {
                                                            let mut reference = Reference::new();
                                                            for attr in e.attributes() {
                                                                if let Ok(a) = attr {
                                                                    match a.key.as_ref() {
                                                                        b"field" => {
                                                                            reference.field = Some(
                                                                                a.value.to_vec(),
                                                                            )
                                                                        }
                                                                        b"selected" => {
                                                                            reference.selected =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"byPosition" => {
                                                                            reference.selected =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"relative" => {
                                                                            reference.relative =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"defaultSubtotal" => {
                                                                            reference.include_default_filter =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"sumSubtotal" => {
                                                                            reference.include_sum_aggregate_filter =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"countASubtotal" => {
                                                                            reference.include_counta_filter =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"avgSubtotal" => {
                                                                            reference.include_avg_aggregate_filter =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"maxSubtotal" => {
                                                                            reference.include_max_aggregate_filter =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"minSubtotal" => {
                                                                            reference.include_min_aggregate_filter =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"productSubtotal" => {
                                                                            reference.include_prod_aggregate_filter =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"countSubtotal" => {
                                                                            reference.include_count_filter =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"stdDevSubtotal" => {
                                                                            reference.include_std_deviation_filter =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"stdDevPSubtotal" => {
                                                                            reference.include_pop_std_deviation_filter =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"varSubtotal" => {
                                                                            reference.include_variance_filter =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        b"varPSubtotal" => {
                                                                            reference.include_pop_variance_filter =
                                                                                *a.value == *b"1";
                                                                        }
                                                                        _ => (),
                                                                    }
                                                                }
                                                            }
                                                            let mut ref_buf =
                                                                Vec::with_capacity(1024);
                                                            //////////////////
                                                            //// REFERENCE nth-1
                                                            //////////////
                                                            loop {
                                                                ref_buf.clear();
                                                                match xml
                                                                    .read_event_into(&mut ref_buf)
                                                                {
                                                                    //////////////////
                                                                    //// REFERENCE SHARED VALUE
                                                                    //////////////
                                                                    Ok(Event::Empty(ref e))
                                                                        if e.local_name()
                                                                            .as_ref()
                                                                            == b"x" =>
                                                                    {
                                                                        //////////////////
                                                                        //// SHARED VALUE
                                                                        //////////////
                                                                        for attr in e.attributes() {
                                                                            if let Ok(a) = attr {
                                                                                match a.key.as_ref() {
                                                                                            b"v" => {
                                                                                                reference.values.push(
                                                                                                    a.value.to_vec()
                                                                                                )
                                                                                            },
                                                                                            _ => ()
                                                                                        }
                                                                            }
                                                                        }
                                                                    }
                                                                    Ok(Event::End(ref e))
                                                                        if e.local_name()
                                                                            .as_ref()
                                                                            == b"reference" =>
                                                                    {
                                                                        pivot
                                                                            .area
                                                                            .references
                                                                            .push(reference);
                                                                        break;
                                                                    }
                                                                    Ok(Event::Eof) => {
                                                                        return Err(
                                                                            XlsxError::XmlEof(
                                                                                "reference".into(),
                                                                            ),
                                                                        )
                                                                    }
                                                                    Err(e) => {
                                                                        return Err(
                                                                            XlsxError::Xml(e),
                                                                        );
                                                                    }
                                                                    _ => (),
                                                                }
                                                            }
                                                        }
                                                        Ok(Event::End(ref e))
                                                            if e.local_name().as_ref()
                                                                == b"pivotArea" =>
                                                        {
                                                            pivot.area = area;
                                                            break;
                                                        }
                                                        Ok(Event::Eof) => {
                                                            return Err(XlsxError::XmlEof(
                                                                "pivotArea".into(),
                                                            ))
                                                        }
                                                        Err(e) => {
                                                            return Err(XlsxError::Xml(e));
                                                        }
                                                        _ => (),
                                                    }
                                                }
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
        use super::init;
        use crate::stream::xlsx::sheet::{Color, GridlineColor, Sheet};
        use std::{
            fs::File,
            io::{Seek, SeekFrom},
        };

        #[test]
        fn test_cell_group_to_list_cell_range() {
            let actual =
                Sheet::cell_group_to_list_cell_range(&"A1:B2 D4:D5 F10".as_bytes().to_vec())
                    .unwrap();
            assert_eq!(actual[0], ((0, 0), (1, 1)));
            assert_eq!(actual[1], ((3, 3), (3, 4)));
            assert_eq!(actual[2], ((5, 9), (5, 9)));
        }

        #[test]
        fn test_list_cell_range_to_cell_group() {
            let actual =
                Sheet::cell_group_to_list_cell_range(&"A1:B2 D4:D5 F10".as_bytes().to_vec())
                    .unwrap();
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

#[derive(XmlWriter)]
#[xml(name = "ex")]
struct Example {
    #[xml(name = "activePane")]
    active_pane: bool,
    #[xml(default_bool = true)]
    x_split: bool,
    value_test: Vec<u8>,
    #[xml(default_bytes = b"test")]
    open_win: Vec<u8>,
}

#[test]
fn test_xml_writer_derive() {
    let sheet = Example {
        active_pane: false,
        x_split: true,
        value_test: b"01234".to_vec(),
        open_win: b"test".to_vec(),
    };

    let mut buffer = Cursor::new(Vec::new());
    let mut writer = Writer::new(&mut buffer);
    let _ = sheet.write_xml(&mut writer, "sheet");

    let xml_output = String::from_utf8(buffer.into_inner()).unwrap();
    let expected_output = r#"<ex activePane="0" value_test="01234"/>"#;
    assert_eq!(xml_output, expected_output);
}