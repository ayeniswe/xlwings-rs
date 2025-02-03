use crate::stream::{
    utils::{XmlReader, XmlWriter},
    xlsx::XlsxError,
};
use derive::{XmlRead, XmlWrite};
use quick_xml::{events::Event, Reader, Writer};
use std::io::{BufRead, Write};

use super::{pane::CTPane, pivot::CTPivotSelection, selection::CTSelection};


/// Represents a sheet view in a spreadsheet, defining visual and behavioral settings for a worksheet.
///
/// This struct corresponds to the `CT_SheetView` complex type in the XML schema. It encapsulates
/// attributes and elements that control how a sheet is displayed, including visibility settings,
/// zoom levels, and UI elements like grid lines, headers, and rulers.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_SheetView">
///     <sequence>
///         <element name="pane" type="CT_Pane" minOccurs="0" maxOccurs="1"/>
///         <element name="selection" type="CT_Selection" minOccurs="0" maxOccurs="4"/>
///         <element name="pivotSelection" type="CT_PivotSelection" minOccurs="0" maxOccurs="4"/>
///     </sequence>
///     <attribute name="windowProtection" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="showFormulas" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="showGridLines" type="xsd:boolean" use="optional" default="true"/>
///     <attribute name="showRowColHeaders" type="xsd:boolean" use="optional" default="true"/>
///     <attribute name="showZeros" type="xsd:boolean" use="optional" default="true"/>
///     <attribute name="rightToLeft" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="tabSelected" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="showRuler" type="xsd:boolean" use="optional" default="true"/>
///     <attribute name="showOutlineSymbols" type="xsd:boolean" use="optional" default="true"/>
///     <attribute name="defaultGridColor" type="xsd:boolean" use="optional" default="true"/>
///     <attribute name="showWhiteSpace" type="xsd:boolean" use="optional" default="true"/>
///     <attribute name="view" type="ST_SheetViewType" use="optional" default="normal"/>
///     <attribute name="topLeftCell" type="ST_CellRef" use="optional"/>
///     <attribute name="colorId" type="xsd:unsignedInt" use="optional" default="64"/>
///     <attribute name="zoomScale" type="xsd:unsignedInt" use="optional" default="100"/>
///     <attribute name="zoomScaleNormal" type="xsd:unsignedInt" use="optional" default="0"/>
///     <attribute name="zoomScaleSheetLayoutView" type="xsd:unsignedInt" use="optional" default="0"/>
///     <attribute name="zoomScalePageLayoutView" type="xsd:unsignedInt" use="optional" default="0"/>
///     <attribute name="workbookViewId" type="xsd:unsignedInt" use="required"/>
/// </complexType>
/// ```
///
/// # Fields
/// ## Attributes
/// - `use_protection`: Enables or disables window protection (`windowProtection`).
/// - `show_formula`: Controls the visibility of formulas (`showFormulas`).
/// - `show_grid`: Toggles the display of grid lines (`showGridLines`).
/// - `show_header`: Toggles the display of row and column headers (`showRowColHeaders`).
/// - `show_zero`: Controls the visibility of zero values (`showZeros`).
/// - `use_rtl`: Enables right-to-left layout (`rightToLeft`).
/// - `show_tab`: Indicates whether the sheet tab is selected (`tabSelected`).
/// - `show_ruler`: Toggles the display of the ruler (`showRuler`).
/// - `show_outline_symbol`: Controls the visibility of outline symbols (`showOutlineSymbols`).
/// - `grid_color`: Enables or disables the default grid color (`defaultGridColor`).
/// - `show_whitespace`: Toggles the display of whitespace (`showWhiteSpace`).
/// - `view`: Specifies the view type (`view`), e.g., "normal", "page layout".
/// - `top_left_cell`: The top-left cell visible in the view (`topLeftCell`).
/// - `color_id`: The color ID for the sheet (`colorId`).
/// - `zoom_scale`: The zoom scale percentage (`zoomScale`).
/// - `zoom_scale_normal`: The zoom scale for normal view (`zoomScaleNormal`).
/// - `zoom_scale_sheet`: The zoom scale for sheet layout view (`zoomScaleSheetLayoutView`).
/// - `zoom_scale_page`: The zoom scale for page layout view (`zoomScalePageLayoutView`).
/// - `view_id`: The unique ID of the workbook view (`workbookViewId`).
///
/// ## Elements
/// - `pane`: Represents the pane settings for the sheet (`pane`).
/// - `selection`: Represents the selected cells or ranges (`selection`).
/// - `pivot_selection`: Represents the pivot table selection (`pivotSelection`).
/// ```
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite, XmlRead)]
pub(crate) struct CTSheetView {
    #[xml(name = "windowProtection", default_bool = false)]
    use_protection: bool,
    #[xml(name = "showFormulas", default_bool = false)]
    show_formula: bool,
    #[xml(name = "showGridLines", default_bool = true)]
    show_grid: bool,
    #[xml(name = "showRowColHeaders", default_bool = true)]
    show_header: bool,
    #[xml(name = "showZeros", default_bool = true)]
    show_zero: bool,
    #[xml(name = "rightToLeft", default_bool = false)]
    use_rtl: bool,
    #[xml(name = "tabSelected", default_bool = false)]
    show_tab: bool,
    #[xml(name = "showRuler", default_bool = true)]
    show_ruler: bool,
    #[xml(name = "showOutlineSymbols", default_bool = true)]
    show_outline_symbol: bool,
    #[xml(name = "defaultGridColor", default_bool = true)]
    use_default_grid_color: bool,
    #[xml(name = "showWhiteSpace", default_bool = true)]
    show_whitespace: bool,
    #[xml(name = "view", default_bytes = b"normal")]
    view: Vec<u8>,
    #[xml(name = "topLeftCell")]
    top_left_cell: Vec<u8>,
    #[xml(name = "colorId", default_bytes = b"64")]
    color_id: Vec<u8>,
    #[xml(name = "zoomScale", default_bytes = b"100")]
    zoom_scale: Vec<u8>,
    #[xml(name = "zoomScaleNormal", default_bytes = b"0")]
    zoom_scale_normal: Vec<u8>,
    #[xml(name = "zoomScaleSheetLayoutView", default_bytes = b"0")]
    zoom_scale_sheet: Vec<u8>,
    #[xml(name = "zoomScalePageLayoutView", default_bytes = b"0")]
    zoom_scale_page: Vec<u8>,
    #[xml(name = "workbookViewId")]
    view_id: Vec<u8>,

    #[xml(following_elements, name = "pane")]
    pane: Option<CTPane>,
    #[xml(name = "selection")]
    selection: Option<CTSelection>,
    #[xml(name = "pivotSelection")]
    pivot_selection: Option<CTPivotSelection>,
}
impl CTSheetView {
    /// Creates a new `CT_SheetView` instance with xml schema default values.
    fn new(id: u32) -> Self {
        Self {
            show_grid: true,
            show_zero: true,
            view: b"normal".into(),
            use_default_grid_color: true,
            show_outline_symbol: true,
            zoom_scale: b"100".into(),
            zoom_scale_normal: b"0".into(),
            zoom_scale_sheet: b"0".into(),
            zoom_scale_page: b"0".into(),
            color_id: b"64".into(),
            view_id: id.to_string().as_bytes().into(),
            show_ruler: true,
            show_header: true,
            show_whitespace: true,
            ..Default::default()
        }
    }
}