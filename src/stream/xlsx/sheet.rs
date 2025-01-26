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
use derive::XmlWrite;
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

/// Axis on a `PivotTable`
#[derive(Debug, Default, PartialEq, Clone, Eq)]
pub(crate) enum Axis {
    #[default]
    AxisRow,
    AxisCol,
    AxisPage,
    AxisValues,
}

// Location of pane within sheet
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) enum PanePosition {
    BottomRight,
    TopRight,
    BottomLeft,
    #[default]
    TopLeft,
}
impl TryFrom<Vec<u8>> for PanePosition {
    type Error = XlsxError;

    fn try_from(value: Vec<u8>) -> Result<Self, Self::Error> {
        match value.as_slice() {
            b"bottomLeft" => Ok(PanePosition::BottomLeft),
            b"bottomRight" => Ok(PanePosition::BottomRight),
            b"topLeft" => Ok(PanePosition::TopLeft),
            b"topRight" => Ok(PanePosition::TopRight),
            v => {
                let value = String::from_utf8_lossy(v);
                Err(XlsxError::MissingVariant(
                    "PanePosition".into(),
                    value.into(),
                ))
            }
        }
    }
}
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) enum PaneState {
    Frozen,
    #[default]
    Split,
    FrozenSplit,
}
impl TryFrom<Vec<u8>> for PaneState {
    type Error = XlsxError;

    fn try_from(value: Vec<u8>) -> Result<Self, Self::Error> {
        match value.as_slice() {
            b"frozen" => Ok(PaneState::Frozen),
            b"split" => Ok(PaneState::Split),
            b"frozenSplit" => Ok(PaneState::FrozenSplit),
            v => {
                let value = String::from_utf8_lossy(v);
                Err(XlsxError::MissingVariant("PaneState".into(), value.into()))
            }
        }
    }
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
impl TryFrom<Vec<u8>> for PivotType {
    type Error = XlsxError;

    fn try_from(value: Vec<u8>) -> Result<Self, Self::Error> {
        match value.as_slice() {
            b"none" => Ok(PivotType::None),
            b"normal" => Ok(PivotType::Normal),
            b"data" => Ok(PivotType::Data),
            b"all" => Ok(PivotType::All),
            b"origin" => Ok(PivotType::Origin),
            b"button" => Ok(PivotType::Button),
            b"topRight" => Ok(PivotType::TopRight),
            v => {
                let value = String::from_utf8_lossy(v);
                Err(XlsxError::MissingVariant("PivotType".into(), value.into()))
            }
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
impl TryFrom<Vec<u8>> for GridlineColor {
    type Error = XlsxError;

    fn try_from(value: Vec<u8>) -> Result<Self, Self::Error> {
        match value.as_slice() {
            b"0" => Ok(GridlineColor::Automatic),
            b"8" => Ok(GridlineColor::Black),
            b"15" => Ok(GridlineColor::Turquoise),
            b"60" => Ok(GridlineColor::Brown),
            b"14" => Ok(GridlineColor::Pink),
            b"59" => Ok(GridlineColor::OliveGreen),
            b"58" => Ok(GridlineColor::DarkGreen),
            b"56" => Ok(GridlineColor::DarkTeal),
            b"18" => Ok(GridlineColor::DarkBlue),
            b"62" => Ok(GridlineColor::Indigo),
            b"63" => Ok(GridlineColor::Gray80),
            b"23" => Ok(GridlineColor::Gray50),
            b"55" => Ok(GridlineColor::Gray40),
            b"22" => Ok(GridlineColor::Gray25),
            b"9" => Ok(GridlineColor::White),
            b"31" => Ok(GridlineColor::IceBlue),
            b"12" => Ok(GridlineColor::Blue),
            b"21" => Ok(GridlineColor::Teal),
            b"30" => Ok(GridlineColor::OceanBlue),
            b"25" => Ok(GridlineColor::Plum),
            b"46" => Ok(GridlineColor::Lavender),
            b"20" => Ok(GridlineColor::Violet),
            b"54" => Ok(GridlineColor::BlueGray),
            b"48" => Ok(GridlineColor::LightBlue),
            b"40" => Ok(GridlineColor::SkyBlue),
            b"44" => Ok(GridlineColor::PaleBlue),
            b"29" => Ok(GridlineColor::Coral),
            b"16" => Ok(GridlineColor::DarkRed),
            b"49" => Ok(GridlineColor::Aqua),
            b"27" => Ok(GridlineColor::LightTurquoise),
            b"28" => Ok(GridlineColor::DarkPurple),
            b"57" => Ok(GridlineColor::SeaGreen),
            b"42" => Ok(GridlineColor::LightGreen),
            b"11" => Ok(GridlineColor::BrightGreen),
            b"13" => Ok(GridlineColor::Yellow),
            b"26" => Ok(GridlineColor::Ivory),
            b"43" => Ok(GridlineColor::LightYellow),
            b"19" => Ok(GridlineColor::DarkYellow),
            b"50" => Ok(GridlineColor::Lime),
            b"53" => Ok(GridlineColor::Orange),
            b"52" => Ok(GridlineColor::LightOrange),
            b"51" => Ok(GridlineColor::Gold),
            b"47" => Ok(GridlineColor::Tan),
            b"45" => Ok(GridlineColor::Rose),
            b"24" => Ok(GridlineColor::Periwinkle),
            b"10" => Ok(GridlineColor::Red),
            b"17" => Ok(GridlineColor::Green),
            v => {
                let value = String::from_utf8_lossy(v);
                Err(XlsxError::MissingVariant(
                    "GridlineColor".into(),
                    value.to_string(),
                ))
            }
        }
    }
}

/// Predefined zoom level for a sheet
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub(crate) enum Zoom {
    /// 0%
    Z0,
    /// 25%
    Z25,
    /// 50%
    Z50,
    /// 75%
    Z75,
    #[default]
    /// 100%
    Z100,
    /// 200%
    Z200,
    /// 400%
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
impl TryFrom<Vec<u8>> for View {
    type Error = XlsxError;

    fn try_from(value: Vec<u8>) -> Result<Self, Self::Error> {
        match value.as_slice() {
            b"normal" => Ok(View::Normal),
            b"pageBreakPreview" => Ok(View::PageBreakPreview),
            b"pageLayout" => Ok(View::PageLayout),
            v => {
                let value = String::from_utf8_lossy(v);
                Err(XlsxError::MissingVariant("View".into(), value.into()))
            }
        }
    }
}

/// Represents a pane in a spreadsheet, defining the split and active pane settings.
///
/// This struct corresponds to the `CT_Pane` complex type in the XML schema. It encapsulates
/// attributes that control the position of splits, the active pane, and the state of the pane.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_Pane">
///     <attribute name="xSplit" type="xsd:double" use="optional" default="0"/>
///     <attribute name="ySplit" type="xsd:double" use="optional" default="0"/>
///     <attribute name="topLeftCell" type="ST_CellRef" use="optional"/>
///     <attribute name="activePane" type="ST_Pane" use="optional" default="topLeft"/>
///     <attribute name="state" type="ST_PaneState" use="optional" default="split"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `x_split`: The horizontal split position (`xSplit`).
/// - `y_split`: The vertical split position (`ySplit`).
/// - `top_left_cell`: The top-left cell visible in the pane (`topLeftCell`).
/// - `active_pane`: The active pane (`activePane`).
/// - `state`: The state of the pane (`state`), e.g., "split", "frozen".
///
/// # Example
/// ```rust
/// use my_crate::Pane;
///
/// let pane = Pane {
///     x_split: b"0".to_vec(),
///     y_split: b"0".to_vec(),
///     top_left_cell: b"A1".to_vec(),
///     active_pane: b"topLeft".to_vec(),
///     state: b"split".to_vec(),
/// };
/// ```
#[derive(Debug, Default, Clone, PartialEq, Eq, XmlWrite)]
pub(crate) struct Pane {
    #[xml(name = "xSplit", default_bytes = b"0")]
    x_split: Vec<u8>,
    #[xml(name = "ySplit", default_bytes = b"0")]
    y_split: Vec<u8>,
    #[xml(name = "topLeftCell")]
    top_left_cell: Vec<u8>,
    #[xml(name = "activePane", default_bytes = b"topLeft")]
    active_pane: Vec<u8>,
    #[xml(name = "state", default_bytes = b"split")]
    state: Vec<u8>,
}
impl Pane {
    /// Creates a new `Pane` instance with xml schema default values.
    pub fn new() -> Self {
        Self {
            ..Default::default()
        }
    }
}

/// Represents a selection within a sheet view, defining the active cell, pane, and selected range.
///
/// This struct corresponds to the `CT_Selection` complex type in the XML schema. It encapsulates
/// attributes that specify the active cell, the pane in which the selection is active, and the
/// range of selected cells.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_Selection">
///     <attribute name="pane" type="ST_Pane" use="optional"/>
///     <attribute name="activeCell" type="ST_CellRef" use="optional"/>
///     <attribute name="activeCellId" type="xsd:unsignedInt" use="optional" default="0"/>
///     <attribute name="sqref" type="ST_Sqref" use="optional" default="A1"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `pane`: The pane in which the selection is active (`pane`).
/// - `active_cell`: The active cell within the selection (`activeCell`).
/// - `active_cell_id`: The ID of the active cell (`activeCellId`).
/// - `sqref`: The range of selected cells (`sqref`).
/// ```
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite)]
pub(crate) struct Selection {
    #[xml(name = "pane")]
    pane: Vec<u8>,
    #[xml(name = "activeCell")]
    cell: Vec<u8>,
    #[xml(name = "activeCellId", default_bytes = b"0")]
    cell_id: Vec<u8>,
    #[xml(name = "sqref", default_bytes = b"A1")]
    sqref: Vec<u8>,
}
impl Selection {
    /// Creates a new `Selection` instance with xml schema default values.
    pub(crate) fn new() -> Self {
        Self {
            ..Default::default()
        }
    }
}

/// Represents an index within a `PivotTable` selection.
///
/// This struct corresponds to the `CT_Index` complex type in the XML schema. It encapsulates
/// an unsigned integer value that represents an index.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_Index">
///     <attribute name="v" use="required" type="xsd:unsignedInt"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `item`: The index value (`v`).
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite)]
pub(crate) struct CTIndex {
    #[xml(name = "v")]
    item: Vec<u8>,
}
impl CTIndex {
    /// Creates a new `CT_Index` instance with xml schema default values.
    pub(crate) fn new() -> Self {
        Self {
            ..Default::default()
        }
    }
}

/// Represents a selected field and item within its parent in a `PivotTable`.
///
/// This struct corresponds to the `CT_PivotAreaReference` complex type in the XML schema. It encapsulates
/// attributes and elements that define the selection of a field, its position, and various filters and
/// subtotals.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_PivotAreaReference">
///     <sequence>
///         <element name="x" minOccurs="0" maxOccurs="unbounded" type="CT_Index"/>
///         <element name="extLst" minOccurs="0" type="CT_ExtensionList"/>
///     </sequence>
///     <attribute name="field" use="optional" type="xsd:unsignedInt"/>
///     <attribute name="count" type="xsd:unsignedInt"/>
///     <attribute name="selected" type="xsd:boolean" default="true"/>
///     <attribute name="byPosition" type="xsd:boolean" default="false"/>
///     <attribute name="relative" type="xsd:boolean" default="false"/>
///     <attribute name="defaultSubtotal" type="xsd:boolean" default="false"/>
///     <attribute name="sumSubtotal" type="xsd:boolean" default="false"/>
///     <attribute name="countASubtotal" type="xsd:boolean" default="false"/>
///     <attribute name="avgSubtotal" type="xsd:boolean" default="false"/>
///     <attribute name="maxSubtotal" type="xsd:boolean" default="false"/>
///     <attribute name="minSubtotal" type="xsd:boolean" default="false"/>
///     <attribute name="productSubtotal" type="xsd:boolean" default="false"/>
///     <attribute name="countSubtotal" type="xsd:boolean" default="false"/>
///     <attribute name="stdDevSubtotal" type="xsd:boolean" default="false"/>
///     <attribute name="stdDevPSubtotal" type="xsd:boolean" default="false"/>
///     <attribute name="varSubtotal" type="xsd:boolean" default="false"/>
///     <attribute name="varPSubtotal" type="xsd:boolean" default="false"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `field`: The field index (`field`).
/// - `count`: The count of references (`count`).
/// - `selected`: Indicates whether the field is selected (`selected`).
/// - `by_position`: Indicates whether the selection is by position (`byPosition`).
/// - `relative`: Indicates whether the selection is relative (`relative`).
/// - `include_default_filter`: Indicates whether to include the default subtotal filter (`defaultSubtotal`).
/// - `include_sum_aggregate_filter`: Indicates whether to include the sum subtotal filter (`sumSubtotal`).
/// - `include_counta_filter`: Indicates whether to include the countA subtotal filter (`countASubtotal`).
/// - `include_avg_aggregate_filter`: Indicates whether to include the average subtotal filter (`avgSubtotal`).
/// - `include_max_aggregate_filter`: Indicates whether to include the max subtotal filter (`maxSubtotal`).
/// - `include_min_aggregate_filter`: Indicates whether to include the min subtotal filter (`minSubtotal`).
/// - `include_prod_aggregate_filter`: Indicates whether to include the product subtotal filter (`productSubtotal`).
/// - `include_count_filter`: Indicates whether to include the count subtotal filter (`countSubtotal`).
/// - `include_std_deviation_filter`: Indicates whether to include the standard deviation subtotal filter (`stdDevSubtotal`).
/// - `include_pop_std_deviation_filter`: Indicates whether to include the population standard deviation subtotal filter (`stdDevPSubtotal`).
/// - `include_variance_filter`: Indicates whether to include the variance subtotal filter (`varSubtotal`).
/// - `include_pop_variance_filter`: Indicates whether to include the population variance subtotal filter (`varPSubtotal`).
/// - `selected_items`: A vector of `SelectedItem` elements, each representing an index (`x`).
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite)]
pub(crate) struct CTPivotAreaReference {
    #[xml(name = "field")]
    field: Vec<u8>,
    #[xml(name = "count")]
    count: Vec<u8>,
    #[xml(name = "selected", default_bool = true)]
    selected: bool,
    #[xml(name = "byPosition", default_bool = false)]
    by_position: bool,
    #[xml(name = "relative", default_bool = false)]
    relative: bool,
    #[xml(name = "defaultSubtotal", default_bool = false)]
    include_default_filter: bool,
    #[xml(name = "sumSubtotal", default_bool = false)]
    include_sum_aggregate_filter: bool,
    #[xml(name = "countASubtotal", default_bool = false)]
    include_counta_filter: bool,
    #[xml(name = "avgSubtotal", default_bool = false)]
    include_avg_aggregate_filter: bool,
    #[xml(name = "maxSubtotal", default_bool = false)]
    include_max_aggregate_filter: bool,
    #[xml(name = "minSubtotal", default_bool = false)]
    include_min_aggregate_filter: bool,
    #[xml(name = "productSubtotal", default_bool = false)]
    include_prod_aggregate_filter: bool,
    #[xml(name = "countSubtotal", default_bool = false)]
    include_count_filter: bool,
    #[xml(name = "stdDevSubtotal", default_bool = false)]
    include_std_deviation_filter: bool,
    #[xml(name = "stdDevPSubtotal", default_bool = false)]
    include_pop_std_deviation_filter: bool,
    #[xml(name = "varSubtotal", default_bool = false)]
    include_variance_filter: bool,
    #[xml(name = "varPSubtotal", default_bool = false)]
    include_pop_variance_filter: bool,

    #[xml(element)]
    selected_items: Vec<CTIndex>,
}
impl CTPivotAreaReference {
    /// Creates a new `CT_PivotAreaReference` instance with xml schema default values.
    pub(crate) fn new() -> Self {
        Self {
            ..Default::default()
        }
    }
}

/// Represents a collection of references within a `PivotTable` pivot area.
///
/// This struct corresponds to the `CT_PivotAreaReferences` complex type in the XML schema. It encapsulates
/// a count of references and a collection of individual `Reference` elements.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_PivotAreaReferences">
///     <sequence>
///         <element name="reference" maxOccurs="unbounded" type="CT_PivotAreaReference"/>
///     </sequence>
///     <attribute name="count" type="xsd:unsignedInt"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `count`: The number of references in the collection (`count`).
/// - `references`: A vector of `Reference` elements, each representing a pivot area reference (`reference`).
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite)]
pub(crate) struct CTPivotAreaReferences {
    #[xml(name = "count")]
    count: Vec<u8>,
    #[xml(element)]
    references: Vec<CTPivotAreaReference>,
}
impl CTPivotAreaReferences {
    /// Creates a new `CT_PivotAreaReferences` instance with xml schema default values.
    pub(crate) fn new() -> Self {
        Self {
            ..Default::default()
        }
    }
}

/// Represents a rule to describe `PivotTable` selection, defining the area and its properties.
///
/// This struct corresponds to the `CT_PivotArea` complex type in the XML schema. It encapsulates
/// attributes and elements that specify the pivot area, including its type, data, labels, and other
/// settings.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_PivotArea">
///     <sequence>
///         <element name="references" minOccurs="0" type="CT_PivotAreaReferences"/>
///         <element name="extLst" minOccurs="0" type="CT_ExtensionList"/>
///     </sequence>
///     <attribute name="field" use="optional" type="xsd:int"/>
///     <attribute name="type" type="ST_PivotAreaType" default="normal"/>
///     <attribute name="dataOnly" type="xsd:boolean" default="true"/>
///     <attribute name="labelOnly" type="xsd:boolean" default="false"/>
///     <attribute name="grandRow" type="xsd:boolean" default="false"/>
///     <attribute name="grandCol" type="xsd:boolean" default="false"/>
///     <attribute name="cacheIndex" type="xsd:boolean" default="false"/>
///     <attribute name="outline" type="xsd:boolean" default="true"/>
///     <attribute name="offset" type="ST_Ref"/>
///     <attribute name="collapsedLevelsAreSubtotals" type="xsd:boolean" default="false"/>
///     <attribute name="axis" type="ST_Axis" use="optional"/>
///     <attribute name="fieldPosition" type="xsd:unsignedInt" use="optional"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `field`: The field index (`field`).
/// - `use_pivot_type`: The type of pivot area (`type`).
/// - `use_data_only`: Indicates whether only data is included (`dataOnly`).
/// - `use_label_only`: Indicates whether only labels are included (`labelOnly`).
/// - `include_row_total`: Indicates whether to include row totals (`grandRow`).
/// - `include_col_total`: Indicates whether to include column totals (`grandCol`).
/// - `cache_index`: Indicates whether to use the cache index (`cacheIndex`).
/// - `outline`: Indicates whether to include outlines (`outline`).
/// - `offset`: The offset reference (`offset`).
/// - `collapsed_are_subtotal`: Indicates whether collapsed levels are subtotals (`collapsedLevelsAreSubtotals`).
/// - `axis`: The axis of the pivot area (`axis`).
/// - `field_pos`: The field position (`fieldPosition`).
/// - `reference_collection`: The collection of references (`references`).
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite)]
pub(crate) struct CTPivotArea {
    #[xml(name = "field")]
    field: Vec<u8>,
    #[xml(name = "type", default_bytes = b"normal")]
    pivot_type: Vec<u8>,
    #[xml(name = "dataOnly", default_bool = true)]
    use_data_only: bool,
    #[xml(name = "labelOnly", default_bool = false)]
    use_label_only: bool,
    #[xml(name = "grandRow", default_bool = false)]
    include_row_total: bool,
    #[xml(name = "grandCol", default_bool = false)]
    include_col_total: bool,
    #[xml(name = "cacheIndex", default_bool = false)]
    cache_index: bool,
    #[xml(name = "outline", default_bool = true)]
    outline: bool,
    #[xml(name = "offset")]
    offset: Vec<u8>,
    #[xml(name = "collapsedLevelsAreSubtotals", default_bool = false)]
    collapsed_are_subtotal: bool,
    #[xml(name = "axis")]
    axis: Vec<u8>,
    #[xml(name = "fieldPosition")]
    field_pos: Vec<u8>,

    #[xml(element)]
    reference_collection: CTPivotAreaReferences,
}
impl CTPivotArea {
    /// Creates a new `CT_PivotArea` instance with xml schema default values.
    pub(crate) fn new() -> Self {
        Self {
            outline: true,
            ..Default::default()
        }
    }
}

/// Represents a selection within a `PivotTable`, defining the active row, column, and other settings.
///
/// This struct corresponds to the `CT_PivotSelection` complex type in the XML schema. It encapsulates
/// attributes that specify the active row, column, pane, and other properties related to the selection
/// within a PivotTable.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_PivotSelection">
///     <sequence>
///         <element name="pivotArea" type="CT_PivotArea"/>
///     </sequence>
///     <attribute name="pane" type="ST_Pane" use="optional" default="topLeft"/>
///     <attribute name="showHeader" type="xsd:boolean" default="false"/>
///     <attribute name="label" type="xsd:boolean" default="false"/>
///     <attribute name="data" type="xsd:boolean" default="false"/>
///     <attribute name="extendable" type="xsd:boolean" default="false"/>
///     <attribute name="count" type="xsd:unsignedInt" default="0"/>
///     <attribute name="axis" type="ST_Axis" use="optional"/>
///     <attribute name="dimension" type="xsd:unsignedInt" default="0"/>
///     <attribute name="start" type="xsd:unsignedInt" default="0"/>
///     <attribute name="min" type="xsd:unsignedInt" default="0"/>
///     <attribute name="max" type="xsd:unsignedInt" default="0"/>
///     <attribute name="activeRow" type="xsd:unsignedInt" default="0"/>
///     <attribute name="activeCol" type="xsd:unsignedInt" default="0"/>
///     <attribute name="previousRow" type="xsd:unsignedInt" default="0"/>
///     <attribute name="previousCol" type="xsd:unsignedInt" default="0"/>
///     <attribute name="click" type="xsd:unsignedInt" default="0"/>
///     <attribute ref="r:id" use="optional"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `pane`: The pane in which the selection is active (`pane`).
/// - `show_header`: Indicates whether to show the header (`showHeader`).
/// - `label`: Indicates whether the selection is a label (`label`).
/// - `data`: Indicates whether the selection is data (`data`).
/// - `extendable`: Indicates whether the selection is extendable (`extendable`).
/// - `count`: The count of selected items (`count`).
/// - `axis`: The axis of the selection (`axis`).
/// - `dimension`: The dimension of the selection (`dimension`).
/// - `start`: The starting index of the selection (`start`).
/// - `min`: The minimum value of the selection (`min`).
/// - `max`: The maximum value of the selection (`max`).
/// - `row`: The active row in the selection (`activeRow`).
/// - `col`: The active column in the selection (`activeCol`).
/// - `prev_row`: The previous active row (`previousRow`).
/// - `prev_col`: The previous active column (`previousCol`).
/// - `click`: The click count (`click`).
/// - `rid`: The relationship ID (`r:id`).
/// - `area`: The pivot area of the selection (`pivotArea`).
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite)]
pub(crate) struct CTPivotSelection {
    #[xml(name = "pane", default_bytes = b"topLeft")]
    pane: Vec<u8>,
    #[xml(name = "showHeader", default_bool = false)]
    show_header: bool,
    #[xml(name = "label", default_bool = false)]
    label: bool,
    #[xml(name = "data", default_bool = false)]
    data: bool,
    #[xml(name = "extendable", default_bool = false)]
    extendable: bool,
    #[xml(name = "count", default_bytes = b"0")]
    count: Vec<u8>,
    #[xml(name = "axis")]
    axis: Vec<u8>,
    #[xml(name = "dimension", default_bytes = b"0")]
    dimension: Vec<u8>,
    #[xml(name = "start", default_bytes = b"0")]
    start: Vec<u8>,
    #[xml(name = "min", default_bytes = b"0")]
    min: Vec<u8>,
    #[xml(name = "max", default_bytes = b"0")]
    max: Vec<u8>,
    #[xml(name = "activeRow", default_bytes = b"0")]
    row: Vec<u8>,
    #[xml(name = "activeCol", default_bytes = b"0")]
    col: Vec<u8>,
    #[xml(name = "previousRow", default_bytes = b"0")]
    prev_row: Vec<u8>,
    #[xml(name = "previousCol", default_bytes = b"0")]
    prev_col: Vec<u8>,
    #[xml(name = "click", default_bytes = b"0")]
    click: Vec<u8>,
    #[xml(name = "r:id")]
    rid: Vec<u8>,

    #[xml(element)]
    area: CTPivotArea,
}
impl CTPivotSelection {
    /// Creates a new `CT_PivotSelection` instance with xml schema default values.
    pub(crate) fn new() -> Self {
        Self {
            ..Default::default()
        }
    }
}

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
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite)]
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

    #[xml(element)]
    pane: Option<Pane>,
    #[xml(element)]
    selection: Option<Selection>,
    #[xml(element)]
    pivot_selection: Option<CTPivotSelection>,
}
impl CTSheetView {
    /// Creates a new `CT_SheetView` instance with xml schema default values.
    fn new(id: u32) -> Self {
        Self {
            show_grid: true,
            show_zero: true,
            show_outline_symbol: true,
            zoom_scale: Zoom::Z100.into(),
            view_id: id.to_string().as_bytes().into(),
            show_ruler: true,
            show_header: true,
            show_whitespace: true,
            ..Default::default()
        }
    }
}

/// Represents the properties of a worksheet, including synchronization, transitions, and formatting.
///
/// This struct corresponds to the `CT_SheetPr` complex type in the XML schema. It encapsulates
/// attributes and elements that define the behavior and appearance of a worksheet.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_SheetPr">
///     <sequence>
///         <element name="tabColor" type="CT_Color" minOccurs="0" maxOccurs="1"/>
///         <element name="outlinePr" type="CT_OutlinePr" minOccurs="0" maxOccurs="1"/>
///         <element name="pageSetUpPr" type="CT_PageSetUpPr" minOccurs="0" maxOccurs="1"/>
///     </sequence>
///     <attribute name="syncHorizontal" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="syncVertical" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="syncRef" type="ST_Ref" use="optional"/>
///     <attribute name="transitionEvaluation" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="transitionEntry" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="published" type="xsd:boolean" use="optional" default="true"/>
///     <attribute name="codeName" type="xsd:string" use="optional"/>
///     <attribute name="filterMode" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="enableFormatConditionsCalculation" type="xsd:boolean" use="optional" default="true"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `sync_horizontal`: Indicates whether horizontal synchronization is enabled (`syncHorizontal`).
/// - `sync_vertical`: Indicates whether vertical synchronization is enabled (`syncVertical`).
/// - `sync_ref`: The reference for synchronization (`syncRef`).
/// - `transition_eval`: Indicates whether transition evaluation is enabled (`transitionEvaluation`).
/// - `transition_entry`: Indicates whether transition entry is enabled (`transitionEntry`).
/// - `published`: Indicates whether the sheet is published (`published`).
/// - `code_name`: The code name of the sheet (`codeName`).
/// - `filter_mode`: Indicates whether filter mode is enabled (`filterMode`).
/// - `enable_cond_format_calc`: Indicates whether conditional formatting calculation is enabled (`enableFormatConditionsCalculation`).
/// - `tab_color`: The color of the sheet tab (`tabColor`).
/// - `outline_pr`: The outline properties of the sheet (`outlinePr`).
/// - `page_setup_pr`: The page setup properties of the sheet (`pageSetUpPr`).
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite)]
pub struct CTSheetPr {
    #[xml(name = "syncHorizontal", default_bool = false)]
    sync_horizontal: bool,
    #[xml(name = "syncVertical", default_bool = false)]
    sync_vertical: bool,
    #[xml(name = "syncRef")]
    sync_ref: Vec<u8>,
    #[xml(name = "transitionEvaluation", default_bool = false)]
    transition_eval: bool,
    #[xml(name = "transitionEntry", default_bool = false)]
    transition_entry: bool,
    #[xml(name = "published", default_bool = true)]
    published: bool,
    #[xml(name = "codeName")]
    code_name: Vec<u8>,
    #[xml(name = "filterMode", default_bool = false)]
    filter_mode: bool,
    #[xml(name = "enableFormatConditionsCalculation", default_bool = true)]
    enable_cond_format_calc: bool,

    #[xml(element)]
    tab_color: Option<Color>,
    #[xml(element)]
    outline_pr: Option<CTOutlinePr>,
    #[xml(element)]
    page_setup_pr: Option<CTPageSetupPr>,
}
impl CTSheetPr {
    /// Creates a new `CT_SheetPr` instance with xml schema default values.
    pub fn new() -> Self {
        Self {
            ..Default::default()
        }
    }
}

/// Represents the dimensions of a worksheet, defining the range of cells that contain data.
///
/// This struct corresponds to the `CT_SheetDimension` complex type in the XML schema. It encapsulates
/// a required attribute `ref` that specifies the cell range of the worksheet's dimensions.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_SheetDimension">
///     <attribute name="ref" type="ST_Ref" use="required"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `range`: The cell range of the worksheet's dimensions (`ref`).
#[derive(Debug, PartialEq, Default, Clone, Eq, XmlWrite)]
pub struct CTSheetDimension {
    #[xml(name = "ref")]
    range: Vec<u8>,
}
impl CTSheetDimension {
    /// Creates a new `CT_SheetDimension` instance with xml schema default values.
    pub fn new() -> Self {
        Self { range: "A1".into() }
    }
}

/// Represents the outline properties of a worksheet, defining how outlines are applied and displayed.
///
/// This struct corresponds to the `CT_OutlinePr` complex type in the XML schema. It encapsulates
/// attributes that control the application of styles, the position of summary rows and columns,
/// and the visibility of outline symbols.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_OutlinePr">
///     <attribute name="applyStyles" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="summaryBelow" type="xsd:boolean" use="optional" default="true"/>
///     <attribute name="summaryRight" type="xsd:boolean" use="optional" default="true"/>
///     <attribute name="showOutlineSymbols" type="xsd:boolean" use="optional" default="true"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `apply_styles`: Indicates whether styles are applied to the outline (`applyStyles`).
/// - `summary_below`: Indicates whether summary rows are displayed below the detail rows (`summaryBelow`).
/// - `summary_right`: Indicates whether summary columns are displayed to the right of the detail columns (`summaryRight`).
/// - `show_outline_symbols`: Indicates whether outline symbols are displayed (`showOutlineSymbols`).
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite)]
pub struct CTOutlinePr {
    #[xml(name = "applyStyles", default_bool = false)]
    apply_styles: bool,
    #[xml(name = "summaryBelow", default_bool = true)]
    summary_below: bool,
    #[xml(name = "summaryRight", default_bool = true)]
    summary_right: bool,
    #[xml(name = "showOutlineSymbols", default_bool = true)]
    show_outline_symbols: bool,
}

impl CTOutlinePr {
    /// Creates a new `CT_OutlinePr` instance with xml schema default values.
    pub fn new() -> Self {
        Self {
            ..Default::default()
        }
    }
}

#[derive(Debug, Default, PartialEq, Clone, Eq)]
pub struct Sheet {
    path: String,
    uid: Vec<u8>,
    code_name: Vec<u8>,
    fit_to_page: bool,
    auto_page_break: bool,
    dimensions: Vec<u8>,
    enable_cond_format_calc: bool,
    published: bool,
    sync_vertical: bool,
    sync_horizontal: bool,
    sync_ref: Vec<u8>,
    transition_eval: bool,
    transition_entry: bool,
    filter_mode: bool,
    apply_outline_style: bool,
    show_summary_below: bool, // summary row should be inserted to above when off
    show_summary_right: bool, // sumamry row should be inserted to left when off
    sheet_views: Vec<CTSheetView>,
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
                if !self.code_name.is_empty() {
                    attrs.push((b"codeName".as_ref(), self.code_name.as_ref()));
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
                if !self.sync_ref.is_empty() {
                    attrs.push((b"syncRef".as_ref(), self.sync_ref.as_ref()));
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
                    if !view.color_id.is_empty() {
                        attrs.push((b"colorId".as_ref(), view.color_id.as_ref()));
                    }
                    if !view.show_whitespace {
                        attrs.push((b"showWhiteSpace".as_ref(), b"0".as_ref()));
                    }
                    if view.view != b"normal" && !view.view.is_empty() {
                        attrs.push((b"view".as_ref(), view.view.as_ref()));
                    }
                    if !view.top_left_cell.is_empty() {
                        attrs.push((b"topLeftCell".as_ref(), view.top_left_cell.as_ref()));
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
            path: path.into(),
            show_summary_below: true,
            show_summary_right: true,
            auto_page_break: true,
            enable_cond_format_calc: true,
            published: true,
            show_outline_symbol: true,
            ..Default::default()
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
                                b"codeName" => self.code_name = a.value.into(),
                                b"enableFormatConditions" => {
                                    self.enable_cond_format_calc = *a.value == *b"1"
                                }
                                b"filterMode" => self.filter_mode = *a.value == *b"1",
                                b"published" => self.published = *a.value == *b"1",
                                b"syncHorizontal" => self.sync_horizontal = *a.value == *b"1",
                                b"syncVertical" => self.sync_vertical = *a.value == *b"1",
                                b"syncRef" => self.sync_ref = a.value.into(),
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
                                b"ref" => self.dimensions = a.value.into(),
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
                                            b"defaultGridColor" => {
                                                sheet_view.use_default_grid_color =
                                                    *a.value == *b"1";
                                            }
                                            b"view" => sheet_view.view = a.value.into(),
                                            b"topLeftCell" => {
                                                sheet_view.top_left_cell = a.value.into();
                                            }
                                            b"colorId" => {
                                                sheet_view.color_id = a.value.into();
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
                                                                pane.active_pane = a.value.into();
                                                            }
                                                            b"state" => {
                                                                pane.state = a.value.into();
                                                            }

                                                            b"topLeftCell" => {
                                                                pane.top_left_cell = a.value.into();
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
                                                            b"pane" => {
                                                                selection.pane = a.value.into()
                                                            }
                                                            b"activeCellId" => {
                                                                selection.cell_id = a.value.into();
                                                            }
                                                            b"activeCell" => {
                                                                selection.cell = a.value.into();
                                                            }
                                                            b"sqref" => {
                                                                selection.sqref = a.value.into();
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
                                                                pivot.pane = a.value.into();
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
                                                                pivot.rid = a.value.into();
                                                            }
                                                            b"axis" => pivot.axis = a.value.into(),

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
                                                                        b"field" =>
                                                                            area.field =a.value.to_vec(),


                                                                        b"type" => {
                                                                            area.pivot_type = a.value.into();
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
                                                                            area.offset = a.value.into();
                                                                        }
                                                                        // This binary string breaks rustfmt with line length
                                                                        b"collapsedLevelsAreSubtotals" => {
                                                                            area.collapsed_are_subtotal = *a.value == *b"1";
                                                                        }
                                                                        b"axis" => {
                                                                            area.axis = a.value.into();
                                                                        }
                                                                        b"fieldPosition" =>
                                                                            area.field_pos =
                                                                                a.value.to_vec(),


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
                                                            pivot.area.reference_collection =
                                                                References::new();
                                                            let mut reference = Reference::new();
                                                            for attr in e.attributes() {
                                                                if let Ok(a) = attr {
                                                                    match a.key.as_ref() {
                                                                        b"field" => {
                                                                            reference.field =
                                                                                a.value.into();
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
                                                                                                reference.selected_items.push(
                                                                                                    SelectedItem { item: a.value.into() }
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
                                                                            .reference_collection
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
                                                            let count = pivot
                                                                .area
                                                                .reference_collection
                                                                .references
                                                                .len();
                                                            pivot.area.reference_collection.count =
                                                                count.to_string().as_bytes().into();
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
