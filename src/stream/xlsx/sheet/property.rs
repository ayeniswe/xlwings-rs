use crate::{
    errors::XlsxError,
    stream::utils::{XmlReader, XmlWriter},
};
use derive::{XmlRead, XmlWrite};
use quick_xml::{
    events::{Event},
    Reader, Writer,
};
use std::io::BufRead;

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
            summary_below: true,
            summary_right: true,
            show_outline_symbols: true,
            ..Default::default()
        }
    }
}
/// Represents the page setup properties of a worksheet, defining how the worksheet is paginated.
///
/// This struct corresponds to the `CT_PageSetUpPr` complex type in the XML schema. It encapsulates
/// attributes that control automatic page breaks and whether the content should be fit to the page.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_PageSetUpPr">
///     <attribute name="autoPageBreaks" type="xsd:boolean" use="optional" default="true"/>
///     <attribute name="fitToPage" type="xsd:boolean" use="optional" default="false"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `auto_page_breaks`: Indicates whether automatic page breaks are enabled (`autoPageBreaks`).
/// - `fit_to_page`: Indicates whether the content should be fit to the page (`fitToPage`).
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite)]
pub struct CTPageSetupPr {
    #[xml(name = "autoPageBreaks", default_bool = true)]
    auto_page_breaks: bool,
    #[xml(name = "fitToPage", default_bool = false)]
    fit_to_page: bool,
}
impl CTPageSetupPr {
    /// Creates a new `CT_PageSetupPr` instance with xml schema default values.
    pub fn new() -> Self {
        Self {
            auto_page_breaks: true,
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
            published: true,
            enable_cond_format_calc: true,
            ..Default::default()
        }
    }
}