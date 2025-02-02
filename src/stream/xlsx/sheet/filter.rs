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
/// Represents a custom filter for a filter column in a spreadsheet.
/// This struct corresponds to the `CT_CustomFilter` complex type in the XML schema.
/// It allows for the application of custom filters with a specified operator and value.
/// 
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_CustomFilter">
///     <attribute name="operator" type="ST_FilterOperator" default="equal" use="optional"/>
///     <attribute name="val" type="ST_Xstring"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `operator`: The operator used for the filter (`ST_FilterOperator`), defaulting to `"equal"` if not specified.
/// - `val`: The value for the filter (`ST_Xstring`), which is required to apply the filter.
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
struct CTCustomFilter {
    operator: Vec<u8>
    val: Vec<u8>,
}
impl CTCustomFilter {
    /// Creates a new `CT_CustomFilter` with the xml schema default values.
    fn new() -> Self {
        CTCustomFilter {
            operator: b"equal".into(),
            ..Default::default()
        }
    }
}
/// Represents custom filters for a filter column in a spreadsheet.
/// This struct corresponds to the `CT_CustomFilters` complex type in the XML schema.
/// It allows users to apply custom filters to the column, with the option to specify whether the filters are combined with an AND logic.
/// 
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_CustomFilters">
///     <sequence>
///         <element name="customFilter" type="CT_CustomFilter" minOccurs="1" maxOccurs="2"/>
///     </sequence>
///     <attribute name="and" type="xsd:boolean" use="optional" default="false"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `custom_filters`: A list of custom filters applied to the column (`customFilter`), with a minimum of 1 and a maximum of 2 allowed.
/// - `and`: Whether the filters are combined using an AND logic (`false` by default).
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
struct CTCustomFilters {
    custom_filters: Vec<CTCustomFilter>,
    and_logic: bool,
}

impl CTCustomFilters {
    /// Creates a new `CT_CustomFilters` with xml schema default values (`and_logic` set to `false`).
    fn new() -> Self {
        CTCustomFilters {
          ..Default::default(),
        }
    }
}
/// Represents a color filter for a filter column in a spreadsheet.
/// This struct corresponds to the `CT_ColorFilter` complex type in the XML schema.
/// It allows for filtering based on cell color and also includes a reference to a custom style (dxfId).
/// 
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_ColorFilter">
///     <attribute name="dxfId" type="ST_DxfId" use="optional"/>
///     <attribute name="cellColor" type="xsd:boolean" use="optional" default="true"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `dxf_id`: A reference to a custom style, allowing users to specify a style that is part of a predefined set of styles.
/// - `cell_color`: Whether the filter should be based on the cell's background color (`true` by default).
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
struct CTColorFilter {
    dxf_id: Vec<u8>,
    cell_color: bool,
}

impl CTColorFilter {
    /// Creates a new `CT_ColorFilter` with xml schema default value.
    fn new() -> Self {
        Self {
            cell_color: true,
            ..Default::default()
        }
    }
}
/// Enum representing the different filter types that can be applied to a filter column.
/// This corresponds to the `<choice>` element in the XML schema for `CT_FilterColumn`.
///
/// The enum ensures that only one filter type is applied at a time, reflecting the choice 
/// structure in the XML schema. Each variant in the enum represents a possible filter option
/// that can be associated with the column. These options include standard filters, top 10 filters, 
/// custom filters, dynamic filters, color filters, icon filters, and extensions.
///
/// # Variants
/// - `Filters(CTFilters)`: Represents a standard filter for the column, mapped to the `filters` element.
/// - `Top10(CTTop10)`: Represents a "Top 10" filter for the column, mapped to the `top10` element.
/// - `CustomFilters(CTCustomFilters)`: Represents custom filters applied to the column, mapped to the `customFilters` element.
/// - `DynamicFilter(CTDynamicFilter)`: Represents a dynamic filter for the column, mapped to the `dynamicFilter` element.
/// - `ColorFilter(CTColorFilter)`: Represents a color-based filter for the column, mapped to the `colorFilter` element.
/// - `IconFilter(CTIconFilter)`: Represents an icon-based filter for the column, mapped to the `iconFilter` element.
///
/// The choice constraint in the XML schema guarantees that only one of these variants is applied to a
/// column at any given time, making it important to handle the selection of filters accordingly.
#[derive(Debug, Clone, PartialEq, XmlRead, XmlWrite)]
enum Filter {
    /// Represents a standard filter applied to the column, based on cell values.
    Filters(CTFilters),
    /// Represents a filter that selects the top 10 items from the column, typically numeric values.
    Top10(CTTop10),
    /// Represents a custom filter applied to the column, allowing advanced filtering logic.
    CustomFilters(CTCustomFilters),
    /// Represents a dynamic filter, which is often based on changing or conditional data.
    DynamicFilter(CTDynamicFilter),
    /// Represents a filter based on cell color, useful for visually distinguishing data.
    ColorFilter(CTColorFilter),
    /// Represents a filter based on icons, used to group or categorize data visually.
    IconFilter(CTIconFilter),
}
/// Represents a filter column in a spreadsheet, defining various filter options and settings.
///
/// This struct corresponds to the `CT_FilterColumn` complex type in the XML schema. It encapsulates
/// attributes for filter settings, button visibility, and allows for different filter options (filters, 
/// top 10, custom filters, dynamic filter, color filter, and icon filter).
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_FilterColumn">
///     <choice minOccurs="0" maxOccurs="1">
///         <element name="filters" type="CT_Filters" minOccurs="0" maxOccurs="1"/>
///         <element name="top10" type="CT_Top10" minOccurs="0" maxOccurs="1"/>
///         <element name="customFilters" type="CT_CustomFilters" minOccurs="0" maxOccurs="1"/>
///         <element name="dynamicFilter" type="CT_DynamicFilter" minOccurs="0" maxOccurs="1"/>
///         <element name="colorFilter" type="CT_ColorFilter" minOccurs="0" maxOccurs="1"/>
///         <element name="iconFilter" type="CT_IconFilter" minOccurs="0" maxOccurs="1"/>
///         <element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
///     </choice>
///     <attribute name="colId" type="xsd:unsignedInt" use="required"/>
///     <attribute name="hiddenButton" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="showButton" type="xsd:boolean" use="optional" default="true"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `col_id`: The column ID for the filter column (`colId`).
/// - `hidden_button`: Whether the button for the column is hidden (`hiddenButton`).
/// - `show_button`: Whether the button for the column is visible (`showButton`).
/// - `filter`: The filter type for the column, which can be one of the `Filter` options
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
struct CTFilterColumn {
    col_id: Vec<u8>,
    hidden_button: bool,
    show_button: bool,
    filter: Option<Filter>,  
}
impl CTFilterColumn {
    /// Creates a new `CT_FilterColumn` with xml schema default values.
    fn new() -> Self {
        CTFilterColumn {
            show_button: true,
            ..Default::default()
        }
    }
}
/// Represents the condition for sorting in a document.
///
/// This struct corresponds to the `CT_SortCondition` complex type in the XML schema.
/// It contains the attributes that define how sorting is performed, including the reference,
/// sort type, and other optional settings like descending order, custom list, and icon set.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_SortCondition">
///     <attribute name="descending" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="sortBy" type="ST_SortBy" use="optional" default="value"/>
///     <attribute name="ref" type="ST_Ref" use="required"/>
///     <attribute name="customList" type="ST_Xstring" use="optional"/>
///     <attribute name="dxfId" type="ST_DxfId" use="optional"/>
///     <attribute name="iconSet" type="ST_IconSetType" use="optional" default="3Arrows"/>
///     <attribute name="iconId" type="xsd:unsignedInt" use="optional"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `descending`: Indicates whether the sort is in descending order. Defaults to `false`.
/// - `sort_by`: Specifies how to sort the data. Defaults to `"value"`.
/// - `reference`: The reference for the range to be sorted. This is required.
/// - `custom_list`: Specifies a custom list for sorting. Optional.
/// - `dxf_id`: Applies the style for sorting. Optional.
/// - `icon_set`: Specifies an icon set for sorting. Defaults to `"3Arrows"`.
/// - `icon_id`: The ID for the icon to be applied. Optional.
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
struct CTSortCondition {
    /// Whether sorting is in descending order.
    descending: bool,
    /// The cell range reference for the sort condition.
    reference: Vec<u8>,
    /// Defines sorting criteria.
    sort_by: Vec<u8>,
    /// Custom sorting order if specified.
    custom_list: Vec<u8>,
    /// Formatting ID for differential styles.
    dxf_id: Vec<u8>,
    /// Icon set used for sorting.
    icon_set: Vec<u8>,
    /// Specific icon index within the icon set.
    icon_id: Vec<u8>,
}
impl CTSortCondition {
    /// Creates a new `CT_SortCondition` with xml schema default values.
    fn new() -> Self {
        Self {
            sort_by: b"value".into()
            icon_set: b"3Arrows".into()
            ..Default::default()
        }
    }
}
/// Represents the sort state in a document.
///
/// This struct corresponds to the `CT_SortState` complex type in the XML schema.
/// It contains the sorting conditions, optional extensions, and attributes related to sorting.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_SortState">
///     <sequence>
///         <element name="sortCondition" minOccurs="0" maxOccurs="64" type="CT_SortCondition"/>
///         <element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
///     </sequence>
///     <attribute name="columnSort" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="caseSensitive" type="xsd:boolean" use="optional" default="false"/>
///     <attribute name="sortMethod" type="ST_SortMethod" use="optional" default="none"/>
///     <attribute name="ref" type="ST_Ref" use="required"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `column_sort`: Indicates whether column sorting is enabled. Defaults to `false`.
/// - `case_sensitive`: Indicates whether sorting is case-sensitive. Defaults to `false`.
/// - `sort_method`: Specifies the sorting method used. Defaults to `"none"`.
/// - `reference`: The reference for the range to be sorted. This is required.
/// - `sort_conditions`: A list of sorting conditions. Can have up to 64 conditions.
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
pub(crate) struct CTSortState {
    column_sort: bool,
    case_sensitive: bool,
    sort_method: Vec<u8>,
    reference: Vec<u8>,
    sort_conditions: Vec<CTSortCondition>,
}
impl CTSortState {
    /// Creates a new `CT_SortState` with xml schema default values.
    fn new() -> Self {
        Self {
            sort_method: b"none".into(),
            ..Default::default()
        }
    }
}
/// Represents an auto filter configuration in a document.
///
/// This struct corresponds to the `CT_AutoFilter` complex type in the XML schema.
/// It contains the filter columns, sorting state, and optional extensions related to the auto filter.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_AutoFilter">
///     <sequence>
///         <element name="filterColumn" minOccurs="0" maxOccurs="unbounded" type="CT_FilterColumn"/>
///         <element name="sortState" minOccurs="0" maxOccurs="1" type="CT_SortState"/>
///         <element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
///     </sequence>
///     <attribute name="ref" type="ST_Ref"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `reference`: The reference for the range of the filter. This is required.
/// - `filter_column`: A list of filter columns. Can be unbounded.
/// - `sort_state`: The sorting state for the filter. Optional.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
struct CTAutoFilter {
    reference: Vec<u8>,
    filter_column: Vec<CTFilterColumn>,
    sort_state: Option<CTSortState>,
}
impl CTAutoFilter {
    fn new() -> Self {
        Self {
            ..Default::default()
        }
    }
}