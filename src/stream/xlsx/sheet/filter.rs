use crate::stream::{utils::{XmlReader, XmlWriter}, xlsx::errors::XlsxError};
use derive::{XmlRead, XmlWrite};
use quick_xml::{events::Event, Reader, Writer};
use std::io::{BufRead, Write};

/// Represents the valid calendar types.
///
/// This enum corresponds to the `ST_CalendarType` simple type in the XML schema.
/// It provides the valid options for different calendar types.
///
/// # XML Schema Mapping
/// ```xml
/// <simpleType name="ST_CalendarType">
///   <restriction base="xsd:string">
///     <enumeration value="none"/>
///     <enumeration value="gregorian"/>
///     <enumeration value="gregorianUs"/>
///     <enumeration value="japan"/>
///     <enumeration value="taiwan"/>
///     <enumeration value="korea"/>
///     <enumeration value="hijri"/>
///     <enumeration value="thai"/>
///     <enumeration value="hebrew"/>
///     <enumeration value="gregorianMeFrench"/>
///     <enumeration value="gregorianArabic"/>
///     <enumeration value="gregorianXlitEnglish"/>
///     <enumeration value="gregorianXlitFrench"/>
///   </restriction>
/// </simpleType>
/// ```
#[derive(Debug, Default, Clone, PartialEq, EnumToBytes)]
#[camel]
pub enum STCalendarType {
    #[default]
    None,
    Gregorian,
    GregorianUs,
    Japan,
    Taiwan,
    Korea,
    Hijri,
    Thai,
    Hebrew,
    GregorianMeFrench,
    GregorianArabic,
    GregorianXlitEnglish,
    GregorianXlitFrench,
}
/// Represents the valid date-time grouping options.
///
/// This enum corresponds to the `ST_DateTimeGrouping` simple type in the XML schema.
/// It can be used to specify how date and time values should be grouped.
///
/// # XML Schema Mapping
/// ```xml
/// <simpleType name="ST_DateTimeGrouping">
///   <restriction base="xsd:string">
///     <enumeration value="year"/>
///     <enumeration value="month"/>
///     <enumeration value="day"/>
///     <enumeration value="hour"/>
///     <enumeration value="minute"/>
///     <enumeration value="second"/>
///   </restriction>
/// </simpleType>
/// ```
#[derive(Debug, Clone, PartialEq, EnumToBytes)]
#[camel]
pub enum STDateTimeGrouping {
    Year,
    Month,
    Day,
    Hour,
    Minute,
    Second,
}
/// Represents a date and time group item.
///
/// This struct corresponds to the `CT_DateGroupItem` complex type in the XML schema.
/// It contains attributes for year, month, day, hour, minute, second, and dateTimeGrouping.
///
/// # XML Schema Mapping
/// ```xml
/// <complexType name="CT_DateGroupItem">
///   <attribute name="year" type="xsd:unsignedShort" use="required"/>
///   <attribute name="month" type="xsd:unsignedShort" use="optional"/>
///   <attribute name="day" type="xsd:unsignedShort" use="optional"/>
///   <attribute name="hour" type="xsd:unsignedShort" use="optional"/>
///   <attribute name="minute" type="xsd:unsignedShort" use="optional"/>
///   <attribute name="second" type="xsd:unsignedShort" use="optional"/>
///   <attribute name="dateTimeGrouping" type="ST_DateTimeGrouping" use="required"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `year`: The year part of the date.
/// - `month`: The month part of the date.
/// - `day`: The day part of the date.
/// - `hour`: The hour part of the time.
/// - `minute`: The minute part of the time.
/// - `second`: The second part of the time.
/// - `date_time_grouping`: How the date and time are grouped.
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
pub(crate) struct CTDateGroupItem {
    year: Vec<u8>,
    month: Option<Vec<u8>>,
    day: Option<Vec<u8>>,
    hour: Option<Vec<u8>>,
    minute: Option<Vec<u8>>,
    second: Option<Vec<u8>>,
    date_time_grouping: Vec<u8>,
}
impl CTDateGroupItem {
    /// Creates a new `CT_DateGroupItem` with XML schema default values.
    pub fn new(year: u16, month: u8, day: u8, hour: u8, minute: u8, second: u16, date_time_grouping: STDateTimeGrouping) -> Self {
        Self {
            month: month.to_string().into(),
            day: day.to_string().into(),
            hour: hour.to_string().into(),
            minute: minute.to_string().into(),
            second: second.to_string().into(),
            year: year.to_string().into(),
            date_time_grouping: date_time_grouping.into()
        }
    }
}
/// Represents a filter with a string value.
///
/// This struct corresponds to the `CT_Filter` complex type in the XML schema.
/// It contains a single attribute `val` that represents the value of the filter.
///
/// # XML Schema Mapping
/// ```xml
/// <complexType name="CT_Filter">
///   <attribute name="val" type="ST_Xstring"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `val`: The value of the filter (a string).
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
pub(crate) struct CTFilter {
    val: Vec<u8>,
}
impl CTFilter {
    /// Creates a new `CT_Filter`  with XML schema default values.
    pub fn new(val: &str) -> Self {
        Self {
            val: val.into()
        }
    }
}
/// Represents a collection of filters and date group items.
///
/// This struct corresponds to the `CT_Filters` complex type in the XML schema.
/// It holds both regular filters and date-related group items
///
/// # XML Schema Mapping
/// ```xml
/// <complexType name="CT_Filters">
///   <sequence>
///     <element name="filter" type="CT_Filter" minOccurs="0" maxOccurs="unbounded"/>
///     <element name="dateGroupItem" type="CT_DateGroupItem" minOccurs="0" maxOccurs="unbounded"/>
///   </sequence>
///   <attribute name="blank" type="xsd:boolean" use="optional" default="false"/>
///   <attribute name="calendarType" type="ST_CalendarType" use="optional" default="none"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `filters`: A collection of regular filters.
/// - `date_group_items`: A collection of date group items.
/// - `blank`: A boolean indicating whether the blank option is enabled.
/// - `calendar_type`: The type of calendar used.
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
pub(crate) struct CTFilters {
    #[xml(default_bool = false)]
    blank: Option<bool>,
    #[xml(default_bytes = b"none")]
    calendar_type: Option<Vec<u8>>,
    #[xml(following_elements, sequence)]
    filters: Vec<CTFilter>,
    date_group_items: Vec<CTDateGroupItem>,
}
impl CTFilters {
    /// Creates a new `CT_Filters` instance with XML schema default values.
    pub fn new(blank: Option<bool>, calendar_type: Option<STCalendarType>, filters: Option<Vec<CTFilter>>, date_group_items: Option<Vec<CTDateGroupItem>>) -> Self {
        Self {
            blank: blank.unwrap_or(Some(false)),
            filters: filters.unwrap_or(Vec::new()),
            date_group_items: date_group_items.unwrap_or(Vec::new()),
            calendar_type: calendar_type.unwrap_or(STCalendarType::None).into(),
        }
    }
}
/// Represents the icon filter configuration.
///
/// This struct corresponds to the `CT_IconFilter` complex type in the XML schema.
/// It is used to define the icon set and optionally an icon ID for filtering.
///
/// # XML Schema Mapping
/// ```xml
/// <complexType name="CT_IconFilter">
///   <attribute name="iconSet" type="ST_IconSetType" use="required"/>
///   <attribute name="iconId" type="xsd:unsignedInt" use="optional"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `icon_set`: The icon set to use for the filter.
/// - `icon_id`: An optional icon ID within the icon set.
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
pub(crate) struct CTIconFilter {
    icon_set: Vec<u8>,
    icon_id: Option<Vec<u8>>,
}
impl CTIconFilter {
    /// Creates a new `CT_IconFilter` with XML schema default values.
    pub fn new(icon_id: Option<Vec<u8>>, icon_set: STIconSetType) -> Self {
        Self {
            icon_id,
            icon_set: icon_set.into(),
        }
    }
}
/// Represents the "Top 10" filter configuration.
///
/// This struct corresponds to the `CT_Top10` complex type in the XML schema.
/// It is used to define the top items to be selected, either based on absolute values or percentages.
///
/// # XML Schema Mapping
/// ```xml
/// <complexType name="CT_Top10">
///   <attribute name="top" type="xsd:boolean" use="optional" default="true"/>
///   <attribute name="percent" type="xsd:boolean" use="optional" default="false"/>
///   <attribute name="val" type="xsd:double" use="required"/>
///   <attribute name="filterVal" type="xsd:double" use="optional"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `top`: Whether to select the top values (`true`) or the bottom values (`false`).
/// - `percent`: Whether to treat the values as percentages.
/// - `val`: Top or bottom value to use as the filter criteria.
/// - `filter_val`: The actual cell value in the range which is used to perform the comparison for this filter.
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
struct CTTop10 {
    #[xml(default_bool = true)]
    top: Option<bool>,
    #[xml(default_bool = false)]
    percent: Option<bool>,
    val: Vec<u8>,
    filter_val: Option<Vec<u8>>,
}
impl CTTop10 {
    /// Creates a new `CT_Top10` with XML schema default values.
    fn new(top: Option<bool>, percent: Option<bool>, val: f32, filter_val: Option<f32>) -> Self {
        let filter_val = if let Some(v) = filter_val {
            Some(v.to_string().to_vec())
        } else {
            None
        }
        Self {
            top,
            percent,
            val: val.to_string().into(),
            filter_val,
        }
    }
}
/// Represents the type of dynamic filter to apply.
///
/// This enum corresponds to the `ST_DynamicFilterType` simple type in the XML schema.
/// It defines the possible types of dynamic filters.
///
/// # XML Schema Mapping
/// ```xml
/// <simpleType name="ST_DynamicFilterType">
///   <restriction base="xsd:string">
///     <enumeration value="null"/>
///     <enumeration value="aboveAverage"/>
///     <enumeration value="belowAverage"/>
///     <enumeration value="tomorrow"/>
///     <enumeration value="today"/>
///     <enumeration value="yesterday"/>
///     <enumeration value="nextWeek"/>
///     <enumeration value="thisWeek"/>
///     <enumeration value="lastWeek"/>
///     <enumeration value="nextMonth"/>
///     <enumeration value="thisMonth"/>
///     <enumeration value="lastMonth"/>
///     <enumeration value="nextQuarter"/>
///     <enumeration value="thisQuarter"/>
///     <enumeration value="lastQuarter"/>
///     <enumeration value="nextYear"/>
///     <enumeration value="thisYear"/>
///     <enumeration value="lastYear"/>
///     <enumeration value="yearToDate"/>
///     <enumeration value="Q1"/>
///     <enumeration value="Q2"/>
///     <enumeration value="Q3"/>
///     <enumeration value="Q4"/>
///     <enumeration value="M1"/>
///     <enumeration value="M2"/>
///     <enumeration value="M3"/>
///     <enumeration value="M4"/>
///     <enumeration value="M5"/>
///     <enumeration value="M6"/>
///     <enumeration value="M7"/>
///     <enumeration value="M8"/>
///     <enumeration value="M9"/>
///     <enumeration value="M10"/>
///     <enumeration value="M11"/>
///     <enumeration value="M12"/>
///   </restriction>
/// </simpleType>
/// ```
#[derive(Debug, Clone, PartialEq, EnumToBytes)]
pub enum STDynamicFilterType {
    #[camel] /// Represents no dynamic filter.
    Null,
    #[camel] /// Filters for values above the average.
    AboveAverage,
    #[camel] /// Filters for values below the average.
    BelowAverage,
    #[camel] /// Filters for tomorrow's values.
    Tomorrow,
    #[camel] /// Filters for today's values.
    Today,
    #[camel] /// Filters for yesterday's values.
    Yesterday,
    #[camel] /// Filters for the upcoming week.
    NextWeek,
    #[camel] /// Filters for this week's values.
    ThisWeek,
    #[camel] /// Filters for the previous week.
    LastWeek,
    #[camel] /// Filters for the upcoming month.
    NextMonth,
    #[camel] /// Filters for this month's values.
    ThisMonth,
    #[camel] /// Filters for the previous month.
    LastMonth,
    #[camel] /// Filters for the upcoming quarter.
    NextQuarter,
    #[camel] /// Filters for this quarter's values.
    ThisQuarter,
    #[camel] /// Filters for the previous quarter.
    LastQuarter,
    #[camel] /// Filters for the upcoming year.
    NextYear,
    #[camel] /// Filters for this year's values.
    ThisYear,
    #[camel] /// Filters for the previous year.
    LastYear,
    #[camel] /// Filters from the start of the year until the current date.
    YearToDate,
    #[camel] /// Represents the first quarter of the year.
    Q1,
    #[camel] /// Represents the second quarter of the year.
    Q2,
    #[camel] /// Represents the third quarter of the year.
    Q3,
    #[camel] /// Represents the fourth quarter of the year.
    Q4,
    #[camel] /// Represents the first month of the year (January).
    M1,
    #[camel] /// Represents the second month of the year (February).
    M2,
    #[camel] /// Represents the third month of the year (March).
    M3,
    #[camel] /// Represents the fourth month of the year (April).
    M4,
    #[camel] /// Represents the fifth month of the year (May).
    M5,
    #[camel] /// Represents the sixth month of the year (June).
    M6,
    #[camel] /// Represents the seventh month of the year (July).
    M7,
    #[camel] /// Represents the eighth month of the year (August).
    M8,
    #[camel] /// Represents the ninth month of the year (September).
    M9,
    #[camel] /// Represents the tenth month of the year (October).
    M10,
    #[camel] /// Represents the eleventh month of the year (November).
    M11,
    #[camel] /// Represents the twelfth month of the year (December).
    M12,
}
/// Represents the configuration for a dynamic filter.
/// This struct corresponds to the `CT_DynamicFilter` complex type in the XML schema.
/// It contains the filter type and optional values for filtering.
///
/// # XML Schema Mapping
/// The struct maps to the following XML schema definition:
/// ```xml
/// <complexType name="CT_DynamicFilter">
/// 	<attribute name="type" type="ST_DynamicFilterType" use="required"/>
/// 	<attribute name="val" type="xsd:double" use="optional"/>
/// 	<attribute name="maxVal" type="xsd:double" use="optional"/>
/// </complexType>
/// ```
///
/// # Fields
/// - `filter_type`: The type of dynamic filter to apply.
/// - `value`: The value to use for filtering.
/// - `max_value`: The maximum value for filtering.
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
struct CTDynamicFilter {
    filter_type: Vec<u8>,
    value: Option<Vec<u8>>,
    max_value: Option<Vec<u8>>,
}
impl CTDynamicFilter {
    /// Creates a new `CTDynamicFilter` with the xml schema default values.
    fn new(filter_type: STDynamicFilterType, max_value: Option<f32>, value: Option<f32>) -> Self {
        let value = if let Some(v) = value {
            Some(v.to_string().to_vec())
        } else {
            None
        }
        let max_value = if let Some(v) = max_value {
            Some(v.to_string().to_vec())
        } else {
            None
        }
        Self {
            filter_type,
            max_value,
            value,
        }
    }
}
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
    operator: Option<Vec<u8>>,
    val: Vec<u8>,
}
impl CTCustomFilter {
    /// Creates a new `CT_CustomFilter` with the xml schema default values.
    fn new(val: &str, operator: Option<STFilterOperator>) -> Self {
        CTCustomFilter {
            operator: operator.unwrap_or(STFilterOperator::Equal).into(),
            val: val.into()
        }
    }
}
/// Represents the filter operators used in SpreadsheetML for filtering data.
///
/// This enum corresponds to the `ST_FilterOperator` simple type in the
/// Office Open XML specification. Each variant specifies a type of comparison
/// operation that can be applied during data filtering.
#[derive(Debug, Clone, PartialEq, EnumToBytes)]
#[camel]
enum FilterOperator {
    Equal,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
    NotEqual,
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
    #[xml(default_bool = false)]
    and_logic: Option<bool>,
    #[xml(element)]
    custom_filters: Vec<CTCustomFilter>,
}
impl CTCustomFilters {
    /// Creates a new `CT_CustomFilters` with xml schema default values.
    fn new(and_logic: Option<bool>, custom_filters: Vec<CTCustomFilter>) -> Self {
        CTCustomFilters {
            and_logic,
            custom_filters
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
#[derive(Debug, Clone, PartialEq, XmlRead)]
enum Filter {
    /// Represents a standard filter applied to the column, based on cell values.
    #[xml(name = "filters")]
    Filters(CTFilters),
    /// Represents a filter that selects the top 10 items from the column, typically numeric values.
    #[xml(name = "top10")]
    Top10(CTTop10),
    /// Represents a custom filter applied to the column, allowing advanced filtering logic.
    #[xml(name = "customFilters")]
    CustomFilters(CTCustomFilters),
    /// Represents a dynamic filter, which is often based on changing or conditional data.
    #[xml(name = "dynamicFilter")]
    DynamicFilter(CTDynamicFilter),
    /// Represents a filter based on cell color, useful for visually distinguishing data.
    #[xml(name = "colorFilter")]
    ColorFilter(CTColorFilter),
    /// Represents a filter based on icons, used to group or categorize data visually.
    #[xml(name = "iconFilter")]
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
#[derive(Debug, Default, Clone, PartialEq, XmlRead)]
struct CTFilterColumn {
    col_id: Vec<u8>,
    hidden_button: bool,
    show_button: bool,
    #[xml(element)]
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
/// Represents the method by which sorting is applied in a document.
///
/// This enum corresponds to the `ST_SortBy` simple type in the XML schema.
/// It defines the possible options for sorting by values, cell color, font color, or icon.
///
/// # XML Schema Mapping
/// The enum maps to the following XML schema definition:
/// ```xml
/// <simpleType name="ST_SortBy">
///     <restriction base="xsd:string">
///         <enumeration value="value"/>
///         <enumeration value="cellColor"/>
///         <enumeration value="fontColor"/>
///         <enumeration value="icon"/>
///     </restriction>
/// </simpleType>
/// ```
///
/// # Variants
/// - `Value`: Represents sorting by value.
/// - `CellColor`: Represents sorting by cell color.
/// - `FontColor`: Represents sorting by font color.
/// - `Icon`: Represents sorting by icon.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum STSortBy {
    /// Represents sorting by value.
    Value,
    /// Represents sorting by cell color.
    CellColor,
    /// Represents sorting by font color.
    FontColor,
    /// Represents sorting by icon.
    Icon,
}
impl TryFrom<Vec<u8>> for STSortBy {
    type Error = XlsxError;
    fn try_from(value: Vec<u8>) -> Result<Self, Self::Error> {
        match value.as_slice() {
            b"value" => Ok(STSortBy::Value),
            b"cellColor" => Ok(STSortBy::CellColor),
            b"fontColor" => Ok(STSortBy::FontColor),
            b"icon" => Ok(STSortBy::Icon),
            v => {
                let value = String::from_utf8_lossy(v);
                Err(XlsxError::MissingVariant("STSortBy".into(), value.into()))
            }
        }
    }
}
/// Represents the type of icon set used for conditional formatting in a document.
///
/// This enum corresponds to the `ST_IconSetType` simple type in the XML schema.
/// It defines the available icon set types that can be used in a document.
///
/// # XML Schema Mapping
/// The enum maps to the following XML schema definition:
/// ```xml
/// <simpleType name="ST_IconSetType">
///     <restriction base="xsd:string">
///         <enumeration value="3Arrows"/>
///         <enumeration value="3Flags"/>
///         <enumeration value="3TrafficLights1"/>
///         <enumeration value="3TrafficLights2"/>
///         <enumeration value="3Signs"/>
///         <enumeration value="3Symbols"/>
///         <enumeration value="3Symbols2"/>
///         <enumeration value="4Arrows"/>
///         <enumeration value="4TrafficLights"/>
///         <enumeration value="5Arrows"/>
///         <enumeration value="5TrafficLights"/>
///         <enumeration value="5Quarters"/>
///     </restriction>
/// </simpleType>
/// ```
///
/// # Variants
/// - `ThreeArrows`: Represents a set of 3 arrows used for conditional formatting.
/// - `ThreeFlags`: Represents a set of 3 flags for conditional formatting.
/// - `ThreeTrafficLights1`: Represents a set of 3 traffic lights (set 1) for conditional formatting.
/// - `ThreeTrafficLights2`: Represents a set of 3 traffic lights (set 2) for conditional formatting.
/// - `ThreeSigns`: Represents a set of 3 signs for conditional formatting.
/// - `ThreeSymbols`: Represents a set of 3 symbols for conditional formatting.
/// - `ThreeSymbols2`: Represents a different set of 3 symbols for conditional formatting.
/// - `FourArrows`: Represents a set of 4 arrows for conditional formatting.
/// - `FourTrafficLights`: Represents a set of 4 traffic lights for conditional formatting.
/// - `FiveArrows`: Represents a set of 5 arrows for conditional formatting.
/// - `FiveTrafficLights`: Represents a set of 5 traffic lights for conditional formatting.
/// - `FiveQuarters`: Represents a set of 5 quarters for conditional formatting.
#[derive(Debug, Clone, PartialEq, Eq, EnumToBytes)]
pub enum STIconSetType {
    /// Represents a set of 3 arrows used for conditional formatting.
    #[name = "3Arrows"]
    ThreeArrows,
    /// Represents a set of 3 flags for conditional formatting.
    #[name = "3Flags"]
    ThreeFlags,
    /// Represents a set of 3 traffic lights (set 1) for conditional formatting.
    #[name = "3TrafficLights1"]
    ThreeTrafficLights1,
    /// Represents a set of 3 traffic lights (set 2) for conditional formatting.
    #[name = "3TrafficLights2"]
    ThreeTrafficLights2,
    /// Represents a set of 3 signs for conditional formatting.
    #[name = "3Signs"]
    ThreeSigns,
    /// Represents a set of 3 symbols for conditional formatting.
    #[name = "3Symbols"]
    ThreeSymbols,
    /// Represents a different set of 3 symbols for conditional formatting.
    #[name = "3Symbols2"]
    ThreeSymbols2,
    /// Represents a set of 4 arrows for conditional formatting.
    #[name = "4Arrows"]
    FourArrows,
    /// Represents a set of 4 traffic lights for conditional formatting.
    #[name = "4TrafficLights"]
    FourTrafficLights,
    /// Represents a set of 5 arrows for conditional formatting.
    #[name = "5Arrows"]
    FiveArrows,
    /// Represents a set of 5 traffic lights for conditional formatting.
    #[name = "5TrafficLights"]
    FiveTrafficLights,
    /// Represents a set of 5 quarters for conditional formatting.
    #[name = "5Quarters"]
    FiveQuarters,
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
/// - `descending`: Indicates whether the sort is in descending order.
/// - `sort_by`: Specifies how to sort the data.
/// - `reference`: The reference for the range to be sorted.
/// - `custom_list`: Specifies a custom list for sorting.
/// - `dxf_id`: Applies the style for sorting.
/// - `icon_set`: Specifies an icon set for sorting.
/// - `icon_id`: The ID for the icon to be applied.
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
struct CTSortCondition {
    descending: bool,
    reference: Vec<u8>,
    sort_by: Vec<u8>,
    custom_list: Vec<u8>,
    dxf_id: Vec<u8>,
    icon_set: Vec<u8>,
    icon_id: Vec<u8>,
}
impl CTSortCondition {
    /// Creates a new `CT_SortCondition` with xml schema default values.
    fn new() -> Self {
        Self {
            sort_by: b"value".into(),
            icon_set: b"3Arrows".into(),
            ..Default::default()
        }
    }
}
/// Represents the sorting method used in a document.
///
/// This enum corresponds to the `ST_SortMethod` simple type in the XML schema.
/// It defines the methods available for sorting, including options for sorting by stroke order,
/// Pinyin, or no sorting.
///
/// # XML Schema Mapping
/// The enum maps to the following XML schema definition:
/// ```xml
/// <simpleType name="ST_SortMethod">
///     <restriction base="xsd:string">
///         <enumeration value="stroke"/>
///         <enumeration value="pinYin"/>
///         <enumeration value="none"/>
///     </restriction>
/// </simpleType>
/// ```
///
/// # Variants
/// - `Stroke`: Represents sorting based on stroke order, typically used for Chinese characters.
/// - `PinYin`: Represents sorting based on the Pinyin romanization system, also used for Chinese characters.
/// - `None`: Represents no sorting method, used as a default.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub enum STSortMethod {
    /// Sorting based on stroke order.
    Stroke,
    /// Sorting based on Pinyin.
    PinYin,
    /// Default value, representing no sorting method.
    #[default]
    None,
}
impl TryFrom<Vec<u8>> for STSortMethod {
    type Error = XlsxError;
    fn try_from(value: Vec<u8>) -> Result<Self, Self::Error> {
        match value.as_slice() {
            b"pinYin" => Ok(STSortMethod::PinYin),
            b"none" => Ok(STSortMethod::Stroke),
            b"stroke" => Ok(STSortMethod::None),
            v => {
                let value = String::from_utf8_lossy(v);
                Err(XlsxError::MissingVariant(
                    "STSortMethod".into(),
                    value.into(),
                ))
            }
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
/// - `column_sort`: Indicates whether column sorting is enabled.
/// - `case_sensitive`: Indicates whether sorting is case-sensitive.
/// - `sort_method`: Specifies the sorting method used.
/// - `reference`: The reference for the range to be sorted.
/// - `sort_conditions`: A list of sorting conditions.
#[derive(Debug, Default, Clone, PartialEq, XmlRead, XmlWrite)]
struct CTSortState {
    column_sort: bool,
    case_sensitive: bool,
    sort_method: Vec<u8>,
    reference: Vec<u8>,
    #[xml(element)]
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
/// - `reference`: The reference for the range of the filter.
/// - `filter_column`: A list of filter columns.
/// - `sort_state`: The sorting state for the filter.
#[derive(Debug, Default, Clone, PartialEq, XmlRead)]
struct CTAutoFilter {
    reference: Vec<u8>,
    #[xml(following_elements)]
    filter_column: Vec<CTFilterColumn>,
    sort_state: Option<CTSortState>,
}
impl CTAutoFilter {
    /// Creates a new `CT_AutoFilter` with xml schema default values.
    fn new() -> Self {
        Self {
            ..Default::default()
        }
    }
}
