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
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlRead, XmlWrite)]
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

    #[xml(element, name = "x")]
    selected_items: Vec<CTIndex>,
}
impl CTPivotAreaReference {
    /// Creates a new `CT_PivotAreaReference` instance with xml schema default values.
    pub(crate) fn new() -> Self {
        Self {
            selected: true,
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
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite, XmlRead)]
pub(crate) struct CTPivotAreaReferences {
    #[xml(name = "count")]
    count: Vec<u8>,
    #[xml(element, name = "references")]
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
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlWrite, XmlRead)]
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

    #[xml(element, name ="references")]
    reference_collection: CTPivotAreaReferences,
}
impl CTPivotArea {
    /// Creates a new `CT_PivotArea` instance with xml schema default values.
    pub(crate) fn new() -> Self {
        Self {
            outline: true,
            use_data_only: true,
            pivot_type: b"normal".into(),
            ..Default::default()
        }
    }
}