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
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlRead, XmlWrite)]
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