use crate::stream::{
    utils::{XmlReader, XmlWriter},
    xlsx::XlsxError,
};
use derive::{XmlRead, XmlWrite};
use quick_xml::{events::Event, Reader, Writer};
use std::io::{BufRead, Write};

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
#[derive(Debug, Default, PartialEq, Clone, Eq, XmlRead, XmlWrite)]
pub(crate) struct CTSelection {
    #[xml(name = "pane")]
    pane: Vec<u8>,
    #[xml(name = "activeCell")]
    cell: Vec<u8>,
    #[xml(name = "activeCellId", default_bytes = b"0")]
    cell_id: Vec<u8>,
    #[xml(name = "sqref", default_bytes = b"A1")]
    sqref: Vec<u8>,
}
impl CTSelection {
    /// Creates a new `CT_Selection` instance with xml schema default values.
    fn new() -> Self {
        Self {
            sqref: b"A1".into(),
            cell_id: b"0".into(),
            ..Default::default()
        }
    }
}