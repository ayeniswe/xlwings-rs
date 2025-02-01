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
/// ```
#[derive(Debug, XmlRead, XmlWrite, Default, Clone, PartialEq, Eq)]
pub(crate) struct CTPane {
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
impl CTPane {
    /// Creates a new `CT_Pane` instance with xml schema default values.
    pub fn new() -> Self {
        Self {
            x_split: b"0".into(),
            y_split: b"0".into(),
            active_pane: b"topLeft".into(),
            state: b"split".into(),
            ..Default::default()
        }
    }
}