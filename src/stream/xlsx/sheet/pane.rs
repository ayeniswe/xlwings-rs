use crate::stream::{
    utils::{XmlReader, XmlWriter},
    xlsx::XlsxError,
};
use derive::{XmlRead, XmlWrite};
use quick_xml::{events::Event, Reader, Writer};
use std::io::{BufRead, Write};

/// Represents the position of a pane in a spreadsheet.
///
/// This enum corresponds to the `ST_Pane` simple type in the XML schema, which defines
/// the possible positions for panes within a worksheet.
///
/// # XML Schema Mapping
/// ```xml
/// <simpleType name="ST_Pane">
///     <restriction base="xsd:string">
///         <enumeration value="bottomLeft"/>
///         <enumeration value="bottomRight"/>
///         <enumeration value="topLeft"/>
///         <enumeration value="topRight"/>
///     </restriction>
/// </simpleType>
/// ```
///
/// # Variants
/// - `BottomLeft` – Bottom left pane, used when both vertical and horizontal splits are applied.
/// - `BottomRight` – Bottom right pane, used when both vertical and horizontal splits are applied.
/// - `TopLeft` – Top left pane, used when both vertical and horizontal splits are applied.
/// - `TopRight` – Top right pane, used when both vertical and horizontal splits are applied.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
enum STPane {
    BottomRight,
    TopRight,
    BottomLeft,
    #[default]
    TopLeft,
}
impl TryFrom<Vec<u8>> for STPane {
    type Error = XlsxError;
    fn try_from(value: Vec<u8>) -> Result<Self, Self::Error> {
        match value.as_slice() {
            b"bottomLeft" => Ok(STPane::BottomLeft),
            b"bottomRight" => Ok(STPane::BottomRight),
            b"topLeft" => Ok(STPane::TopLeft),
            b"topRight" => Ok(STPane::TopRight),
            v => {
                let value = String::from_utf8_lossy(v);
                Err(XlsxError::MissingVariant("STPane".into(), value.into()))
            }
        }
    }
}
/// Defines the state of a pane in a spreadsheet.
///
/// This enum corresponds to the `ST_PaneState` simple type in the XML schema, which
/// specifies the state of the sheet's pane.
///
/// # XML Schema Mapping
/// ```xml
/// <simpleType name="ST_PaneState">
///     <restriction base="xsd:string">
///         <enumeration value="split"/>
///         <enumeration value="frozen"/>
///         <enumeration value="frozenSplit"/>
///     </restriction>
/// </simpleType>
/// ```
///
/// # Variants
/// - `Frozen` – Panes are frozen without prior splitting; unfreezing results in a single pane.
/// - `Split` – Panes are split but not frozen; split bars are adjustable by the user.
/// - `FrozenSplit` – Panes are frozen after being split; unfreezing retains the split, which becomes adjustable.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub enum STPaneState {
    /// Panes are frozen without prior splitting; unfreezing results in a single pane.
    Frozen,
    /// Panes are split but not frozen; split bars are adjustable by the user.
    #[default]
    Split,
    /// Panes are frozen after being split; unfreezing retains the split, which becomes adjustable.
    FrozenSplit,
}
impl TryFrom<Vec<u8>> for STPaneState {
    type Error = XlsxError;
    fn try_from(value: Vec<u8>) -> Result<Self, Self::Error> {
        match value.as_slice() {
            b"frozen" => Ok(STPaneState::Frozen),
            b"split" => Ok(STPaneState::Split),
            b"frozenSplit" => Ok(STPaneState::FrozenSplit),
            v => {
                let value = String::from_utf8_lossy(v);
                Err(XlsxError::MissingVariant("STPaneState".into(), value.into()))
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
    fn new() -> Self {
        Self {
            x_split: b"0".into(),
            y_split: b"0".into(),
            active_pane: b"topLeft".into(),
            state: b"split".into(),
            ..Default::default()
        }
    }
}
