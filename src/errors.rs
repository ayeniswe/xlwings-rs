use thiserror::Error;

/// Hiearchy of the entire crate's error types
#[derive(Error, Debug)]
pub enum XcelmateError {
    #[error("excel columns can not exceed 16,384")]
    ExcelMaxColumnExceeded,
    #[error("excel rows can not exceed 1,048,576")]
    ExcelMaxRowExceeded,
    #[error("can not parse excel dimension: {0}")]
    ExcelDimensionParseError(String),
    #[error("styles.xml is missing")]
    StylesMissing,
    /// A not specified enum variant. This can happen if we do not get
    /// a full coverage of the possible variants that can exist when parsing strings
    #[error("({0}) missing variant for: {1}")]
    MissingVariant(String, String),
    /// Stream reading has reached the end so more than likely enclosed tags are incorrect or missing
    #[error("malformed stream for tag: {0}")]
    XmlEof(String),

    /// The `std::io` error wrapper
    #[error(transparent)]
    StdErr(#[from] std::io::Error),
    /// The `std::char` int error wrapper
    #[error(transparent)]
    ParseChar(#[from] std::char::ParseCharError),
    /// The `std::num` int error wrapper
    #[error(transparent)]
    ParseInt(#[from] std::num::ParseIntError),
    /// The `quick_xml` crate error wrapper
    #[error(transparent)]
    Xml(#[from] quick_xml::Error),
    /// The `quick_xml::events::attributes` crate error wrapper
    #[error(transparent)]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),
    /// The `zip` crate error wrapper
    #[error(transparent)]
    Zip(#[from] zip::result::ZipError),
}

impl From<String> for XcelmateError {
    fn from(v: String) -> Self {
        Self::XmlEof(v)
    }
}
