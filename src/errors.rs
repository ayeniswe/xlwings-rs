use thiserror::Error;

/// Hiearchy of the entire crate's error types
#[derive(Error, Debug)]
pub enum XcelmateError {
    /// The .xlsx error wrapper
    #[error(transparent)]
    Xcelmate(#[from] crate::stream::xlsx::errors::XlsxError),
}
