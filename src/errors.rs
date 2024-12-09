use thiserror::Error;

#[derive(Error, Debug)]
pub enum ExcelError {
    #[error("Row exceeds Excel's allowed limits 1,048,576.")]
    RowLimitError,
    #[error("Column exceeds Excel's allowed limits 16,384.")]
    ColumnLimitError,
    #[error("String exceeds Excel's limit of 32,767 characters.")]
    MaxStringLengthExceeded,
}
