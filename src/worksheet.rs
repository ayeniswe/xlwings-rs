use std::{
    borrow::Cow,
    sync::{Arc, Mutex},
};

use crate::{errors::ExcelError, helper::SharedStringTable};
use rust_xlsxwriter::XlsxError;
use xlwings_serde::{
    sheet::{Cell, Row},
    Relationship, SharedString, Sheet,
};

const COL_MAX: u16 = 16_384;
const ROW_MAX: u32 = 1_048_576;
const MAX_STRING_LEN: usize = 32_767;
const DEFAULT_STYLE_IDX: &str = "1";

pub struct Worksheet {
    shared_string_table: Arc<Mutex<SharedStringTable>>,
    sheet: Sheet,
    relationship: Relationship,
}
impl Worksheet {
    /// 0-indexed
    pub fn write(
        &mut self,
        row: u32,
        col: u16,
        data: impl IntoExcelData,
    ) -> Result<&mut Self, ExcelError> {
        data.write(self, row, col)
    }
    /// Convert integer to respective Excel column letter
    fn number_to_letter(&self, col: u16) -> String {
        let mut letter = String::new();
        let mut number = col + 1;
        while number > 0 {
            // Adjust for 0-based indexing
            number -= 1;
            let remainder = (number % 26) as u8;
            letter.insert(0, (b'A' + remainder) as char);
            number /= 26;
        }
        letter
    }
    fn store_string(
        &mut self,
        row: u32,
        col: u16,
        string: String,
    ) -> Result<&mut Self, ExcelError> {
        if col >= COL_MAX {
            return Err(ExcelError::ColumnLimitError);
        }
        if row >= ROW_MAX {
            return Err(ExcelError::RowLimitError);
        }
        if string.chars().count() >= MAX_STRING_LEN {
            return Err(ExcelError::MaxStringLengthExceeded);
        }

        // handle 0 based index
        let row = row + 1;
        let col = col + 1;

        ///// impl shared string table
        let row = Row {
            index: row.to_string(),
            cells: vec![Cell {
                column: self.number_to_letter(col) + &row.to_string(),
                style_index: DEFAULT_STYLE_IDX.to_string(),
                r#type: None,
                value: Some(string),
            }],
        };
        self.sheet.data.rows.push(row);
        Ok(self)
    }
}

#[test]
fn t() {
    
}

pub trait IntoExcelData {
    /// Trait method to handle writing an unformatted type to Excel.
    ///
    /// # Errors
    ///
    /// - [`ExcelError::RowLimitError`] - Row exceeds Excel's
    /// - [`ExcelError::ColumnLimitError`] - Column exceeds Excel's
    ///   worksheet limits.
    /// - [`ExcelError::MaxStringLengthExceeded`] - String exceeds Excel's limit
    ///   of 32,767 characters.
    ///
    fn write(
        self,
        worksheet: &mut Worksheet,
        row: u32,
        col: u16,
    ) -> Result<&mut Worksheet, ExcelError>;
}
macro_rules! write_string_trait_impl {
    ($($t:ty)*) => ($(
        impl IntoExcelData for $t {
            fn write(
                self,
                worksheet: &mut Worksheet,
                row: u32,
                col: u16,
            ) -> Result<&mut Worksheet, ExcelError> {
                worksheet.store_string(row, col, self.into())
            }
        }
    )*)
}
write_string_trait_impl!(&str &String String Cow<'_, str>);
