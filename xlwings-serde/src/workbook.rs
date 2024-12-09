use crate::{PreprocessNamespace, MAIN_NAMESPACE};
use serde::Deserialize;
use yaserde::{ser::to_string, YaDeserialize, YaSerialize};

/// Deserialize the .xlsx file `xl/workbook.xml`
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(rename = "workbook", namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mv" = "urn:schemas-microsoft-com:mac:vml",
    "x14" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x15" = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xm" = "http://schemas.microsoft.com/office/excel/2006/main",
    "mx" = "http://schemas.microsoft.com/office/mac/excel/2008/main"
})]
pub struct Book {
    #[yaserde(rename = "workbookPr")]
    book_properties: BookProperty,
    #[yaserde(rename = "sheets")]
    sheet: Sheet,
    #[yaserde(rename = "definedNames")]
    defined_names: DefinedName,
    #[yaserde(rename = "calcPr")]
    calculation_properties: CalculationProperty,
}
impl ToString for Book {
    fn to_string(&self) -> String {
        let original_namespaces = format!(
            r#"<workbook {MAIN_NAMESPACE} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"#
        );

        let cleared_namespaces = r#"<workbook xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"#;
        to_string(self).unwrap()
        .replace(&cleared_namespaces, &original_namespaces)
    }
}
impl PreprocessNamespace for Book {}

#[derive(YaSerialize, YaDeserialize, Deserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mv" = "urn:schemas-microsoft-com:mac:vml",
    "x14" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x15" = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xm" = "http://schemas.microsoft.com/office/excel/2006/main",
    "mx" = "http://schemas.microsoft.com/office/mac/excel/2008/main"
})]
struct BookProperty {}
#[derive(YaSerialize, YaDeserialize, Deserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mv" = "urn:schemas-microsoft-com:mac:vml",
    "x14" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x15" = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xm" = "http://schemas.microsoft.com/office/excel/2006/main",
    "mx" = "http://schemas.microsoft.com/office/mac/excel/2008/main"
})]
struct DefinedName {}
#[derive(YaSerialize, YaDeserialize, Deserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mv" = "urn:schemas-microsoft-com:mac:vml",
    "x14" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x15" = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xm" = "http://schemas.microsoft.com/office/excel/2006/main",
    "mx" = "http://schemas.microsoft.com/office/mac/excel/2008/main"
})]
struct CalculationProperty {}
#[derive(YaSerialize, YaDeserialize, Deserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mv" = "urn:schemas-microsoft-com:mac:vml",
    "x14" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x15" = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xm" = "http://schemas.microsoft.com/office/excel/2006/main",
    "mx" = "http://schemas.microsoft.com/office/mac/excel/2008/main"
})]
struct Sheet {
    #[yaserde(rename = "sheet")]
    sheets: Vec<SheetChildren>,
}
#[derive(YaSerialize, YaDeserialize, Deserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mv" = "urn:schemas-microsoft-com:mac:vml",
    "x14" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x15" = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xm" = "http://schemas.microsoft.com/office/excel/2006/main",
    "mx" = "http://schemas.microsoft.com/office/mac/excel/2008/main"
})]
struct SheetChildren {
    #[yaserde(rename = "state", attribute = true)]
    state: String,
    #[yaserde(rename = "name", attribute = true)]
    name: String,
    #[yaserde(rename = "sheetId", attribute = true)]
    sheet_id: String,
    #[yaserde(rename = "id", attribute = true, prefix = "r")]
    r_id: String,
}
