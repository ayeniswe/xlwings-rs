use crate::{PreprocessNamespace, MAIN_NAMESPACE};
use yaserde::{ser::to_string, YaDeserialize, YaSerialize};

/// Deserialize the .xlsx file(s) `xl/worksheets/sheet1.xml`
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(rename = "worksheet", namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "mv" = "urn:schemas-microsoft-com:mac:vml",
    "x14" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x15" = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "xm" = "http://schemas.microsoft.com/office/excel/2006/main",
    "mx" = "http://schemas.microsoft.com/office/mac/excel/2008/main"
})]
pub struct Sheet {
    #[yaserde(rename = "sheetPr")]
    property: Property,
    #[yaserde(rename = "sheetViews")]
    views: SheetView,
    #[yaserde(rename = "sheetFormatPr")]
    format_property: FormatProperty,
    #[yaserde(rename = "sheetData")]
    pub data: Data,
    #[yaserde(rename = "drawing")]
    drawing: Drawing,
}
impl ToString for Sheet {
    fn to_string(&self) -> String {
        let original_namespaces = format!(
            r#"<worksheet {MAIN_NAMESPACE} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"#
        );

        let cleared_namespaces = r#"<worksheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:mx="http://schemas.microsoft.com/office/mac/excel/2008/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"#;
        to_string(self)
            .unwrap()
            .replace(&cleared_namespaces, &original_namespaces)
    }
}
impl PreprocessNamespace for Sheet {}

#[derive(YaDeserialize, YaSerialize, Debug)]
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
struct Drawing {
    #[yaserde(rename = "id", attribute = true, prefix = "r")]
    id: String,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
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
pub struct Data {
    #[yaserde(rename = "row")]
    pub rows: Vec<Row>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
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
pub struct Row {
    #[yaserde(rename = "r", attribute = true)]
    pub index: String,
    #[yaserde(rename = "c")]
    pub cells: Vec<Cell>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
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
pub struct Cell {
    #[yaserde(rename = "r", attribute = true)]
    pub column: String,
    #[yaserde(rename = "s", attribute = true)]
    pub style_index: String,
    #[yaserde(rename = "t", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    pub r#type: Option<String>,
    #[yaserde(rename = "v")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
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
struct FormatProperty {
    #[yaserde(rename = "customHeight", attribute = true)]
    height: String,
    #[yaserde(rename = "defaultColWidth", attribute = true)]
    column_width: String,
    #[yaserde(rename = "defaultRowHeight", attribute = true)]
    row_height: String,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
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
struct SheetView {
    #[yaserde(rename = "sheetView")]
    view: SheetViewChildren,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
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
struct SheetViewChildren {
    #[yaserde(rename = "workbookViewId", attribute = true)]
    id: String,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
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
struct Property {
    #[yaserde(rename = "outlinePr")]
    outline: OutlineProperty,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
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
struct OutlineProperty {
    #[yaserde(rename = "summaryBelow", attribute = true)]
    summary_below: String,
    #[yaserde(rename = "summaryRight", attribute = true)]
    summary_right: String,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
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
struct SharedStringItem {
    #[yaserde(rename = "t")]
    text: Text,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
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
struct Text {
    #[yaserde(text = true)]
    value: String,
}
