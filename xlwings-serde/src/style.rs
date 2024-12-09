use yaserde::{ser::to_string, YaDeserialize, YaSerialize};

use crate::{PreprocessNamespace, MAIN_NAMESPACE};

/// Deserialize the .xlsx file `xl/styles.xml`
///
/// Responsible for the styling of the entire workbook
/// such as cells fg and bg coloring, bordering, and fonts
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(rename= "styleSheet", namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
pub struct Style {
    #[yaserde(rename = "fonts")]
    fonts: Vec<Font>,
    #[yaserde(rename = "fills")]
    fills: Vec<Fill>,
    #[yaserde(rename = "borders")]
    borders: Vec<Border>,
    #[yaserde(rename = "cellStyleXfs")]
    cellstyles_xfs: ExtendedFormat,
    #[yaserde(rename = "cellXfs")]
    cell_xfs: ExtendedFormat,
    #[yaserde(rename = "cellStyles")]
    cellstyles: Vec<CellStyle>,
    #[yaserde(rename = "dxfs")]
    dxfs: DExtendedFormat,
}
impl ToString for Style {
    fn to_string(&self) -> String {
        let original_namespaces = format!(
            r#"<styleSheet {MAIN_NAMESPACE} xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">"#
        );
        let cleared_namespaces = r#"<styleSheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">"#;
        to_string(self)
            .unwrap()
            .replace(&cleared_namespaces, &original_namespaces)
    }
}
impl PreprocessNamespace for Style {}

#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct DExtendedFormat {
    #[yaserde(rename = "count", attribute = true)]
    count: String,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct ExtendedFormat {
    #[yaserde(rename = "count", attribute = true)]
    count: String,
    #[yaserde(rename = "xf")]
    xf: Vec<ExtendedFormatChildren>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct ExtendedFormatChildren {
    #[yaserde(rename = "borderId", attribute = true)]
    border_id: String,
    #[yaserde(rename = "fillId", attribute = true)]
    fill_id: String,
    #[yaserde(rename = "fontId", attribute = true)]
    font_id: String,
    #[yaserde(rename = "numFmtId", attribute = true)]
    numfmt_id: String,
    #[yaserde(rename = "xfId", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    xf_id: Option<String>,
    #[yaserde(rename = "applyBorder", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    apply_border: Option<String>,
    #[yaserde(rename = "applyAlignment", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    apply_alignment: Option<String>,
    #[yaserde(rename = "applyFill", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    apply_fill: Option<String>,
    #[yaserde(rename = "applyFont", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    apply_font: Option<String>,
    #[yaserde(rename = "applyNumberFormat", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    apply_numfmt: Option<String>,
    #[yaserde(rename = "alignment")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    alignment: Option<Alignment>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct Alignment {
    #[yaserde(rename = "readingOrder", attribute = true)]
    reading_order: String,
    #[yaserde(rename = "shrinkToFit", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    shrink_to_fit: Option<String>,
    #[yaserde(rename = "vertical", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    vertical: Option<String>,
    #[yaserde(rename = "wrapText", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    wrap_text: Option<String>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct CellStyle {
    #[yaserde(rename = "count", attribute = true)]
    count: String,
    #[yaserde(rename = "cellStyle")]
    cellstyle: CellStyleChildren,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct CellStyleChildren {
    #[yaserde(rename = "xfId", attribute = true)]
    xf_id: String,
    #[yaserde(rename = "name", attribute = true)]
    name: String,
    #[yaserde(rename = "builtinId", attribute = true)]
    builtin_id: String,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct Border {
    #[yaserde(rename = "count", attribute = true)]
    count: String,
    #[yaserde(rename = "border")]
    border: Vec<BorderChildren>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct Font {
    #[yaserde(rename = "count", attribute = true)]
    count: String,
    #[yaserde(rename = "font")]
    font: Vec<FontChildren>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct Fill {
    #[yaserde(rename = "count", attribute = true)]
    count: String,
    #[yaserde(rename = "fill")]
    fill: Vec<FillChildren>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct FillChildren {
    #[yaserde(rename = "patternFill")]
    pattern_fill: PatternFill,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct BorderChildren {
    #[yaserde(rename = "left")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    left: Option<BorderDirection>,
    #[yaserde(rename = "right")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    right: Option<BorderDirection>,
    #[yaserde(rename = "top")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    top: Option<BorderDirection>,
    #[yaserde(rename = "bottom")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    bottom: Option<BorderDirection>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct BorderDirection {
    #[yaserde(rename = "style", attribute = true)]
    style: String,
    #[yaserde(rename = "color")]
    color: Color,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct PatternFill {
    #[yaserde(rename = "patternType", attribute = true)]
    r#type: String,
    #[yaserde(rename = "fgColor")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    foreground_color: Option<ForegroundColor>,
    #[yaserde(rename = "bgColor")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    background_color: Option<BackgroundColor>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct BackgroundColor {
    #[yaserde(rename = "rgb", attribute = true)]
    rgb: String,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct ForegroundColor {
    #[yaserde(rename = "rgb", attribute = true)]
    rgb: String,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct FontChildren {
    #[yaserde(rename = "sz")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    size: Option<Size>,
    #[yaserde(rename = "color")]
    color: Color,
    #[yaserde(rename = "name")]
    name: Name,
    #[yaserde(rename = "scheme")]
    scheme: Scheme,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct Size {
    #[yaserde(rename = "val", attribute = true)]
    value: String,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct Color {
    #[yaserde(rename = "rgb", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    rgb: Option<String>,
    #[yaserde(rename = "theme", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    theme: Option<String>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct Name {
    #[yaserde(rename = "val", attribute = true)]
    value: String,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(namespaces = {
    "mc" = "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14ac" = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
})]
struct Scheme {
    #[yaserde(rename = "val", attribute = true)]
    value: String,
}
