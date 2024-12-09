use yaserde::{ser::to_string, YaDeserialize, YaSerialize};

use crate::PreprocessNamespace;

/// Deserialize the .xlsx file(s) `xl/theme/theme1.xml`
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(prefix = "a", rename = "theme", namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
pub struct Theme {
    #[yaserde(rename = "name", attribute = true)]
    name: String,
    #[yaserde(rename = "themeElements", prefix = "a")]
    element: Element,
}
impl ToString for Theme {
    fn to_string(&self) -> String {
        to_string(self).unwrap()
    }
}
impl PreprocessNamespace for Theme {}

#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct Element {
    #[yaserde(rename = "clrScheme", prefix = "a")]
    color_scheme: ColorScheme,
    #[yaserde(rename = "fontScheme", prefix = "a")]
    font_scheme: FontScheme,
    #[yaserde(rename = "fmtScheme", prefix = "a")]
    format_scheme: FormatScheme,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct ColorScheme {
    #[yaserde(rename = "name", attribute = true)]
    name: String,
    #[yaserde(rename = "dk1", prefix = "a")]
    dark1: Scheme,
    #[yaserde(rename = "lt1", prefix = "a")]
    light1: Scheme,
    #[yaserde(rename = "dk2", prefix = "a")]
    dark2: Scheme,
    #[yaserde(rename = "lt2", prefix = "a")]
    light2: Scheme,
    #[yaserde(rename = "accent1", prefix = "a")]
    accent1: Scheme,
    #[yaserde(rename = "accent2", prefix = "a")]
    accent2: Scheme,
    #[yaserde(rename = "accent3", prefix = "a")]
    accent3: Scheme,
    #[yaserde(rename = "accent4", prefix = "a")]
    accent4: Scheme,
    #[yaserde(rename = "accent5", prefix = "a")]
    accent5: Scheme,
    #[yaserde(rename = "accent6", prefix = "a")]
    accent6: Scheme,
    #[yaserde(rename = "hlink", prefix = "a")]
    hyperlink: Scheme,
    #[yaserde(rename = "folHlink", prefix = "a")]
    followed_hyperlink: Scheme,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct FontScheme {
    #[yaserde(rename = "name", attribute = true)]
    name: String,
    #[yaserde(rename = "majorFont", prefix = "a")]
    major_font: Font,
    #[yaserde(rename = "minorFont", prefix = "a")]
    minor_font: Font,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct FormatScheme {
    #[yaserde(rename = "name", attribute = true)]
    name: String,
    #[yaserde(rename = "fillStyleLst", prefix = "a")]
    fill_style: FillStyle,
    #[yaserde(rename = "lnStyleLst", prefix = "a")]
    line_styles: LineStyle,
    #[yaserde(rename = "effectStyleLst", prefix = "a")]
    effect_style: EffectStyle,
    #[yaserde(rename = "bgFillStyleLst", prefix = "a")]
    bg_fill_style: FillStyle,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct EffectStyle {
    #[yaserde(rename = "effectStyle", prefix = "a")]
    effects: Vec<Effect>,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct Effect {
    #[yaserde(rename = "effectLst", prefix = "a")]
    r#type: EffectType,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct EffectType {
    #[yaserde(rename = "outerShdw", prefix = "a")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    outer_shadow: Option<Shadow>,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct Shadow {
    #[yaserde(rename = "blurRad", attribute = true)]
    blur_radius: String,
    #[yaserde(rename = "dist", attribute = true)]
    distance: String,
    #[yaserde(rename = "dir", attribute = true)]
    direction: String,
    #[yaserde(rename = "algn", attribute = true)]
    alignment: String,
    #[yaserde(rename = "rotWithShape", attribute = true)]
    rotate_with_shape: String,
    #[yaserde(rename = "srgbClr", prefix = "a")]
    color: ShadowColor,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct ShadowColor {
    #[yaserde(rename = "val", attribute = true)]
    value: String,
    #[yaserde(rename = "alpha", prefix = "a")]
    alpha: PropertyValue,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct FillStyle {
    #[yaserde(rename = "solidFill", prefix = "a")]
    solid_fill: Vec<SolidFill>,
    #[yaserde(rename = "gradFill", prefix = "a")]
    gradient_fill: Vec<GradientFill>,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct LineStyle {
    #[yaserde(rename = "ln", prefix = "a")]
    solid_fill: Vec<Line>,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct Line {
    #[yaserde(rename = "w", attribute = true)]
    width: String,
    #[yaserde(rename = "cap", attribute = true)]
    cap_style: String,
    #[yaserde(rename = "cmpd", attribute = true)]
    compound_style: String,
    #[yaserde(rename = "algn", attribute = true)]
    alignment: String,
    #[yaserde(rename = "solidFill", prefix = "a")]
    solid_fill: SolidFill,
    #[yaserde(rename = "miter", prefix = "a")]
    miter: Miter,
    #[yaserde(rename = "prstDash", prefix = "a")]
    dash_style: PropertyValue,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct SolidFill {
    #[yaserde(rename = "schemeClr", prefix = "a")]
    solid_fill: Color,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct GradientFill {
    #[yaserde(rename = "rotWithShape", attribute = true)]
    rotate_with_shape: String,
    #[yaserde(rename = "gsLst", prefix = "a")]
    gradient: GradientList,
    #[yaserde(rename = "lin", prefix = "a")]
    linear: Linear,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct GradientList {
    #[yaserde(rename = "gs", prefix = "a")]
    gradients: Vec<Gradient>,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct Linear {
    #[yaserde(rename = "ang", attribute = true)]
    angle: String,
    #[yaserde(rename = "scaled", attribute = true)]
    scaled: String,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct Gradient {
    #[yaserde(rename = "pos", attribute = true)]
    pos: String,
    #[yaserde(rename = "schemeClr", prefix = "a")]
    scheme: Color,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct Font {
    #[yaserde(rename = "latin", prefix = "a")]
    latin: FontInfo,
    #[yaserde(rename = "ea", prefix = "a")]
    ea: FontInfo,
    #[yaserde(rename = "cs", prefix = "a")]
    cs: FontInfo,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct Scheme {
    #[yaserde(rename = "srgbClr", prefix = "a")]
    srgb: Color,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct FontInfo {
    #[yaserde(rename = "typeface", attribute = true)]
    typeface: String,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct Color {
    #[yaserde(rename = "val", attribute = true)]
    value: String,
    #[yaserde(rename = "lumMod", prefix = "a")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    lum: Option<PropertyValue>,
    #[yaserde(rename = "tint", prefix = "a")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    tint: Option<PropertyValue>,
    #[yaserde(rename = "satMod", prefix = "a")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    saturation: Option<PropertyValue>,
    #[yaserde(rename = "shade", prefix = "a")]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    shade: Option<PropertyValue>,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct PropertyValue {
    #[yaserde(rename = "val", attribute = true)]
    value: String,
}
#[derive(YaSerialize, YaDeserialize, Debug)]
#[yaserde(namespaces = {
    "r" = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a" = "http://schemas.openxmlformats.org/drawingml/2006/main"
})]
struct Miter {
    #[yaserde(rename = "lim", attribute = true)]
    limit: String,
}
