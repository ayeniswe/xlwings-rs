use crate::{
    errors::XcelmateError,
    stream::utils::{xml_reader, Key, Save, XmlWriter},
};
use bimap::{BiBTreeMap, BiHashMap, BiMap};
use quick_xml::{
    events::{attributes::Attributes, BytesDecl, BytesEnd, BytesStart, Event},
    name::QName,
    Reader, Writer,
};
use std::{
    collections::HashMap,
    io::{BufRead, Read, Seek, Write},
    ops::RangeInclusive,
    sync::Arc,
};
use zip::{
    write::{FileOptionExtension, FileOptions},
    ZipArchive,
};

/// The `Rgb` promotes better api usage with hexadecimal coloring
#[derive(Debug, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) enum Rgb {
    Custom(u8, u8, u8),
}
impl ToString for Rgb {
    fn to_string(&self) -> String {
        match self {
            Rgb::Custom(r, g, b) => format!(
                "FF{}{}{}",
                format!("{:02X}", r),
                &format!("{:02X}", g),
                &format!("{:02X}", b)
            ),
        }
    }
}
/// The `Color` denotes the type of coloring system to
/// use since excel has builtin coloring to choose that will map to `theme` but
/// for custom specfic coloring `rgb` is used
///
/// Default is equivalent to `black`
#[derive(Debug, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) enum Color {
    /// Builtin theme from excel color palette selector which includes theme id and tint value
    Theme {
        id: u32,
        tint: Option<String>,
    },
    /// RGB color model
    Rgb(Rgb),
    Index(u32),
    Auto(u32),
}
impl Default for Color {
    fn default() -> Self {
        Color::Theme { id: 1, tint: None }
    }
}
impl<W: Write> XmlWriter<W> for Color {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &'a str,
    ) -> Result<&'a mut Writer<W>, XcelmateError> {
        let writer = writer.create_element(tag_name);
        match self {
            Color::Theme { id, tint } => {
                let writer = writer.with_attribute(("theme", id.to_string().as_str()));
                if let Some(tint) = tint {
                    let writer = writer.with_attribute(("tint", tint.as_str()));
                    Ok(writer.write_empty()?)
                } else {
                    Ok(writer.write_empty()?)
                }
            }
            Color::Rgb(rgb) => {
                let writer = writer.with_attribute(("rgb", rgb.to_string().as_str()));
                Ok(writer.write_empty()?)
            }
            Color::Index(idx) => {
                let writer = writer.with_attribute(("indexed", idx.to_string().as_str()));
                Ok(writer.write_empty()?)
            }
            Color::Auto(val) => {
                let writer = writer.with_attribute(("auto", val.to_string().as_str()));
                Ok(writer.write_empty()?)
            }
        }
    }
}

/// Some `FontProperty` values can be used in conditional scenarios so being able to override base styles
/// requires a tri value
#[derive(Debug, Default, Clone, PartialEq, Eq, PartialOrd, Hash, Ord)]
pub(crate) enum FormatState {
    /// Sets attribute val="1"
    Enabled,
    /// Sets attribute val="0"
    Disabled,
    #[default]
    /// Value will not show
    None,
}

/// The `FontProperty` denotes all styling options
/// that can be added to text
#[derive(Debug, Default, Clone, PartialEq, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct FontProperty {
    pub(crate) strikethrough: FormatState,
    pub(crate) outline: FormatState,
    pub(crate) shadow: FormatState,
    pub(crate) subscript: FormatState,
    pub(crate) baseline: FormatState,
    pub(crate) superscript: FormatState,
    pub(crate) bold: FormatState,
    pub(crate) underline: FormatState,
    /// Double underline
    pub(crate) double: FormatState,
    pub(crate) italic: FormatState,
    pub(crate) size: String,
    pub(crate) color: Color,
    /// Font type
    pub(crate) font: String,
    /// Font family
    pub(crate) family: u32,
    /// Font scheme
    pub(crate) scheme: String,
    /// Allow duplicate with counter since it will always hash different
    pub(crate) dup_cnt: usize,
}

impl<W: Write> XmlWriter<W> for FontProperty {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &str,
    ) -> Result<&'a mut Writer<W>, XcelmateError> {
        writer
            .create_element(tag_name)
            .write_inner_content::<_, XcelmateError>(|writer| {
                match self.strikethrough {
                    FormatState::Enabled => writer.create_element("strike").write_empty()?,
                    FormatState::Disabled => writer
                        .create_element("strike")
                        .with_attribute(("val", "0"))
                        .write_empty()?,
                    FormatState::None => writer,
                };
                match self.outline {
                    FormatState::Enabled => writer.create_element("outline").write_empty()?,
                    FormatState::Disabled => writer
                        .create_element("outline")
                        .with_attribute(("val", "0"))
                        .write_empty()?,
                    FormatState::None => writer,
                };
                match self.shadow {
                    FormatState::Enabled => writer.create_element("shadow").write_empty()?,
                    FormatState::Disabled => writer
                        .create_element("shadow")
                        .with_attribute(("val", "0"))
                        .write_empty()?,
                    FormatState::None => writer,
                };
                match (&self.superscript, &self.subscript, &self.baseline) {
                    (FormatState::Enabled, _, _) => writer
                        .create_element("vertAlign")
                        .with_attribute(("val", "superscript"))
                        .write_empty()?,
                    (_, FormatState::Enabled, _) => writer
                        .create_element("vertAlign")
                        .with_attribute(("val", "subscript"))
                        .write_empty()?,
                    (_, _, FormatState::Enabled) => writer
                        .create_element("vertAlign")
                        .with_attribute(("val", "baseline"))
                        .write_empty()?,
                    _ => writer,
                };
                match self.bold {
                    FormatState::Enabled => writer.create_element("b").write_empty()?,
                    FormatState::Disabled => writer
                        .create_element("b")
                        .with_attribute(("val", "0"))
                        .write_empty()?,
                    FormatState::None => writer,
                };
                match self.italic {
                    FormatState::Enabled => writer.create_element("i").write_empty()?,
                    FormatState::Disabled => writer
                        .create_element("i")
                        .with_attribute(("val", "0"))
                        .write_empty()?,
                    FormatState::None => writer,
                };
                match (&self.underline, &self.double) {
                    (FormatState::Enabled, _) => writer.create_element("u").write_empty()?,
                    (FormatState::Disabled, _) => writer
                        .create_element("u")
                        .with_attribute(("val", "none"))
                        .write_empty()?,
                    (_, FormatState::Enabled) => writer
                        .create_element("u")
                        .with_attribute(("val", "double"))
                        .write_empty()?,
                    _ => writer,
                };
                if !self.size.is_empty() {
                    writer
                        .create_element("sz")
                        .with_attribute(("val", self.size.as_str()))
                        .write_empty()?;
                }
                self.color.write_xml(writer, "color")?;
                if !self.font.is_empty() {
                    writer
                        .create_element(if tag_name == "font" { "name" } else { "rFont" }) //the similarity of rich text and font tags are identical except for this
                        .with_attribute(("val", self.font.as_str()))
                        .write_empty()?;
                }
                if self.family != u32::default() {
                    writer
                        .create_element("family")
                        .with_attribute(("val", self.family.to_string().as_str()))
                        .write_empty()?;
                }
                if !self.scheme.is_empty() {
                    writer
                        .create_element("scheme")
                        .with_attribute(("val", self.scheme.as_str()))
                        .write_empty()?;
                }
                Ok(())
            })?;

        Ok(writer)
    }
}

/// The range for number formats that are based on local currency
const LOCALIZED_RANGE_NUMBER_FORMAT: RangeInclusive<usize> = 41..=44;
/// The highest reserved id for number formats before custom number formats are detected
const MAX_RESERVED_NUMBER_FORMAT: usize = 163;
/// The formatting style to use on numbers
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct NumberFormat {
    id: u32,
    format_code: String,
}
impl<W: Write> XmlWriter<W> for NumberFormat {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &str,
    ) -> Result<&'a mut Writer<W>, XcelmateError> {
        writer
            .create_element(tag_name)
            .with_attributes(vec![
                ("numFmtId", self.id.to_string().as_str()),
                ("formatCode", self.format_code.as_str()),
            ])
            .write_empty()?;
        Ok(writer)
    }
}

/// The pattern fill styling to apply to a cell
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
enum PatternFill {
    #[default]
    None,
    Solid,
    Gray,
}
impl<W: Write> XmlWriter<W> for PatternFill {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &'a str,
    ) -> Result<&'a mut Writer<W>, XcelmateError> {
        match self {
            PatternFill::None => Ok(writer
                .create_element(tag_name)
                .with_attribute(("patternType", "none"))
                .write_empty()?),
            PatternFill::Gray => Ok(writer
                .create_element(tag_name)
                .with_attribute(("patternType", "gray125"))
                .write_empty()?),
            _ => Ok(writer),
        }
    }
}

/// The background/foreground fill of a cell. Also can include gradients
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct Fill {
    r#type: PatternFill,
    foreground: Option<Color>,
    background: Option<Color>,
}
impl<W: Write> XmlWriter<W> for Fill {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &'a str,
    ) -> Result<&'a mut Writer<W>, XcelmateError> {
        let writer = writer
            .create_element(tag_name)
            .write_inner_content::<_, XcelmateError>(|writer| {
                let writer_fill = writer.create_element("patternFill");
                match (&self.r#type, &self.background, &self.foreground) {
                    (PatternFill::None, Some(bg), Some(fg)) => writer_fill
                        .write_inner_content::<_, XcelmateError>(|writer| {
                            fg.write_xml(writer, "fgColor")?;
                            bg.write_xml(writer, "bgColor")?;
                            Ok(())
                        })?,
                    (PatternFill::Solid, Some(bg), Some(fg)) => writer_fill
                        .with_attribute(("patternType", "solid"))
                        .write_inner_content::<_, XcelmateError>(|writer| {
                            fg.write_xml(writer, "fgColor")?;
                            bg.write_xml(writer, "bgColor")?;
                            Ok(())
                        })?,
                    _ => self.r#type.write_xml(writer, tag_name)?,
                };
                Ok(())
            });
        Ok(writer?)
    }
}

/// The type of line styling for a border
#[derive(Debug, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
enum BorderStyle {
    /// Thin border
    Thin,
    /// Medium border
    Medium,
    /// Thick border
    Thick,
    /// Double line border
    Double,
    /// Dashed border
    Dashed,
    /// Dotted border
    Dotted,
    /// Dash-dot border
    DashDot,
    /// Dash-dot-dot border
    DashDotDot,
    /// Slant dash-dot border
    SlantDashDot,
    /// Hairline border
    Hair,
    /// Medium dashed border
    MediumDashed,
    /// Medium dash-dot border
    MediumDashDot,
    /// Medium dash-dot-dot border
    MediumDashDotDot,
}
impl ToString for BorderStyle {
    fn to_string(&self) -> String {
        match self {
            BorderStyle::Thin => "thin".into(),
            BorderStyle::Medium => "medium".into(),
            BorderStyle::Thick => "thick".into(),
            BorderStyle::Double => "double".into(),
            BorderStyle::Dashed => "dashed".into(),
            BorderStyle::Dotted => "dotted".into(),
            BorderStyle::DashDot => "dashDot".into(),
            BorderStyle::DashDotDot => "dashDotDot".into(),
            BorderStyle::SlantDashDot => "slantDashDot".into(),
            BorderStyle::Hair => "hair".into(),
            BorderStyle::MediumDashed => "mediumDashed".into(),
            BorderStyle::MediumDashDot => "mediumDashDot".into(),
            BorderStyle::MediumDashDotDot => "mediumDashDotDot".into(),
        }
    }
}
/// The border region to apply styling to
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
struct BorderRegion {
    style: Option<BorderStyle>,
    color: Option<Color>,
}
impl<W: Write> XmlWriter<W> for BorderRegion {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &'a str,
    ) -> Result<&'a mut Writer<W>, XcelmateError> {
        if let (Some(style), Some(color)) = (&self.style, &self.color) {
            let writer = writer
                .create_element(tag_name)
                .with_attribute(("style", style.to_string().as_str()))
                .write_inner_content::<_, XcelmateError>(|writer| {
                    color.write_xml(writer, "color")?;
                    Ok(())
                });
            return Ok(writer?);
        }
        Ok(writer)
    }
}
/// The styling for all border regions of a cell
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct Border {
    left: BorderRegion,
    right: BorderRegion,
    top: BorderRegion,
    bottom: BorderRegion,
    diagonal: BorderRegion,
    vertical: BorderRegion,
    horizontal: BorderRegion,
}
impl<W: Write> XmlWriter<W> for Border {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &'a str,
    ) -> Result<&'a mut Writer<W>, XcelmateError> {
        let writer = writer
            .create_element(tag_name)
            .write_inner_content::<_, XcelmateError>(|writer| {
                self.left.write_xml(writer, "left")?;
                self.right.write_xml(writer, "right")?;
                self.top.write_xml(writer, "top")?;
                self.bottom.write_xml(writer, "bottom")?;
                self.vertical.write_xml(writer, "vertical")?;
                self.horizontal.write_xml(writer, "horizontal")?;
                self.diagonal.write_xml(writer, "diagonal")?;
                Ok(())
            });
        Ok(writer?)
    }
}
/// The horizontal alignment of a cell
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) enum HorizontalAlignment {
    #[default]
    Left,
    Center,
    Right,
}
impl ToString for HorizontalAlignment {
    fn to_string(&self) -> String {
        match self {
            HorizontalAlignment::Left => "left".into(),
            HorizontalAlignment::Center => "center".into(),
            HorizontalAlignment::Right => "right".into(),
        }
    }
}

/// The vertical alignment of a cell
#[derive(Debug, Default, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) enum VerticalAlignment {
    Top,
    Center,
    #[default]
    Bottom,
}
impl ToString for VerticalAlignment {
    fn to_string(&self) -> String {
        match self {
            VerticalAlignment::Top => "top".into(),
            VerticalAlignment::Center => "center".into(),
            VerticalAlignment::Bottom => "bottom".into(),
        }
    }
}

/// The alignment attributes of a cell
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct Alignment {
    wrap: bool,
    valign: VerticalAlignment,
    indent: bool,
    halign: HorizontalAlignment,
}

/// The styling traits of a cell
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct CellXf {
    number_format: Option<Arc<NumberFormat>>,
    font: Arc<FontProperty>,
    fill: Arc<Fill>,
    border: Arc<Border>,
    quote_prefix: bool,
    align: Option<Alignment>,
}

/// The styling groups for differential conditional formatting
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct DiffXf {
    font: Option<FontProperty>,
    fill: Option<Fill>,
    border: Option<Border>,
    dup_cnt: usize,
}
impl<W: Write> XmlWriter<W> for DiffXf {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &'a str,
    ) -> Result<&'a mut Writer<W>, XcelmateError> {
        let writer = writer
            .create_element(tag_name)
            .write_inner_content::<_, XcelmateError>(|writer| {
                if let Some(font) = &self.font {
                    font.write_xml(writer, "font")?;
                }
                if let Some(fill) = &self.fill {
                    fill.write_xml(writer, "fill")?;
                }
                if let Some(border) = &self.border {
                    border.write_xml(writer, "border")?;
                }
                Ok(())
            });
        Ok(writer?)
    }
}

/// The grouping of custom table styles
#[derive(Debug, PartialEq, Default, Clone, Eq)]
pub(crate) struct TableStyle {
    default_style: String,
    default_pivot_style: String,
    styles: HashMap<String, Arc<TableCustomStyle>>,
}

/// Table design pieces
#[derive(Debug, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
enum TableStyleElement {
    Table(Arc<DiffXf>),
    Header(Arc<DiffXf>),
    FirstRow(Arc<DiffXf>),
    SecondRow(Arc<DiffXf>),
}
impl ToString for TableStyleElement {
    fn to_string(&self) -> String {
        match self {
            TableStyleElement::Table(_) => "wholeTable".into(),
            TableStyleElement::Header(_) => "headerRow".into(),
            TableStyleElement::FirstRow(_) => "firstRowStripe".into(),
            TableStyleElement::SecondRow(_) => "secondRowStripe".into(),
        }
    }
}

/// A custom table style
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct TableCustomStyle {
    name: String,
    uid: String,
    pivot: usize,
    elements: Vec<TableStyleElement>,
}

/// The `Stylesheet` provides a mapping of styles properties such as fonts, colors, themes, etc
#[derive(Default)]
pub(crate) struct Stylesheet {
    number_formats_builtin: Option<BiMap<Arc<NumberFormat>, Key>>,
    number_formats: Option<BiMap<Arc<NumberFormat>, Key>>,
    fonts: BiBTreeMap<Arc<FontProperty>, Key>,
    fills: BiBTreeMap<Arc<Fill>, Key>,
    borders: BiBTreeMap<Arc<Border>, Key>,
    cell_xf: BiBTreeMap<Arc<CellXf>, Key>,
    diff_xf: BiBTreeMap<Arc<DiffXf>, Key>,
    table_style: Option<TableStyle>,
}
impl<W: Write> XmlWriter<W> for Stylesheet {
    fn write_xml<'a>(
        &self,
        writer: &'a mut Writer<W>,
        tag_name: &str,
    ) -> Result<&'a mut Writer<W>, XcelmateError> {
        writer.write_event(Event::Decl(BytesDecl::new(
            "1.0",
            Some("UTF-8"),
            Some("yes"),
        )))?;
        writer
            .create_element(tag_name)
            .with_attributes(vec![
                (
                    "xmlns",
                    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                ),
                (
                    "xmlns:mc",
                    "http://schemas.openxmlformats.org/markup-compatibility/2006",
                ),
                ("mc:Ignorable", "x14ac x16r2 xr xr9"),
                (
                    "xmlns:x14ac",
                    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
                ),
                (
                    "xmlns:x16r2",
                    "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",
                ),
                (
                    "xmlns:xr",
                    "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
                ),
                (
                    "xmlns:xr9",
                    "http://schemas.microsoft.com/office/spreadsheetml/2016/revision9",
                ),
            ])
            .write_inner_content::<_, XcelmateError>(|writer| {
                // <numFmts>
                if let Some(numfmt) = &self.number_formats {
                    // Includes builtin in total count (should remove this)
                    let _ = writer
                        .create_element("numFmts")
                        .with_attribute(("count", numfmt.len().to_string().as_str()))
                        .write_inner_content::<_, XcelmateError>(|writer| {
                            for n in numfmt.left_values() {
                                n.write_xml(writer, "numFmt")?;
                            }
                            Ok(())
                        });
                }
                // <fonts>
                let _ = writer
                    .create_element("fonts")
                    .with_attribute(("count", self.fonts.len().to_string().as_str()))
                    .write_inner_content::<_, XcelmateError>(|writer| {
                        for (font, _) in self.fonts.right_range(0..self.fonts.len()) {
                            font.write_xml(writer, "font")?;
                        }
                        Ok(())
                    });
                // <fills>
                let _ = writer
                    .create_element("fills")
                    .with_attribute(("count", self.fills.len().to_string().as_str()))
                    .write_inner_content::<_, XcelmateError>(|writer| {
                        for (fill, _) in self.fills.right_range(0..self.fonts.len()) {
                            fill.write_xml(writer, "fill")?;
                        }
                        Ok(())
                    });
                // <borders>
                let _ = writer
                    .create_element("borders")
                    .with_attribute(("count", self.borders.len().to_string().as_str()))
                    .write_inner_content::<_, XcelmateError>(|writer| {
                        for (border, _) in self.borders.right_range(0..self.borders.len()) {
                            border.write_xml(writer, "border")?;
                        }
                        Ok(())
                    });
                // <cellStyleXfs>
                let _ = writer
                    .create_element("cellStyleXfs")
                    .with_attribute(("count", "1"))
                    .write_inner_content::<_, XcelmateError>(|writer| {
                        writer
                            .create_element("xf")
                            .with_attributes(vec![
                                ("numFmtId", "0"),
                                ("fontId", "0"),
                                ("fillId", "0"),
                                ("borderId", "0"),
                            ])
                            .write_empty()?;
                        Ok(())
                    });
                // <cellXfs>
                let _ = writer
                    .create_element("cellXfs")
                    .with_attribute(("count", self.cell_xf.len().to_string().as_str()))
                    .write_inner_content::<_, XcelmateError>(|writer| {
                        for (xf, _) in self.cell_xf.right_range(0..self.cell_xf.len()) {
                            let writer = writer.create_element("xf");

                            let numfmt_id = if let Some(numfmt) = &xf.number_format {
                                self.get_key_from_number_format_ref(numfmt.clone()).unwrap()
                            } else {
                                0
                            };
                            let writer = writer.with_attributes(vec![
                                ("numFmtId", numfmt_id.to_string().as_str()),
                                (
                                    "fontId",
                                    self.get_key_from_font_ref(xf.font.clone())
                                        .unwrap()
                                        .to_string()
                                        .as_str(),
                                ),
                                (
                                    "fillId",
                                    self.get_key_from_fill_ref(xf.fill.clone())
                                        .unwrap()
                                        .to_string()
                                        .as_str(),
                                ),
                                (
                                    "borderId",
                                    self.get_key_from_border_ref(xf.border.clone())
                                        .unwrap()
                                        .to_string()
                                        .as_str(),
                                ),
                            ]);
                            let writer = if xf.quote_prefix {
                                writer.with_attribute(("quotePrefix", "1"))
                            } else {
                                writer
                            };

                            if let Some(align) = &xf.align {
                                writer.write_inner_content::<_, XcelmateError>(|writer| {
                                    let mut attrs = vec![];
                                    if align.wrap {
                                        attrs.push(("wrapText", "1"))
                                    }
                                    if align.indent {
                                        attrs.push(("indent", "1"))
                                    }
                                    match align.valign {
                                        VerticalAlignment::Top => attrs.push(("vertical", "top")),
                                        VerticalAlignment::Center => {
                                            attrs.push(("vertical", "center"))
                                        }
                                        VerticalAlignment::Bottom => (),
                                    }
                                    match align.halign {
                                        HorizontalAlignment::Left => (),
                                        HorizontalAlignment::Center => {
                                            attrs.push(("horizontal", "center"))
                                        }
                                        HorizontalAlignment::Right => {
                                            attrs.push(("horizontal", "right"))
                                        }
                                    }
                                    writer
                                        .create_element("alignment")
                                        .with_attributes(attrs)
                                        .write_empty()?;
                                    Ok(())
                                })?;
                            } else {
                                writer.write_empty()?;
                            };
                        }
                        Ok(())
                    });
                // <cellStyles>
                let _ = writer
                    .create_element("cellStyles")
                    .with_attribute(("count", "1"))
                    .write_inner_content::<_, XcelmateError>(|writer| {
                        writer
                            .create_element("cellStyle")
                            .with_attributes(vec![
                                ("name", "Normal"),
                                ("xfId", "0"),
                                ("builtinId", "0"),
                            ])
                            .write_empty()?;
                        Ok(())
                    });
                // <dxfs>
                let _ = writer
                    .create_element("dxfs")
                    .with_attribute(("count", self.diff_xf.len().to_string().as_str()))
                    .write_inner_content::<_, XcelmateError>(|writer| {
                        for (diff_xf, _) in self.diff_xf.right_range(0..self.diff_xf.len()) {
                            let _ = writer
                                .create_element("dxf")
                                .write_inner_content::<_, XcelmateError>(|writer| {
                                    if let Some(font) = &diff_xf.font {
                                        font.write_xml(writer, "font")?;
                                    }
                                    if let Some(fill) = &diff_xf.fill {
                                        fill.write_xml(writer, "fill")?;
                                    }
                                    if let Some(border) = &diff_xf.border {
                                        border.write_xml(writer, "border")?;
                                    }
                                    Ok(())
                                });
                        }
                        Ok(())
                    });
                // <tableStyles>
                if let Some(table_style) = &self.table_style {
                    let table_style_writer =
                        writer.create_element("tableStyles").with_attributes(vec![
                            ("count", table_style.styles.len().to_string().as_str()),
                            (
                                "defaultPivotStyle",
                                table_style.default_pivot_style.as_str(),
                            ),
                            ("defaultTableStyle", table_style.default_style.as_str()),
                        ]);
                    if !table_style.styles.is_empty() {
                        // <tableStyle>
                        let _ =
                            table_style_writer.write_inner_content::<_, XcelmateError>(|writer| {
                                for (_, style) in &table_style.styles {
                                    let _ = writer
                                        .create_element("tableStyle")
                                        .with_attributes(vec![
                                            ("pivot", style.pivot.to_string().as_str()),
                                            ("count", style.elements.len().to_string().as_str()),
                                            ("xr9:uid", style.uid.as_str()),
                                            ("name", style.name.as_str()),
                                        ])
                                        // <tableStyleElement>
                                        .write_inner_content::<_, XcelmateError>(|writer| {
                                            for ele in &style.elements {
                                                let dxf = match ele {
                                                    TableStyleElement::Table(dxf)
                                                    | TableStyleElement::Header(dxf)
                                                    | TableStyleElement::FirstRow(dxf)
                                                    | TableStyleElement::SecondRow(dxf) => dxf,
                                                };
                                                let dxf_id = self
                                                    .get_key_from_differential_ref(dxf.clone())
                                                    .unwrap()
                                                    .to_string();
                                                writer
                                                    .create_element("tableStyleElement")
                                                    .with_attributes(vec![
                                                        ("type", ele.to_string().as_str()),
                                                        ("dxfId", dxf_id.as_str()),
                                                    ])
                                                    .write_empty()?;
                                            }
                                            Ok(())
                                        });
                                }
                                Ok(())
                            });
                    } else {
                        table_style_writer.write_empty()?;
                    }
                }
                Ok(())
            })?;
        Ok(writer)
    }
}
impl<W: Write + Seek, EX: FileOptionExtension> Save<W, EX> for Stylesheet {
    fn save(
        &mut self,
        writer: &mut zip::ZipWriter<W>,
        options: FileOptions<EX>,
    ) -> Result<(), XcelmateError> {
        writer.start_file("xl/styles.xml", options)?;
        self.write_xml(&mut Writer::new(writer), "styleSheet")?;
        Ok(())
    }
}
impl Stylesheet {
    pub(crate) fn read_stylesheet<'a, RS: Read + Seek>(
        &mut self,
        zip: &'a mut ZipArchive<RS>,
    ) -> Result<(), XcelmateError> {
        let mut xml = match xml_reader(zip, "xl/styles.xml") {
            None => return Err(XcelmateError::StylesMissing),
            Some(x) => x?,
        };
        let mut buf = Vec::with_capacity(1024);
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"numFmts" => {
                    self.number_formats = Some(BiMap::new());
                }
                ////////////////////
                // NUMBER FORMATS
                /////////////
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"numFmt" => {
                    ////////////////////
                    // NUMBER FORMATS Attrs
                    /////////////
                    let mut numfmt = NumberFormat::default();
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"numFmtId") => {
                                    numfmt.id = a.unescape_value()?.parse::<u32>()?
                                }
                                QName(b"formatCode") => {
                                    numfmt.format_code = a.unescape_value()?.to_string()
                                }
                                _ => (),
                            }
                        }
                    }
                    self.add_number_format_ref_to_table(Arc::new(numfmt));
                }
                ////////////////////
                // FONT
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                    // Allow duplicates by increment dup count so all duplicate reflect their respective duplicate count
                    let mut font = Stylesheet::read_font(&mut xml, e.name())?;
                    if let Some(id) = self.get_key_from_font_ref(font.clone().into()) {
                        let dup_cnt = self.get_font_ref_from_key(id).unwrap().dup_cnt + 1;
                        font.dup_cnt = dup_cnt;
                        let _ = self.add_font_ref_to_table(font.into());
                    } else {
                        let _ = self.add_font_ref_to_table(font.into());
                    }
                }
                ////////////////////
                // FILL
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                    let fill = Stylesheet::read_fill(&mut xml, e.name())?;
                    self.add_fill_ref_to_table(fill.into());
                }
                ////////////////////
                // BORDER
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                    self.add_border_ref_to_table(
                        Stylesheet::read_border(&mut xml, e.name())?.into(),
                    );
                }
                ////////////////////
                // CELL REFERENCES
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cellXfs" => {
                    let mut cell_xf_buf = Vec::with_capacity(1024);
                    loop {
                        cell_xf_buf.clear();
                        let mut cell_xf = CellXf::default();
                        let event = xml.read_event_into(&mut cell_xf_buf);
                        match event {
                            ////////////////////
                            // CELL REFERENCES nth-1
                            /////////////
                            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                                if e.local_name().as_ref() == b"xf" =>
                            {
                                for attr in e.attributes() {
                                    if let Ok(a) = attr {
                                        match a.key {
                                            QName(b"numFmtId") => {
                                                let key = a.unescape_value()?.parse::<usize>()?;
                                                cell_xf.number_format =
                                                    self.get_number_format_ref_from_key(key);
                                            }
                                            QName(b"fontId") => {
                                                let key = a.unescape_value()?.parse::<usize>()?;
                                                cell_xf.font = self.get_font_ref_from_key(key).expect("all font styles should have been captured previously");
                                            }
                                            QName(b"fillId") => {
                                                let key = a.unescape_value()?.parse::<usize>()?;
                                                cell_xf.fill = self.get_fill_ref_from_key(key).expect("all fill styles should have been captured previously");
                                            }
                                            QName(b"borderId") => {
                                                let key = a.unescape_value()?.parse::<usize>()?;
                                                cell_xf.border = self.get_border_ref_from_key(key).expect("all border styles should have been captured previously");
                                            }
                                            QName(b"quotePrefix") => {
                                                let val = a.unescape_value()?.parse::<usize>()?;
                                                if val == 1 {
                                                    cell_xf.quote_prefix = true;
                                                }
                                            }
                                            _ => (),
                                        }
                                    }
                                }
                                ////////////////////
                                // CELL REFERENCES nth-2
                                /////////////
                                if let Ok(Event::Start(_)) = event {
                                    let mut val_buf = Vec::with_capacity(1024);
                                    loop {
                                        val_buf.clear();
                                        let event = xml.read_event_into(&mut val_buf);
                                        match event {
                                            Ok(Event::Empty(ref e))
                                                if e.local_name().as_ref() == b"alignment" =>
                                            {
                                                let mut align = Alignment::default();
                                                for attr in e.attributes() {
                                                    if let Ok(a) = attr {
                                                        match a.key {
                                                            QName(b"vertical") => {
                                                                let val =
                                                                    a.unescape_value()?.to_string();
                                                                match val.as_str() {
                                                                    "center" => align.valign =
                                                                        VerticalAlignment::Center,
                                                                    "top" => {
                                                                        align.valign =
                                                                            VerticalAlignment::Top
                                                                    }
                                                                    _ => (),
                                                                };
                                                            }
                                                            QName(b"wrapText") => {
                                                                let val = a
                                                                    .unescape_value()?
                                                                    .parse::<usize>()?;
                                                                if val == 1 {
                                                                    align.wrap = true;
                                                                }
                                                            }
                                                            QName(b"horizontal") => {
                                                                let val =
                                                                    a.unescape_value()?.to_string();
                                                                match val.as_str() {
                                                                    "center" => align.halign =
                                                                        HorizontalAlignment::Center,
                                                                    "right" => align.halign =
                                                                        HorizontalAlignment::Right,
                                                                    _ => (),
                                                                };
                                                            }
                                                            QName(b"indent") => {
                                                                let val = a
                                                                    .unescape_value()?
                                                                    .parse::<usize>()?;
                                                                if val == 1 {
                                                                    align.indent = true;
                                                                }
                                                            }
                                                            _ => (),
                                                        }
                                                    }
                                                }
                                                cell_xf.align = Some(align);
                                            }
                                            Ok(Event::End(ref e))
                                                if e.local_name().as_ref() == b"xf" =>
                                            {
                                                break
                                            }
                                            Ok(Event::Eof) => {
                                                return Err(XcelmateError::XmlEof(
                                                    "alignment".into(),
                                                ))
                                            }
                                            Err(e) => return Err(XcelmateError::Xml(e)),
                                            _ => (),
                                        }
                                    }
                                }
                                self.add_cell_ref_to_table(Arc::new(cell_xf));
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cellXfs" => break,
                            Ok(Event::Eof) => return Err(XcelmateError::XmlEof("cellXfs".into())),
                            Err(e) => return Err(XcelmateError::Xml(e)),
                            _ => (),
                        }
                    }
                }
                ////////////////////
                // DIFFERENTIAL REFERENCE
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dxf" => {
                    let mut dx_buf = Vec::with_capacity(1024);
                    let mut diff_xf = DiffXf::default();
                    loop {
                        dx_buf.clear();
                        match xml.read_event_into(&mut dx_buf) {
                            ////////////////////
                            // DIFFERENTIAL REFERENCE nth-1
                            /////////////
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                                diff_xf.font = Some(Stylesheet::read_font(&mut xml, e.name())?);
                            }
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                                diff_xf.border = Some(Stylesheet::read_border(&mut xml, e.name())?);
                            }
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                                diff_xf.fill = Some(Stylesheet::read_fill(&mut xml, e.name())?);
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dxf" => break,
                            Ok(Event::Eof) => return Err(XcelmateError::XmlEof("dxf".into())),
                            Err(e) => return Err(XcelmateError::Xml(e)),
                            _ => (),
                        }
                    }

                    // Allow duplicates by increment dup count so all duplicate reflect their respective duplicate count
                    if let Some(id) = self.get_key_from_differential_ref(diff_xf.clone().into()) {
                        let dup_cnt = self.get_differential_ref_from_key(id).unwrap().dup_cnt + 1;
                        diff_xf.dup_cnt = dup_cnt;
                        let _ = self.add_differential_ref_to_table(diff_xf.into());
                    } else {
                        let _ = self.add_differential_ref_to_table(diff_xf.into());
                    }
                }
                ////////////////////
                // TABLE STYLE
                /////////////
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"tableStyles" => {
                    let mut table_style = TableStyle::default();
                    ////////////////////
                    // TABLE STYLE Attrs
                    /////////////
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"defaultTableStyle") => {
                                    table_style.default_style = a.unescape_value()?.to_string();
                                }
                                QName(b"defaultPivotStyle") => {
                                    table_style.default_pivot_style =
                                        a.unescape_value()?.to_string();
                                }
                                _ => (),
                            }
                        }
                    }
                    self.table_style = Some(table_style);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tableStyles" => {
                    let mut table_style = TableStyle::default();
                    ////////////////////
                    // TABLE STYLE Attrs
                    /////////////
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"defaultTableStyle") => {
                                    table_style.default_style = a.unescape_value()?.to_string();
                                }
                                QName(b"defaultPivotStyle") => {
                                    table_style.default_pivot_style =
                                        a.unescape_value()?.to_string();
                                }
                                _ => (),
                            }
                        }
                    }

                    let mut table_style_buf = Vec::with_capacity(1024);
                    let mut custom_style = TableCustomStyle::default();
                    loop {
                        table_style_buf.clear();
                        match xml.read_event_into(&mut table_style_buf) {
                            ////////////////////
                            // TABLE STYLE nth-1
                            /////////////
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tableStyle" => {
                                for attr in e.attributes() {
                                    if let Ok(a) = attr {
                                        match a.key {
                                            QName(b"name") => {
                                                custom_style.name = a.unescape_value()?.to_string();
                                            }
                                            QName(b"pivot") => {
                                                custom_style.pivot =
                                                    a.unescape_value()?.parse::<usize>()?;
                                            }
                                            QName(b"xr9:uid") => {
                                                custom_style.uid = a.unescape_value()?.to_string();
                                            }
                                            _ => (),
                                        }
                                    }
                                }
                            }
                            Ok(Event::Empty(ref e))
                                if e.local_name().as_ref() == b"tableStyleElement" =>
                            {
                                let mut key = 0;
                                let mut r#type = String::new();
                                for attr in e.attributes() {
                                    if let Ok(a) = attr {
                                        match a.key {
                                            QName(b"type") => {
                                                r#type = a.unescape_value()?.to_string();
                                            }
                                            QName(b"dxfId") => {
                                                key = a.unescape_value()?.parse::<usize>()?;
                                            }
                                            _ => (),
                                        }
                                    }
                                }
                                let diff = self.get_differential_ref_from_key(key).expect(
                                    "all differential should have been captured previously",
                                );
                                let ele = match r#type.as_str() {
                                    "wholeTable" => Ok(TableStyleElement::Table(diff)),
                                    "headerRow" => Ok(TableStyleElement::Header(diff)),
                                    "firstRowStripe" => Ok(TableStyleElement::FirstRow(diff)),
                                    "secondRowStripe" => Ok(TableStyleElement::SecondRow(diff)),
                                    v => Err(XcelmateError::MissingVariant(
                                        "TableStyleElement".into(),
                                        v.into(),
                                    )),
                                }?;
                                custom_style.elements.push(ele);
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"tableStyle" => {
                                table_style
                                    .styles
                                    .insert(custom_style.name.clone(), custom_style.into());
                                // Reset custom style buffer
                                custom_style = TableCustomStyle::default();
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"tableStyles" => {
                                break
                            }
                            Ok(Event::Eof) => {
                                return Err(XcelmateError::XmlEof("tableStyles".into()))
                            }
                            Err(e) => return Err(XcelmateError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.add_table_style(table_style.into());
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"styleSheet" => break,
                Ok(Event::Eof) => return Err(XcelmateError::XmlEof("styleSheet".into())),
                Err(e) => return Err(XcelmateError::Xml(e)),
                _ => (),
            }
        }
        Ok(())
    }

    pub(crate) fn get_custom_table_style(&self, name: &str) -> Option<Arc<TableCustomStyle>> {
        if let Some(t) = &self.table_style {
            t.styles.get(name).cloned()
        } else {
            None
        }
    }

    pub(crate) fn add_custom_table_style(
        &mut self,
        name: &str,
        style: Arc<TableCustomStyle>,
    ) -> Arc<TableCustomStyle> {
        if let Some(table) = &mut self.table_style {
            table.styles.insert(name.into(), style.clone());
        } else {
            self.table_style = Some(TableStyle {
                styles: HashMap::from_iter(vec![(name.into(), style.clone())]),
                ..Default::default()
            });
        }
        style
    }

    pub(crate) fn add_table_style(&mut self, table: TableStyle) {
        self.table_style = Some(table);
    }

    pub(crate) fn get_key_from_cell_ref(&self, key: Arc<CellXf>) -> Option<usize> {
        if let Some(i) = self.cell_xf.get_by_left(&key) {
            Some(*i)
        } else {
            None
        }
    }

    pub(crate) fn get_cell_ref_from_key(&self, key: Key) -> Option<Arc<CellXf>> {
        if let Some(i) = self.cell_xf.get_by_right(&key) {
            Some(i.clone())
        } else {
            None
        }
    }

    pub(crate) fn add_cell_ref_to_table(&mut self, item: Arc<CellXf>) -> Arc<CellXf> {
        self.cell_xf.insert(item.clone(), self.cell_xf.len());
        item
    }

    pub(crate) fn get_key_from_differential_ref(&self, key: Arc<DiffXf>) -> Option<usize> {
        if let Some(i) = self.diff_xf.get_by_left(&key) {
            Some(*i)
        } else {
            None
        }
    }

    pub(crate) fn get_differential_ref_from_key(&self, key: Key) -> Option<Arc<DiffXf>> {
        if let Some(i) = self.diff_xf.get_by_right(&key) {
            Some(i.clone())
        } else {
            None
        }
    }

    pub(crate) fn add_differential_ref_to_table(&mut self, item: Arc<DiffXf>) -> Arc<DiffXf> {
        self.diff_xf.insert(item.clone(), self.diff_xf.len());
        item
    }

    pub(crate) fn get_key_from_number_format_ref(&self, key: Arc<NumberFormat>) -> Option<usize> {
        if let Some(n) = &self.number_formats {
            if let Some(i) = n.get_by_left(&key) {
                Some(*i)
            } else {
                None
            }
        } else {
            None
        }
    }

    pub(crate) fn get_number_format_ref_from_key(&self, key: Key) -> Option<Arc<NumberFormat>> {
        if (LOCALIZED_RANGE_NUMBER_FORMAT).contains(&key) || key > MAX_RESERVED_NUMBER_FORMAT {
            if let Some(n) = &self.number_formats {
                if let Some(i) = n.get_by_right(&key) {
                    Some(i.clone())
                } else {
                    None
                }
            } else {
                None
            }
        } else {
            if let Some(n) = &self.number_formats_builtin {
                if let Some(i) = n.get_by_right(&key) {
                    Some(i.clone())
                } else {
                    None
                }
            } else {
                None
            }
        }
    }

    pub(crate) fn add_number_format_ref_to_table(
        &mut self,
        item: Arc<NumberFormat>,
    ) -> Arc<NumberFormat> {
        let key = item.id as usize;
        if (LOCALIZED_RANGE_NUMBER_FORMAT).contains(&key) || key >= MAX_RESERVED_NUMBER_FORMAT {
            if let Some(number_formats) = &mut self.number_formats {
                number_formats.insert(item.clone(), key);
            } else {
                self.number_formats = Some(BiHashMap::from_iter(vec![(item.clone(), key)]));
            }

            item
        } else {
            if let Some(number_formats) = &mut self.number_formats_builtin {
                number_formats.insert(item.clone(), key);
            } else {
                self.number_formats = Some(BiHashMap::from_iter(vec![(item.clone(), key)]));
            }

            item
        }
    }

    pub(crate) fn get_key_from_font_ref(&self, key: Arc<FontProperty>) -> Option<usize> {
        if let Some(i) = self.fonts.get_by_left(&key) {
            Some(*i)
        } else {
            None
        }
    }

    pub(crate) fn get_font_ref_from_key(&self, key: Key) -> Option<Arc<FontProperty>> {
        if let Some(i) = self.fonts.get_by_right(&key) {
            Some(i.clone())
        } else {
            None
        }
    }

    pub(crate) fn add_font_ref_to_table(&mut self, item: Arc<FontProperty>) -> Arc<FontProperty> {
        self.fonts.insert(item.clone(), self.fonts.len());
        item
    }

    pub(crate) fn get_key_from_fill_ref(&self, key: Arc<Fill>) -> Option<usize> {
        if let Some(i) = self.fills.get_by_left(&key) {
            Some(*i)
        } else {
            None
        }
    }

    pub(crate) fn get_fill_ref_from_key(&self, key: Key) -> Option<Arc<Fill>> {
        if let Some(i) = self.fills.get_by_right(&key) {
            Some(i.clone())
        } else {
            None
        }
    }

    pub(crate) fn add_fill_ref_to_table(&mut self, item: Arc<Fill>) -> Arc<Fill> {
        self.fills.insert(item.clone(), self.fills.len());
        item
    }

    pub(crate) fn get_key_from_border_ref(&self, key: Arc<Border>) -> Option<usize> {
        if let Some(i) = self.borders.get_by_left(&key) {
            Some(*i)
        } else {
            None
        }
    }

    pub(crate) fn get_border_ref_from_key(&self, key: Key) -> Option<Arc<Border>> {
        if let Some(i) = self.borders.get_by_right(&key) {
            Some(i.clone())
        } else {
            None
        }
    }

    pub(crate) fn add_border_ref_to_table(&mut self, item: Arc<Border>) -> Arc<Border> {
        self.borders.insert(item.clone(), self.borders.len());
        item
    }

    pub(crate) fn read_color(attributes: Attributes) -> Result<Color, XcelmateError>{
        ////////////////////
        // COLOR Attrs
        /////////////
        let mut color = Color::default();
        for attr in attributes {
            if let Ok(a) = attr {
                match a.key {
                    QName(b"rgb") => {
                        color =
                            Stylesheet::to_rgb(a.unescape_value()?.to_string())?;
                    }
                    QName(b"theme") => {
                        color = Color::Theme {
                            id: a.unescape_value()?.parse::<u32>()?,
                            tint: None,
                        };
                    }
                    QName(b"auto") => {
                        color =
                            Color::Auto(a.unescape_value()?.parse::<u32>()?);
                    }
                    QName(b"indexed") => {
                        color =
                            Color::Index(a.unescape_value()?.parse::<u32>()?);
                    }
                    QName(b"tint") => match color {
                        Color::Theme { id, .. } => {
                            color = Color::Theme {
                                id,
                                tint: Some(a.unescape_value()?.to_string()),
                            };
                        }
                        _ => (),
                    },
                    _ => (),
                }
            }
        }
        Ok(color)
    }

    /// Read either left, right, top, bottom, diagonal, vertical, or horizontal of borders
    fn read_border<B: BufRead>(
        xml: &mut Reader<B>,
        QName(mut closing): QName,
    ) -> Result<Border, XcelmateError> {
        fn read_region<B: BufRead>(
            xml: &mut Reader<B>,
            region: &BytesStart,
            border_region: &mut BorderRegion,
        ) -> Result<(), XcelmateError> {
            for attr in region.attributes() {
                if let Ok(a) = attr {
                    ////////////////////
                    // BORDER Attrs
                    /////////////
                    match a.key {
                        QName(b"style") => {
                            let val = a.unescape_value()?.to_string();
                            match val.as_str() {
                                "thin" => border_region.style = Some(BorderStyle::Thin),
                                "medium" => border_region.style = Some(BorderStyle::Medium),
                                "thick" => border_region.style = Some(BorderStyle::Thick),
                                "double" => border_region.style = Some(BorderStyle::Double),
                                "dashed" => border_region.style = Some(BorderStyle::Dashed),
                                "dotted" => border_region.style = Some(BorderStyle::Dotted),
                                "dashDot" => border_region.style = Some(BorderStyle::DashDot),
                                "dashDotDot" => border_region.style = Some(BorderStyle::DashDotDot),
                                "slantDashDot" => {
                                    border_region.style = Some(BorderStyle::SlantDashDot)
                                }
                                "hair" => border_region.style = Some(BorderStyle::Hair),
                                "mediumDashed" => {
                                    border_region.style = Some(BorderStyle::MediumDashed)
                                }
                                "mediumDashDot" => {
                                    border_region.style = Some(BorderStyle::MediumDashDot)
                                }
                                "mediumDashDotDot" => {
                                    border_region.style = Some(BorderStyle::MediumDashDotDot)
                                }
                                _ => (), // Ignore unsupported or unknown values
                            }
                        }
                        _ => (),
                    }
                }
            }
            let mut border_region_buf = Vec::with_capacity(1024);
            loop {
                border_region_buf.clear();
                match xml.read_event_into(&mut border_region_buf) {
                    ////////////////////
                    // BORDER (LRTB) nth-1
                    /////////////
                    Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"color" => {
                        border_region.color = Some(Stylesheet::read_color(e.attributes())?);
                    }
                    Ok(Event::End(ref e)) if e.local_name().as_ref() == region.name().as_ref() => {
                        return Ok(())
                    }
                    Ok(Event::Eof) => {
                        let mut name = String::new();
                        let _ = region.as_ref().read_to_string(&mut name)?;
                        return Err(XcelmateError::XmlEof(name));
                    }
                    Err(e) => return Err(XcelmateError::Xml(e)),
                    _ => (),
                }
            }
        }

        let mut buf = Vec::with_capacity(1024);
        let mut border = Border::default();
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"left" => {
                    read_region(xml, e, &mut border.left)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"right" => {
                    read_region(xml, e, &mut border.right)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"top" => {
                    read_region(xml, e, &mut border.top)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"bottom" => {
                    read_region(xml, e, &mut border.bottom)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"vertical" => {
                    read_region(xml, e, &mut border.vertical)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"diagonal" => {
                    read_region(xml, e, &mut border.diagonal)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"horizontal" => {
                    read_region(xml, e, &mut border.horizontal)?;
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == closing => return Ok(border),
                Ok(Event::Eof) => {
                    let mut name = String::new();
                    let _ = closing.read_to_string(&mut name)?;
                    return Err(XcelmateError::XmlEof(name));
                }
                _ => (),
            }
        }
    }

    /// Read font styling
    pub(crate) fn read_font<B: BufRead>(
        xml: &mut Reader<B>,
        QName(mut closing): QName,
    ) -> Result<FontProperty, XcelmateError> {
        let mut buf = Vec::with_capacity(1024);
        let mut font = FontProperty::default();
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                ////////////////////
                // FONT nth-1
                /////////////
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"sz" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => font.size = a.unescape_value()?.to_string(),
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"strike" => {
                    font.strikethrough = FormatState::Enabled;
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => match a.unescape_value()?.to_string().as_str() {
                                    "0" => {
                                        font.strikethrough = FormatState::Disabled;
                                    }
                                    _ => (),
                                },
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"outline" => {
                    font.outline = FormatState::Enabled;
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => match a.unescape_value()?.to_string().as_str() {
                                    "0" => {
                                        font.outline = FormatState::Disabled;
                                    }
                                    _ => (),
                                },
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"shadow" => {
                    font.shadow = FormatState::Enabled;
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => match a.unescape_value()?.to_string().as_str() {
                                    "0" => {
                                        font.shadow = FormatState::Disabled;
                                    }
                                    _ => (),
                                },
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"vertAlign" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => match a.unescape_value()?.to_string().as_str() {
                                    "subscript" => {
                                        font.subscript = FormatState::Enabled;
                                    }
                                    "superscript" => font.superscript = FormatState::Enabled,
                                    "baseline" => font.baseline = FormatState::Enabled,
                                    _ => (),
                                },
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"b" => {
                    font.bold = FormatState::Enabled;
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => match a.unescape_value()?.to_string().as_str() {
                                    "0" => {
                                        font.bold = FormatState::Disabled;
                                    }
                                    _ => (),
                                },
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"i" => {
                    font.italic = FormatState::Enabled;
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => match a.unescape_value()?.to_string().as_str() {
                                    "0" => {
                                        font.italic = FormatState::Disabled;
                                    }
                                    _ => (),
                                },
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"u" => {
                    // we do not know if underline is set to not show so we set it to true incase we encountee nonr in attributes
                    font.underline = FormatState::Enabled;
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => {
                                    match a.unescape_value()?.to_string().as_str() {
                                        "double" => {
                                            font.double = FormatState::Enabled;
                                            // No longer can be true if doubled
                                            font.underline = FormatState::None;
                                        }
                                        "none" => font.underline = FormatState::Disabled,
                                        _ => (),
                                    }
                                }
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"color" => {
                    font.color = Stylesheet::read_color(e.attributes())?;
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"name" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => font.font = a.unescape_value()?.to_string(),
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"rFont" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => font.font = a.unescape_value()?.to_string(),
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"family" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => {
                                    font.family = a.unescape_value()?.parse::<u32>()?
                                }
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"scheme" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => font.scheme = a.unescape_value()?.to_string(),
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == closing => return Ok(font),
                Ok(Event::Eof) => {
                    let mut name = String::new();
                    let _ = closing.read_to_string(&mut name)?;
                    return Err(XcelmateError::XmlEof(name));
                }
                Err(e) => return Err(XcelmateError::Xml(e)),
                _ => (),
            }
        }
    }

    /// Read fill styling
    fn read_fill<B: BufRead>(
        xml: &mut Reader<B>,
        QName(mut closing): QName,
    ) -> Result<Fill, XcelmateError> {
        let mut buf = Vec::with_capacity(1024);
        let mut fill = Fill::default();
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                ////////////////////
                // FILL nth-1
                /////////////
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e))
                    if e.local_name().as_ref() == b"patternFill" =>
                {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"patternType") => {
                                    let val = a.unescape_value()?.to_string();
                                    match val.as_str() {
                                        "solid" => fill.r#type = PatternFill::Solid,
                                        "none" => fill.r#type = PatternFill::None,
                                        "gray125" => fill.r#type = PatternFill::Gray,
                                        _ => (),
                                    }
                                }
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"fgColor" => {
                    fill.foreground = Some(Stylesheet::read_color(e.attributes())?);
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"bgColor" => {
                    fill.background = Some(Stylesheet::read_color(e.attributes())?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == closing => return Ok(fill),
                Ok(Event::Eof) => {
                    let mut name = String::new();
                    let _ = closing.read_to_string(&mut name)?;
                    return Err(XcelmateError::XmlEof(name));
                }
                Err(e) => return Err(XcelmateError::Xml(e)),
                _ => (),
            }
        }
    }

    /// Convert from hexadecimal to a tuple of RGB model
    pub(crate) fn to_rgb(value: String) -> Result<Color, XcelmateError> {
        // The first two letter are ignored since they response to alpha
        let base16 = 16u32;
        let red = u8::from_str_radix(&value[2..4], base16)?;
        let green = u8::from_str_radix(&value[4..6], base16)?;
        let blue = u8::from_str_radix(&value[6..8], base16)?;
        Ok(Color::Rgb(Rgb::Custom(red, green, blue)))
    }

    /// Convert from u8 to a hexadecimal of RGB model scale
    pub(crate) fn from_rgb(r: u8, g: u8, b: u8) -> String {
        format!("{:02X}", r) + &format!("{:02X}", g) + &format!("{:02X}", b)
    }
}

#[cfg(test)]
mod stylesheet_unittests {
    use super::Stylesheet;
    use std::fs::File;
    use zip::ZipArchive;

    fn init(path: &str) -> Stylesheet {
        let file = File::open(path).unwrap();
        let mut zip = ZipArchive::new(file).unwrap();
        let mut stylesheet = Stylesheet::default();
        stylesheet.read_stylesheet(&mut zip).unwrap();
        stylesheet
    }

    mod stylesheet_api {
        use super::init;
        use crate::stream::utils::Save;
        use crate::stream::xlsx::stylesheet::{
            Border, BorderRegion, BorderStyle, CellXf, DiffXf, Fill, FontProperty, FormatState,
            NumberFormat, PatternFill,
        };
        use crate::stream::xlsx::{
            stylesheet::{Color, Rgb},
            Stylesheet,
        };
        use quick_xml::{events::Event, Reader};
        use std::fs::File;
        use std::io::Cursor;
        use std::sync::Arc;
        use zip::write::SimpleFileOptions;
        use zip::{CompressionMethod, ZipWriter};

        #[test]
        fn get_custom_table_style() {
            let style = init("tests/workbook04.xlsx");
            let actual = style.get_custom_table_style("Customer Contact List");
            assert!(actual.is_some())
        }

        #[test]
        fn test_to_rgb() {
            let result = Stylesheet::to_rgb("FF573345".into()).unwrap();
            assert_eq!(result, Color::Rgb(Rgb::Custom(87, 51, 69)));
        }

        #[test]
        fn test_read_border_region_for_empty_borders() {
            let xml_content = r#"
                <root>
                    <border>
                        <left></left>
                        <right></right>
                        <top></top>
                        <bottom></bottom>
                    </border>
                </root>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                        let border = Stylesheet::read_border(&mut xml, e.name()).unwrap();

                        assert_eq!(border.left, BorderRegion::default());

                        assert_eq!(border.right, BorderRegion::default());

                        assert_eq!(border.top, BorderRegion::default());

                        assert_eq!(border.bottom, BorderRegion::default());
                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_border_region_for_all_borders() {
            let xml_content = r#"
                <root>
                    <border>
                        <left style="double">
                            <color rgb="FF234567" />
                        </left>
                        <right style="thick">
                            <color rgb="FF234567" />
                        </right>
                        <top style="thin">
                            <color rgb="FF234567" />
                        </top>
                        <bottom style="dashed">
                            <color theme="1" tint="0.78785898899" />
                        </bottom>
                        <vertical style="dashed">
                            <color theme="2" tint="0.78785898899" />
                        </vertical>
                        <horizontal style="dashed">
                            <color theme="3" tint="0.78785898899" />
                        </horizontal>
                        <diagonal style="dashed">
                            <color theme="4" tint="0.78785898899" />
                        </diagonal>
                    </border>
                </root>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                        let border = Stylesheet::read_border(&mut xml, e.name()).unwrap();
                        assert_eq!(
                            border.left,
                            BorderRegion {
                                style: Some(BorderStyle::Double),
                                color: Some(Color::Rgb(Rgb::Custom(35, 69, 103)))
                            }
                        );
                        assert_eq!(
                            border.right,
                            BorderRegion {
                                style: Some(BorderStyle::Thick),
                                color: Some(Color::Rgb(Rgb::Custom(35, 69, 103)))
                            }
                        );
                        assert_eq!(
                            border.top,
                            BorderRegion {
                                style: Some(BorderStyle::Thin),
                                color: Some(Color::Rgb(Rgb::Custom(35, 69, 103)))
                            }
                        );
                        assert_eq!(
                            border.bottom,
                            BorderRegion {
                                style: Some(BorderStyle::Dashed),
                                color: Some(Color::Theme {
                                    id: 1,
                                    tint: Some("0.78785898899".into())
                                })
                            }
                        );
                        assert_eq!(
                            border.vertical,
                            BorderRegion {
                                style: Some(BorderStyle::Dashed),
                                color: Some(Color::Theme {
                                    id: 2,
                                    tint: Some("0.78785898899".into())
                                })
                            }
                        );
                        assert_eq!(
                            border.horizontal,
                            BorderRegion {
                                style: Some(BorderStyle::Dashed),
                                color: Some(Color::Theme {
                                    id: 3,
                                    tint: Some("0.78785898899".into())
                                })
                            }
                        );
                        assert_eq!(
                            border.diagonal,
                            BorderRegion {
                                style: Some(BorderStyle::Dashed),
                                color: Some(Color::Theme {
                                    id: 4,
                                    tint: Some("0.78785898899".into())
                                })
                            }
                        );
                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_border_region_for_malformed_xml() {
            let xml_content = r#"
                <root>
                    <border>
                        <horizontal style="double">
                            <color rgb="FF234567" />
                        </horizontal
                    </border>
                </root>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                        let actual = Stylesheet::read_border(&mut xml, e.name())
                            .err()
                            .unwrap()
                            .to_string();
                        assert_eq!(actual, "ill-formed document: expected `</horizontal>`, but `</horizontal\n                    </border>` was found".to_string());
                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_border_region_for_eof() {
            let xml_content = r#"
                <root>
                    <border>
                    <vertical>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                        let actual = Stylesheet::read_border(&mut xml, e.name())
                            .err()
                            .unwrap()
                            .to_string();
                        assert_eq!(actual, "malformed stream for tag: vertical".to_string());
                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_font_for_malformed_xml() {
            let xml_content = r#"
                <root>
                    <fonts>
                        <font>
                            <sz val="12" />
                        </font
                    </fonts>
                </root>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                        let actual = Stylesheet::read_font(&mut xml, e.name())
                            .err()
                            .unwrap()
                            .to_string();
                        assert_eq!(actual, "ill-formed document: expected `</font>`, but `</font\n                    </fonts>` was found".to_string());
                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_font_for_eof() {
            let xml_content = r#"
                <root>
                    <fonts>
                        <font>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                        let actual = Stylesheet::read_font(&mut xml, e.name())
                            .err()
                            .unwrap()
                            .to_string();
                        assert_eq!(actual, "malformed stream for tag: font".to_string());
                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_font_for_all_styling() {
            let xml_content = r#"
                <root>
                    <fonts>
                        <font>
                            <b/>
                            <i/>
                            <u val="double"/>
                            <color theme="1"/>
                            <sz val="21"/>
                            <name val="Calibri"/>
                            <family val="2"/>
                            <scheme val="minor"/>
                        </font>
                    </fonts>
                </root>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                        let actual = Stylesheet::read_font(&mut xml, e.name()).unwrap();
                        assert_eq!(
                            actual,
                            FontProperty {
                                bold: FormatState::Enabled,
                                double: FormatState::Enabled,
                                italic: FormatState::Enabled,
                                size: "21".into(),
                                color: Color::Theme { id: 1, tint: None },
                                font: "Calibri".into(),
                                family: 2,
                                scheme: "minor".into(),
                                ..Default::default()
                            }
                        );

                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_font_for_no_underline() {
            let xml_content = r#"
                <root>
                    <fonts>
                        <font>
                            <b/>
                            <i/>
                            <u val="none"/>
                            <color theme="1"/>
                            <sz val="21"/>
                            <name val="Calibri"/>
                            <family val="2"/>
                            <scheme val="minor"/>
                        </font>
                    </fonts>
                </root>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                        let actual = Stylesheet::read_font(&mut xml, e.name()).unwrap();
                        assert_eq!(
                            actual,
                            FontProperty {
                                bold: FormatState::Enabled,
                                italic: FormatState::Enabled,
                                underline: FormatState::Disabled,
                                size: "21".into(),
                                color: Color::Theme { id: 1, tint: None },
                                font: "Calibri".into(),
                                family: 2,
                                scheme: "minor".into(),
                                ..Default::default()
                            }
                        );

                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_font_for_underline() {
            let xml_content = r#"
                <root>
                    <fonts>
                        <font>
                            <b/>
                            <i/>
                            <u/>
                            <color theme="1"/>
                            <sz val="21"/>
                            <name val="Calibri"/>
                            <family val="2"/>
                            <scheme val="minor"/>
                        </font>
                    </fonts>
                </root>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                        let actual = Stylesheet::read_font(&mut xml, e.name()).unwrap();
                        assert_eq!(
                            actual,
                            FontProperty {
                                bold: FormatState::Enabled,
                                underline: FormatState::Enabled,
                                italic: FormatState::Enabled,
                                size: "21".into(),
                                color: Color::Theme { id: 1, tint: None },
                                font: "Calibri".into(),
                                family: 2,
                                scheme: "minor".into(),
                                ..Default::default()
                            }
                        );

                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_fill_for_malformed_xml() {
            let xml_content = r#"
                <root>
                    <fills>
                        <fill>
                            <patternFill patternType="none" />
                        </fill
                    </fills>
                </root>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fills" => {
                        let actual = Stylesheet::read_fill(&mut xml, e.name())
                            .err()
                            .unwrap()
                            .to_string();
                        assert_eq!(actual, "ill-formed document: expected `</fill>`, but `</fill\n                    </fills>` was found".to_string());
                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_fill_for_eof() {
            let xml_content = r#"
                <root>
                    <fills>
                        <fill>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fills" => {
                        let actual = Stylesheet::read_fill(&mut xml, e.name())
                            .err()
                            .unwrap()
                            .to_string();
                        assert_eq!(actual, "malformed stream for tag: fills".to_string());
                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_fill_for_type_none() {
            let xml_content = r#"
                <root>
                    <fills>
                        <fill>
                            <patternFill patternType="none" />
                        </fill>
                    </fills>
                </root>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                        let actual = Stylesheet::read_fill(&mut xml, e.name()).unwrap();
                        assert_eq!(
                            actual,
                            Fill {
                                r#type: PatternFill::None,
                                foreground: None,
                                background: None
                            }
                        );

                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_fill_for_type_gray() {
            let xml_content = r#"
                <root>
                    <fills>
                        <fill>
                            <patternFill patternType="gray125" />
                        </fill>
                    </fills>
                </root>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                        let actual = Stylesheet::read_fill(&mut xml, e.name()).unwrap();
                        assert_eq!(
                            actual,
                            Fill {
                                r#type: PatternFill::Gray,
                                foreground: None,
                                background: None
                            }
                        );

                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_read_fill_for_type_solid() {
            let xml_content = r#"
                <root>
                    <fills>
                        <fill>
                            <patternFill patternType="solid">
                                <fgColor rgb="FF435678"/>
                                <bgColor rgb="FF432378"/>
                            </patternFill>
                        </fill>
                    </fills>
                </root>
                "#;
            let mut xml = Reader::from_reader(Cursor::new(xml_content));
            let mut buf = Vec::with_capacity(1024);

            loop {
                match xml.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                        let actual = Stylesheet::read_fill(&mut xml, e.name()).unwrap();
                        assert_eq!(
                            actual,
                            Fill {
                                r#type: PatternFill::Solid,
                                foreground: Some(Color::Rgb(Rgb::Custom(67, 86, 120))),
                                background: Some(Color::Rgb(Rgb::Custom(67, 35, 120)))
                            }
                        );

                        break;
                    }
                    _ => (),
                }
            }
        }

        #[test]
        fn test_get_cell_ref_from_key_and_not_exists() {
            let style = init("tests/workbook03.xlsx");
            let actual = style.get_cell_ref_from_key(29);
            assert_eq!(actual, None)
        }

        #[test]
        fn test_get_cell_ref_from_key_and_exists() {
            let style = init("tests/workbook03.xlsx");
            let actual = style.get_cell_ref_from_key(1).unwrap();
            let actual_key = style.get_key_from_cell_ref(actual.clone()).unwrap();
            assert_eq!(actual_key, 1);
            assert_eq!(
                actual,
                CellXf {
                    number_format: None,
                    font: Arc::new(FontProperty {
                        size: "11".into(),
                        color: Color::Rgb(Rgb::Custom(156, 0, 6,)),
                        font: "Calibri".into(),
                        family: 2,
                        scheme: "minor".into(),
                        ..Default::default()
                    }),
                    fill: Arc::new(Fill {
                        r#type: PatternFill::Solid,
                        foreground: Some(Color::Rgb(Rgb::Custom(255, 199, 206))),
                        ..Default::default()
                    }),
                    border: Arc::new(Border {
                        ..Default::default()
                    }),
                    ..Default::default()
                }
                .into()
            );
        }

        #[test]
        fn test_get_differential_ref_from_key_and_exists() {
            let style = init("tests/workbook04.xlsx");
            let actual = style.get_differential_ref_from_key(1).unwrap();
            let actual_key = style.get_key_from_differential_ref(actual.clone()).unwrap();
            assert_eq!(actual_key, 1);
            assert_eq!(
                actual,
                DiffXf {
                    font: Some(FontProperty {
                        strikethrough: FormatState::Disabled,
                        outline: FormatState::Disabled,
                        shadow: FormatState::Disabled,
                        baseline: FormatState::Enabled,
                        underline: FormatState::Disabled,
                        size: "11".into(),
                        color: Color::Theme { id: 0, tint: None },
                        font: "Posterama".into(),
                        family: 2,
                        scheme: "major".into(),
                        ..Default::default()
                    }),
                    dup_cnt: 1, // verifies duplicates allowed

                    ..Default::default()
                }
                .into()
            );
        }

        #[test]
        fn test_get_differential_ref_from_key_and_not_exists() {
            let style = init("tests/workbook04.xlsx");
            let actual = style.get_differential_ref_from_key(11);
            assert_eq!(actual, None)
        }

        #[test]
        fn test_get_number_format_ref_from_key_and_exists() {
            let style = init("tests/workbook03.xlsx");
            let actual = style.get_number_format_ref_from_key(43);
            let actual_key = style
                .get_key_from_number_format_ref(actual.clone().unwrap())
                .unwrap();
            assert_eq!(actual_key, 43);
            assert_eq!(
                actual,
                Some(Arc::new(NumberFormat {
                    id: 43,
                    format_code: r#"_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)"#.into()
                }))
            )
        }

        #[test]
        fn test_get_number_format_ref_from_key_and_not_exists() {
            let style = init("tests/workbook04.xlsx");
            let actual = style.get_number_format_ref_from_key(11);
            assert_eq!(actual, None)
        }

        #[test]
        fn test_get_font_ref_from_key_and_exists() {
            let style = init("tests/workbook03.xlsx");
            let actual = style.get_font_ref_from_key(3);
            let actual_key = style
                .get_key_from_font_ref(actual.clone().unwrap())
                .unwrap();
            assert_eq!(actual_key, 3);
            assert_eq!(
                actual,
                Some(Arc::new(FontProperty {
                    size: "18".into(),
                    color: Color::Theme { id: 3, tint: None },
                    font: "Calibri Light".into(),
                    family: 2,
                    scheme: "major".into(),
                    ..Default::default()
                }))
            )
        }

        #[test]
        fn test_get_font_ref_from_key_and_not_exists() {
            let style = init("tests/workbook03.xlsx");
            let actual = style.get_font_ref_from_key(30);
            assert_eq!(actual, None)
        }

        #[test]
        fn test_get_fill_ref_from_key_and_exists() {
            let style = init("tests/workbook03.xlsx");
            let actual = style.get_fill_ref_from_key(3);
            let actual_key = style
                .get_key_from_fill_ref(actual.clone().unwrap())
                .unwrap();
            assert_eq!(actual_key, 3);
            assert_eq!(
                actual,
                Some(Arc::new(Fill {
                    r#type: PatternFill::Solid,
                    foreground: Some(Color::Rgb(Rgb::Custom(255, 199, 206))),
                    background: None
                }))
            )
        }

        #[test]
        fn test_get_fill_ref_from_key_and_not_exists() {
            let style = init("tests/workbook03.xlsx");
            let actual = style.get_fill_ref_from_key(30);
            assert_eq!(actual, None)
        }

        #[test]
        fn test_get_border_ref_from_key_and_exists() {
            let style = init("tests/workbook03.xlsx");
            let actual = style.get_border_ref_from_key(3);
            let actual_key = style
                .get_key_from_border_ref(actual.clone().unwrap())
                .unwrap();
            assert_eq!(actual_key, 3);
            assert_eq!(
                actual,
                Some(Arc::new(Border {
                    bottom: BorderRegion {
                        style: Some(BorderStyle::Medium),
                        color: Some(Color::Theme {
                            id: 4,
                            tint: Some("0.39997558519241921".into())
                        })
                    },
                    ..Default::default()
                }))
            )
        }

        #[test]
        fn test_get_border_ref_from_key_and_not_exists() {
            let style = init("tests/workbook03.xlsx");
            let actual = style.get_border_ref_from_key(30);
            assert_eq!(actual, None)
        }

        #[test]
        fn save_file() {
            let mut style = init("tests/workbook04.xlsx");
            let mut zip = ZipWriter::new(Cursor::new(Vec::<u8>::new()));
            style
                .save(
                    &mut zip,
                    SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
                )
                .unwrap();

            // Verify all data is written
            assert!(zip.finish().unwrap().into_inner().len() > 22);
        }
    }
}
