use crate::{
    errors::XcelmateError,
    stream::utils::{xml_reader, Key},
};
use bimap::BiMap;
use quick_xml::{events::Event, name::QName, Reader};
use std::{
    collections::HashMap,
    io::{BufReader, Read, Seek},
    sync::Arc,
};
use zip::{read::ZipFile, ZipArchive};

/// The `Rgb` promotes better api usage with hexadecimal coloring
#[derive(Debug, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) enum Rgb {
    Custom((u8, u8, u8)),
}

/// The `Color` denotes the type of coloring system to
/// use since excel has builtin coloring to choose that will map to `theme` but
/// for custom specfic coloring `rgb` is used
///
/// Default is `Theme((1, None))` = black
#[derive(Debug, PartialEq, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) enum Color {
    /// Builtin theme from excel color palette selector which includes theme id and tint value
    Theme { id: u32, tint: Option<String> },
    /// RGB color model
    Rgb(Rgb),
}
impl Default for Color {
    fn default() -> Self {
        Color::Theme { id: 1, tint: None }
    }
}

/// The `FontProperty` denotes all styling options
/// that can be added to text
#[derive(Debug, Default, Clone, PartialEq, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct FontProperty {
    pub(crate) bold: bool,
    pub(crate) underline: bool,
    /// Double underline
    pub(crate) double: bool,
    pub(crate) italic: bool,
    pub(crate) size: String,
    pub(crate) color: Color,
    /// Font type
    pub(crate) font: String,
    /// Font family
    pub(crate) family: u32,
    /// Font scheme
    pub(crate) scheme: String,
}

/// The formatting style to use on numbers
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
struct NumberFormat {
    id: u32,
    format_code: FormatType,
}

/// The enum helps determine what to include in final write to file for number formats
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
enum FormatType {
    /// A builtin in number format code that will not appear in <numfmt> tag list
    #[default]
    Builtin,
    Custom(String),
}

/// The pattern fill styling to apply to a cell
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
enum PatternFill {
    #[default]
    None,
    Solid,
    Gray,
}

/// The background/foreground fill of a cell. Also can include gradients
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct Fill {
    r#type: PatternFill,
    foreground: Option<Color>,
    background: Option<Color>,
}

/// The type of line styling for a border
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
enum BorderStyle {
    /// No border
    #[default]
    None,
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

/// The border region to apply styling to
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
struct BorderRegion {
    style: BorderStyle,
    color: Color,
}

/// The styling for all border regions of a cell
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct Border {
    left: BorderRegion,
    right: BorderRegion,
    top: BorderRegion,
    bottom: BorderRegion,
}

/// The styling traits of a cell
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct CellXf {
    number_format: Option<Arc<NumberFormat>>,
    font: Arc<FontProperty>,
    fill: Arc<Fill>,
    border: Arc<Border>,
}

/// The styling groups for differential conditional formatting
#[derive(Debug, PartialEq, Default, Clone, Eq, PartialOrd, Hash, Ord)]
pub(crate) struct DiffXf {
    font: Option<FontProperty>,
    fill: Option<Fill>,
    border: Option<Border>,
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
    number_formats: Option<BiMap<Arc<NumberFormat>, Key>>,
    fonts: BiMap<Arc<FontProperty>, Key>,
    fills: BiMap<Arc<Fill>, Key>,
    borders: BiMap<Arc<Border>, Key>,
    cell_xf: BiMap<Arc<CellXf>, Key>,
    diff_xf: HashMap<Arc<DiffXf>, Key>,
    diff_xf_with_dups: Vec<Arc<DiffXf>>, // Duplicates can exist
    table_style: Option<TableStyle>,
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
        let mut skipped_first_font = false;
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"numFmts" => {
                    self.number_formats = Some(BiMap::new());
                    // Preinsert builtin number formats to prevent failed lookup
                    self.number_formats.as_mut().unwrap().insert(
                        NumberFormat {
                            id: 9,
                            format_code: FormatType::Builtin,
                        }
                        .into(),
                        9,
                    );
                }
                ////////////////////
                // NUMBER FORMATS
                /////////////
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"numFmt" => {
                    if let Some(n) = self.number_formats.as_mut() {
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
                                        numfmt.format_code =
                                            FormatType::Custom(a.unescape_value()?.to_string())
                                    }
                                    _ => (),
                                }
                            }
                        }
                        let key = numfmt.id as usize;
                        let numfmt = Arc::new(numfmt);
                        n.insert(numfmt, key);
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fonts" => {
                    // Preset the first default font to be unique to avoid duplicate overwritting since first font is always there
                    let font = Arc::new(FontProperty::default());
                    self.fonts.insert(font, 0);
                }
                ////////////////////
                // FONT
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                    let font = Stylesheet::read_font(&mut xml, e.name())?;
                    // Skip the first default font since we dont care about it and keep correct index mapping
                    if skipped_first_font {
                        let key = self.fonts.len();
                        let font = Arc::new(font);
                        self.fonts.insert(font, key);
                    } else {
                        skipped_first_font = true
                    }
                }
                ////////////////////
                // FILL
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                    let fill = Stylesheet::read_fill(&mut xml)?;
                    let key = self.fills.len();
                    let fill = Arc::new(fill);
                    self.fills.insert(fill, key);
                }
                ////////////////////
                // BORDER
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                    let mut border = Border::default();
                    border.left = Stylesheet::read_border_region(&mut xml, QName(b"left"))?;
                    border.right = Stylesheet::read_border_region(&mut xml, QName(b"right"))?;
                    border.top = Stylesheet::read_border_region(&mut xml, QName(b"top"))?;
                    border.bottom = Stylesheet::read_border_region(&mut xml, QName(b"bottom"))?;

                    let key = self.borders.len();
                    let border = Arc::new(border);
                    self.borders.insert(border, key);
                }
                ////////////////////
                // CELL REFERENCES
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cellXfs" => {
                    let mut cell_xf_buf = Vec::with_capacity(1024);
                    loop {
                        cell_xf_buf.clear();
                        let mut cell_xf = CellXf::default();
                        match xml.read_event_into(&mut cell_xf_buf) {
                            ////////////////////
                            // CELL REFERENCES nth-1
                            /////////////
                            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"xf" => {
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
                                            _ => (),
                                        }
                                    }
                                }
                                let key = self.cell_xf.len();
                                let cell_xf = Arc::new(cell_xf);
                                self.cell_xf.insert(cell_xf, key);
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cellXfs" => break,
                            Ok(Event::Eof) => return Err(XcelmateError::XmlEof("cellXfs".into())),
                            Err(e) => return Err(XcelmateError::Xml(e)),
                            _ => (),
                        }
                    }
                }
                ////////////////////
                // TABLE CUSTOM REFERENCE
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dxf" => {
                    let mut dx_buf = Vec::with_capacity(1024);
                    let mut diff_xf = DiffXf::default();
                    loop {
                        dx_buf.clear();
                        match xml.read_event_into(&mut dx_buf) {
                            ////////////////////
                            // TABLE CUSTOM REFERENCE nth-1
                            /////////////
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                                diff_xf.font = Some(Stylesheet::read_font(&mut xml, e.name())?);
                            }
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                                let mut border = Border::default();
                                border.left =
                                    Stylesheet::read_border_region(&mut xml, QName(b"left"))?;
                                border.right =
                                    Stylesheet::read_border_region(&mut xml, QName(b"right"))?;
                                border.top =
                                    Stylesheet::read_border_region(&mut xml, QName(b"top"))?;
                                border.bottom =
                                    Stylesheet::read_border_region(&mut xml, QName(b"bottom"))?;
                                diff_xf.border = Some(border);
                            }
                            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                                diff_xf.fill = Some(Stylesheet::read_fill(&mut xml)?);
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dxf" => break,
                            Ok(Event::Eof) => return Err(XcelmateError::XmlEof("dxf".into())),
                            Err(e) => return Err(XcelmateError::Xml(e)),
                            _ => (),
                        }
                    }
                    let diff_xf = Arc::new(diff_xf);

                    // Since duplication can occur with custom tabling styles we must use a vector/map instead of bimap
                    let key = self.diff_xf_with_dups.len();
                    self.diff_xf_with_dups.push(diff_xf.clone());

                    self.diff_xf.insert(diff_xf, key);
                }
                ////////////////////
                // TABLE STYLE
                /////////////
                Ok(Event::Empty(ref e))
                    if e.local_name().as_ref() == b"tableStyles" =>
                {
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
                    Ok(Event::Start(ref e))
                    if e.local_name().as_ref() == b"tableStyles" =>
                {
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
                                    .insert(custom_style.name.clone(), Arc::new(custom_style));
                                // Reset style interation
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
                    self.table_style = Some(table_style);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"styleSheet" => break,
                Ok(Event::Eof) => return Err(XcelmateError::XmlEof("styleSheet".into())),
                Err(e) => return Err(XcelmateError::Xml(e)),
                _ => (),
            }
        }
        Ok(())
    }

    /// Get custom table style
    pub(crate) fn get_custom_table_style(&mut self, name: &str) -> Option<Arc<TableCustomStyle>> {
        if let Some(t) = &self.table_style {
            t.styles.get(name).cloned()
        } else {
            None
        }
    }

    /// Get the cell format ref from key
    pub(crate) fn get_cell_ref_from_key(&mut self, key: Key) -> Option<Arc<CellXf>> {
        if let Some(i) = self.cell_xf.get_by_right(&key) {
            Some(i.clone())
        } else {
            None
        }
    }

    /// Get the differential format ref from key
    pub(crate) fn get_differential_ref_from_key(&mut self, key: Key) -> Option<Arc<DiffXf>> {
        if let Some(i) = self.diff_xf_with_dups.get(key) {
            Some(i.clone())
        } else {
            None
        }
    }

    /// Get the number format ref from key
    pub(crate) fn get_number_format_ref_from_key(&mut self, key: Key) -> Option<Arc<NumberFormat>> {
        if let Some(n) = &self.number_formats {
            if let Some(i) = n.get_by_right(&key) {
                Some(i.clone())
            } else {
                None
            }
        } else {
            None
        }
    }

    /// Get the font format ref from key
    pub(crate) fn get_font_ref_from_key(&mut self, key: Key) -> Option<Arc<FontProperty>> {
        if let Some(i) = self.fonts.get_by_right(&key) {
            Some(i.clone())
        } else {
            None
        }
    }

    /// Get the fill format ref from key
    pub(crate) fn get_fill_ref_from_key(&mut self, key: Key) -> Option<Arc<Fill>> {
        if let Some(i) = self.fills.get_by_right(&key) {
            Some(i.clone())
        } else {
            None
        }
    }
    /// Get the border format ref from key
    pub(crate) fn get_border_ref_from_key(&mut self, key: Key) -> Option<Arc<Border>> {
        if let Some(i) = self.borders.get_by_right(&key) {
            Some(i.clone())
        } else {
            None
        }
    }

    /// Read either left, right, top, or bottom of borders
    fn read_border_region(
        xml: &mut Reader<BufReader<ZipFile>>,
        QName(mut closing): QName,
    ) -> Result<BorderRegion, XcelmateError> {
        let mut buf = Vec::with_capacity(1024);
        let mut border_region = BorderRegion::default();
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                ////////////////////
                // BORDER (LRTB)
                /////////////
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == closing => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            ////////////////////
                            // BORDER (LRTB) Attrs
                            /////////////
                            match a.key {
                                QName(b"style") => {
                                    let val = a.unescape_value()?.to_string();
                                    match val.as_str() {
                                        "none" => border_region.style = BorderStyle::None,
                                        "thin" => border_region.style = BorderStyle::Thin,
                                        "medium" => border_region.style = BorderStyle::Medium,
                                        "thick" => border_region.style = BorderStyle::Thick,
                                        "double" => border_region.style = BorderStyle::Double,
                                        "dashed" => border_region.style = BorderStyle::Dashed,
                                        "dotted" => border_region.style = BorderStyle::Dotted,
                                        "dashDot" => border_region.style = BorderStyle::DashDot,
                                        "dashDotDot" => {
                                            border_region.style = BorderStyle::DashDotDot
                                        }
                                        "slantDashDot" => {
                                            border_region.style = BorderStyle::SlantDashDot
                                        }
                                        "hair" => border_region.style = BorderStyle::Hair,
                                        "mediumDashed" => {
                                            border_region.style = BorderStyle::MediumDashed
                                        }
                                        "mediumDashDot" => {
                                            border_region.style = BorderStyle::MediumDashDot
                                        }
                                        "mediumDashDotDot" => {
                                            border_region.style = BorderStyle::MediumDashDotDot
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
                                for attr in e.attributes() {
                                    if let Ok(a) = attr {
                                        match a.key {
                                            QName(b"rgb") => {
                                                border_region.color = Stylesheet::to_rgb(
                                                    a.unescape_value()?.to_string(),
                                                )?
                                            }
                                            QName(b"theme") => {
                                                border_region.color = Color::Theme {
                                                    id: a.unescape_value()?.parse::<u32>()?,
                                                    tint: None,
                                                };
                                            }
                                            QName(b"tint") => match border_region.color {
                                                Color::Theme { id, .. } => {
                                                    border_region.color = Color::Theme {
                                                        id,
                                                        tint: Some(a.unescape_value()?.to_string()),
                                                    }
                                                }
                                                _ => (),
                                            },
                                            _ => (),
                                        }
                                    }
                                }
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == closing => {
                                return Ok(border_region)
                            }
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
                Ok(Event::Empty(_)) => return Ok(border_region),
                _ => (),
            }
        }
    }

    /// Read font styling
    pub(crate) fn read_font(
        xml: &mut Reader<BufReader<ZipFile>>,
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
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"b" => font.bold = true,
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"i" => font.italic = true,
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"u" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"val") => {
                                    font.double = true;
                                    // No longer can be true if doubled
                                    font.underline = false;
                                }
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"color" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"rgb") => {
                                    font.color =
                                        Stylesheet::to_rgb(a.unescape_value()?.to_string())?
                                }
                                QName(b"theme") => {
                                    font.color = Color::Theme {
                                        id: a.unescape_value()?.parse::<u32>()?,
                                        tint: None,
                                    };
                                }
                                QName(b"tint") => match font.color {
                                    Color::Theme { id, .. } => {
                                        font.color = Color::Theme {
                                            id,
                                            tint: Some(a.unescape_value()?.to_string()),
                                        }
                                    }
                                    _ => (),
                                },
                                _ => (),
                            }
                        }
                    }
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
    fn read_fill(xml: &mut Reader<BufReader<ZipFile>>) -> Result<Fill, XcelmateError> {
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
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"rgb") => {
                                    fill.foreground =
                                        Some(Stylesheet::to_rgb(a.unescape_value()?.to_string())?)
                                }
                                QName(b"theme") => {
                                    fill.foreground = Some(Color::Theme {
                                        id: a.unescape_value()?.parse::<u32>()?,
                                        tint: None,
                                    });
                                }
                                QName(b"tint") => match fill.foreground {
                                    Some(Color::Theme { id, .. }) => {
                                        fill.foreground = Some(Color::Theme {
                                            id,
                                            tint: Some(a.unescape_value()?.to_string()),
                                        })
                                    }
                                    _ => (),
                                },
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"bgColor" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                QName(b"rgb") => {
                                    fill.background =
                                        Some(Stylesheet::to_rgb(a.unescape_value()?.to_string())?);
                                }
                                QName(b"theme") => {
                                    fill.background = Some(Color::Theme {
                                        id: a.unescape_value()?.to_string().parse::<u32>()?,
                                        tint: None,
                                    });
                                }
                                QName(b"tint") => match fill.background {
                                    Some(Color::Theme { id, .. }) => {
                                        fill.background = Some(Color::Theme {
                                            id,
                                            tint: Some(a.unescape_value()?.to_string()),
                                        })
                                    }
                                    _ => (),
                                },
                                _ => (),
                            }
                        }
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fill" => return Ok(fill),
                Ok(Event::Eof) => return Err(XcelmateError::XmlEof("fill".into())),
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
        Ok(Color::Rgb(Rgb::Custom((red, green, blue))))
    }
}

#[cfg(test)]
mod stylesheet_edges {
    use super::Stylesheet;
    use crate::stream::xlsx::stylesheet::{Color, FontProperty};
    use std::fs::File;
    use zip::ZipArchive;

    fn init(path: &str) -> Stylesheet {
        let file = File::open(path).unwrap();
        let mut zip = ZipArchive::new(file).unwrap();
        let mut stylesheet = Stylesheet::default();
        stylesheet.read_stylesheet(&mut zip).unwrap();
        stylesheet
    }

    #[test]
    fn first_default_font_should_be_skipped() {
        let mut style = init("tests/workbook03.xlsx");
        let font = style.get_font_ref_from_key(2).unwrap();
        assert_eq!(
            *font,
            FontProperty {
                size: "11".into(),
                color: Color::Theme { id: 1, tint: None },
                font: "Calibri".into(),
                family: 2,
                scheme: "minor".into(),
                ..Default::default()
            }
        );
    }
}
