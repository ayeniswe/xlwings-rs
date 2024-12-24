use crate::{
    errors::XcelmateError,
    stream::utils::{xml_reader, Key},
};
use quick_xml::{events::Event, name::QName};
use std::{
    collections::HashMap,
    io::{Read, Seek},
    sync::Arc,
};
use zip::ZipArchive;

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
    format_code: String,
}

/// The `Stylesheet` provides a mapping of styles properties such as fonts, colors, themes, etc
#[derive(Default)]
pub(crate) struct Stylesheet {
    number_formats: Option<HashMap<Arc<NumberFormat>, Key>>,
    number_formats_reverse: Option<HashMap<Key, Arc<NumberFormat>>>,
    fonts: HashMap<Arc<FontProperty>, Key>,
    fonts_reverse: HashMap<Key, Arc<FontProperty>>,
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
                    self.number_formats = Some(HashMap::new());
                    self.number_formats_reverse = Some(HashMap::new());
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                    let mut fonts_buf = Vec::with_capacity(1024);
                    let mut font = FontProperty::default();
                    loop {
                        fonts_buf.clear();
                        match xml.read_event_into(&mut fonts_buf) {
                            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"sz" => {
                                for attr in e.attributes() {
                                    if let Ok(a) = attr {
                                        match a.key {
                                            QName(b"val") => {
                                                font.size = a.unescape_value()?.to_string()
                                            }
                                            _ => (),
                                        }
                                    }
                                }
                            }
                            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"b" => {
                                font.bold = true
                            }
                            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"i" => {
                                font.italic = true
                            }
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
                                                let val = a.unescape_value()?.to_string();
                                                // The first two letter are ignored since they response to alpha
                                                let base16 = 16u32;
                                                let red = u8::from_str_radix(&val[2..4], base16)?;
                                                let green = u8::from_str_radix(&val[4..6], base16)?;
                                                let blue = u8::from_str_radix(&val[6..8], base16)?;
                                                font.color =
                                                    Color::Rgb(Rgb::Custom((red, green, blue)));
                                            }
                                            QName(b"theme") => {
                                                font.color = Color::Theme {
                                                    id: a
                                                        .unescape_value()?
                                                        .to_string()
                                                        .parse::<u32>()?,
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
                                            QName(b"val") => {
                                                font.font = a.unescape_value()?.to_string()
                                            }
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
                                            QName(b"val") => {
                                                font.scheme = a.unescape_value()?.to_string()
                                            }
                                            _ => (),
                                        }
                                    }
                                }
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"font" => break,
                            Ok(Event::Eof) => return Err(XcelmateError::XmlEof("font".into())),
                            Err(e) => return Err(XcelmateError::Xml(e)),
                            _ => (),
                        }
                    }
                    // Skip the first default font since we dont care about it and keep correct index mapping
                    if skipped_first_font {
                        let key = self.fonts.len() + 1;
                        let font = Arc::new(font);
                        self.fonts_reverse.insert(key, font.clone());
                        self.fonts.insert(font, key);
                    } else {
                        skipped_first_font = true
                    }
                }
                Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"numFmt" => {
                    if let (Some(n), Some(nr)) = (
                        self.number_formats.as_mut(),
                        self.number_formats_reverse.as_mut(),
                    ) {
                        let mut numfmt = NumberFormat::default();
                        for attr in e.attributes() {
                            if let Ok(a) = attr {
                                match a.key {
                                    QName(b"numFmtId") => {
                                        numfmt.id =
                                            a.unescape_value()?.to_string().parse::<u32>()?
                                    }
                                    QName(b"formatCode") => {
                                        numfmt.format_code = a.unescape_value()?.to_string()
                                    }
                                    _ => (),
                                }
                            }
                        }
                        // Update tables for biderctional use
                        let key = n.len(); // we should offset since
                        let numfmt = Arc::new(numfmt);
                        nr.insert(key, numfmt.clone());
                        n.insert(numfmt, key);
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"styleSheet" => break,
                Ok(Event::Eof) => return Err(XcelmateError::XmlEof("styleSheet".into())),
                Err(e) => return Err(XcelmateError::Xml(e)),
                _ => (),
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod stylesheet_edges {
    use crate::stream::xlsx::stylesheet::{Color, FontProperty};

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

    #[test]
    fn first_default_font_should_be_skipped() {
        let mut style = init("tests/workbook03.xlsx");
        let font = style.fonts_reverse.get(&2).unwrap();
        assert_eq!(
            **font,
            FontProperty {
                size: "11".into(),
                color: Color::Theme{id: 1, tint: None},
                font: "Calibri".into(),
                family: 2,
                scheme: "minor".into(),
                ..Default::default()
            }
        )
    }
}
