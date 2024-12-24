use std::io::{Read, Seek};

use crate::{errors::XcelmateError, stream::utils::xml_reader};
use quick_xml::events::Event;
use zip::ZipArchive;

/// The `Stylesheet` provides a mapping of styles
///

#[derive(Default)]
pub(crate) struct Stylesheet {

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
        let mut idx = 0; // Track index for referencing updates, etc
        loop {
            buf.clear();
            match xml.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"numFmts" => {
                    for attr in e.attributes() {
                        if let Ok(a) = attr {
                            match a.key {
                                // We dont care about unique count since that will be the len() of the table in SharedStringTable
                                _ => (),
                            }
                        }
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
mod shared_string_api {
    use std::fs::File;
    use zip::ZipArchive;

    use super::Stylesheet;

    fn init(path: &str) -> Stylesheet {
        let file = File::open(path).unwrap();
        let mut zip = ZipArchive::new(file).unwrap();
        let mut stylesheet = Stylesheet::default();
        stylesheet.read_stylesheet(&mut zip).unwrap();
        stylesheet
    }

    #[test]
    fn todo() {
        todo!()
    }
}
