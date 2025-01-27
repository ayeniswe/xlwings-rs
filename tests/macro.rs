mod xml_writer_derive {
    use derive::XmlWrite;
    use quick_xml::Writer;
    use std::io::Cursor;
    use std::io::Write;
    use xcelmate::stream::utils::XmlWriter;
    use xcelmate::stream::xlsx::errors::XlsxError;

    #[derive(XmlWrite)]
    #[xml(name = "ex")]
    struct Example {
        #[xml(skip)]
        help: String,
        #[xml(name = "activePane")]
        active_pane: bool,
        #[xml(default_bool = true)]
        x_split: bool,
        missing: Vec<u8>,
        value_test: Vec<u8>,
        #[xml(default_bytes = b"test")]
        open_win: Vec<u8>,
        #[xml(element, name = "view")]
        sub_field: SubField,
        #[xml(element, name = "SubField")]
        subfield2: Vec<SubField>,
        #[xml(element, name = "SubField")]
        subfield3: Option<SubField>,
        #[xml(element, name = "SubField")]
        subfield4: Option<SubField>,
    }
    #[derive(XmlWrite)]
    struct SubField {
        #[xml(name = "mainValue")]
        value: bool,
    }

    #[test]
    fn test_xml_write_derive() {
        let sheet = Example {
            help: "DO NOT SHOW".into(),
            active_pane: false,
            x_split: true,
            value_test: b"01234".to_vec(),
            open_win: b"test".to_vec(),
            missing: Vec::new(),
            sub_field: SubField { value: true },
            subfield2: vec![SubField { value: true }, SubField { value: false }],
            subfield3: Some(SubField { value: false }),
            subfield4: None,
        };

        let mut buffer = Cursor::new(Vec::new());
        let mut writer = Writer::new(&mut buffer);
        let _ = sheet.write_xml(&mut writer, "sheet");

        let xml_output = String::from_utf8(buffer.into_inner()).unwrap();
        let expected_output = r#"<ex activePane="0" value_test="01234"><view mainValue="1"/><SubField mainValue="1"/><SubField mainValue="0"/><SubField mainValue="0"/></ex>"#;
        assert_eq!(xml_output, expected_output);
    }
}

mod xml_reader_derive {
    use derive::XmlRead;
    use quick_xml::{events::Event, Reader};
    use std::io::BufRead;
    use std::io::Cursor;
    // use xcelmate::stream::utils::XmlReader;
    use xcelmate::stream::xlsx::errors::XlsxError;

    #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
    #[xml(name = "ex")]
    struct Example {
        x_split: bool,
        // #[xml(skip)]
        // help: String,
        // #[xml(name = "activePane")]
        // active_pane: bool,
        // missing: Vec<u8>,
        // value_test: Vec<u8>,
        // #[xml(default_bytes = b"test")]
        // open_win: Vec<u8>,
        #[xml(element, name = "view")]
        sub_field: SubField,
        // #[xml(element, name = "SubField")]
        // subfield2: Vec<SubField>,
        // #[xml(element, name = "SubField")]
        // subfield3: Option<SubField>,
        // #[xml(element, name = "SubField")]
        // subfield4: Option<SubField>,
    }
    #[derive(Default, PartialEq, Eq, Debug)]
    struct SubField {
        // #[xml(name = "mainValue")]
        value: bool,
    }
    pub trait XmlReader<B: BufRead> {
        /// Allows us to read xml into a custom object
        fn read_xml<'a>(&mut self, tag_name: &'a str, xml: &'a mut Reader<B>, closing_name: &'a str)
            -> Result<(), XlsxError>;
    }
    impl<B: BufRead> XmlReader<B> for Vec<SubField> {
        fn read_xml<'a>(&mut self, tag_name: &'a str, xml: &'a mut Reader<B>, closing_name: &'a str)
            -> Result<(), XlsxError> {
            // Keep memory usage to a minimum
            let mut buf = Vec::with_capacity(1024);
            loop {
                let mut item = SubField::default();
                buf.clear();
                let event = xml.read_event_into(&mut buf);
                match event {
                    Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == tag_name.as_bytes() => {
                        // Read the tag attributes
                        for attr in e.attributes() {
                            if let Ok(a) = attr {
                                match a.key.as_ref() {
                                     b"mainValue" => item.value = *a.value == *b"1",
                                    _ => (),
                                }
                            }
                        }
                        // Read the nested tag contents
                        if let Ok(Event::Start(_)) = event {
                            //
                        }
                        self.push(item);
                    }
                    Ok(Event::End(ref e)) if e.local_name().as_ref() == closing_name.as_bytes() => break,
                    Ok(Event::Eof) => return Err(XlsxError::XmlEof(tag_name.into())),
                    Err(e) => return Err(XlsxError::Xml(e)),
                    _ => (),
                }
            }
            Ok(())
        }
    }
    #[test]
    fn test_xml_reader_derive() {
        let xml_content = r#"
        <ex x_split="1">
            <view mainValue="1" />
        </ex>
        "#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml).unwrap();
        assert_eq!(
            example,
            Example {
                x_split: true,
                sub_field: SubField { value: true }
            }
        );
    }
}
