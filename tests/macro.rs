mod xml_writer_derive {
    use derive::XmlWrite;
    use quick_xml::Writer;
    use std::io::Cursor;
    use std::io::Write;
    use xcelmate::stream::xlsx::errors::XlsxError;
    use xcelmate::stream::utils::XmlWriter;

    #[derive(XmlWrite)]
    #[xml(name = "ex")]
    struct Example {
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
