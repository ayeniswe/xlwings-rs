mod xml_writer_derive {
    use crate::stream::{utils::XmlWriter, xlsx::errors::XlsxError};
    use derive::XmlWrite;
    use quick_xml::Writer;
    use std::io::Cursor;
    use std::io::Write;

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
    use crate::stream::{utils::XmlReader, xlsx::errors::XlsxError};
    use derive::XmlRead;
    use quick_xml::{events::Event, Reader};
    use std::io::{BufRead, Cursor};

    #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
    struct Example {
        active_pane: bool,
        #[xml(val)]
        inner: Vec<u8>,
    }

    #[test]
    fn test_xml_reader_inner_text() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            active_pane: bool,
            #[xml(val)]
            inner: Vec<u8>,
        }
        let xml_content = r#"
        <Example active_pane="1">Hello World</Example>"#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example
            .read_xml("Example", &mut xml, "Example", &mut None)
            .unwrap();
        assert_eq!(
            example,
            Example {
                active_pane: true,
                inner: b"Hello World".to_vec()
            }
        );
    }
    #[test]
    fn test_xml_reader_empty_tag_attributes() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            active_pane: bool,
            window: Vec<u8>,
        }
        let xml_content = r#"
        <Example active_pane="1" window="hello" />"#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example
            .read_xml("Example", &mut xml, "Example", &mut None)
            .unwrap();
        assert_eq!(
            example,
            Example {
                active_pane: true,
                window: b"hello".to_vec()
            }
        );
    }
    #[test]
    fn test_xml_reader_start_tag_attributes() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            active_pane: bool,
            window: Vec<u8>,
        }
        let xml_content = r#"
        <Example active_pane="1" window="hello" ></Example>"#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example
            .read_xml("Example", &mut xml, "Example", &mut None)
            .unwrap();
        assert_eq!(
            example,
            Example {
                active_pane: true,
                window: b"hello".to_vec()
            }
        );
    }
    #[test]
    fn test_xml_reader_top_level_tag_name_alter_has_no_effect() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        #[xml(name = "e")]
        struct Example {
            active_pane: bool,
        }
        let xml_content = r#"
        <ex active_pane="1" ></ex>"#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
        assert_eq!(example, Example { active_pane: true });
    }
    #[test]
    fn test_xml_reader_element_tag_name_alter() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            #[xml(name = "active")]
            active_pane: bool,
        }
        let xml_content = r#"
        <ex active="1" ></ex>"#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
        assert_eq!(example, Example { active_pane: true });
    }
    #[test]
    fn test_xml_reader_element_skip() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            #[xml(skip)]
            active_pane: bool,
        }
        let xml_content = r#"
        <ex active_pane="1"></ex>"#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
        assert_eq!(example, Example { active_pane: false });
    }
    #[test]
    fn test_xml_reader_element_read() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            active_pane: bool,
            #[xml(element)]
            side: SideExample,
        }
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct SideExample {
            active_pane: bool,
            window: Vec<u8>,
        }
        let xml_content = r#"
        <ex active_pane="1">
            <side window="hello" active_pane="true" />
        </ex>
        "#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
        assert_eq!(
            example,
            Example {
                active_pane: true,
                side: SideExample {
                    active_pane: true,
                    window: b"hello".to_vec()
                }
            }
        );
    }
    #[test]
    fn test_xml_reader_element_as_enum() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Holder {
            active_pane: bool,
            #[xml(element)]
            value: SomeExample,
        }

        #[derive(XmlRead, PartialEq, Eq, Debug)]
        enum SomeExample {
            #[xml(name = "side")]
            Side(SideExample),
            Regular(Example),
        }
        impl Default for SomeExample {
            fn default() -> Self {
                SomeExample::Regular(Example::default())
            }
        }

        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            active_pane: bool,
        }
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct SideExample {
            window: Vec<u8>,
        }

        let xml_content = r#"
        <ex active_pane="1">
            <value>
                <side window="hello"/>
            </value>
        </ex>
        "#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Holder {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
        assert_eq!(
            example,
            Holder {
                active_pane: true,
                value: SomeExample::Side(SideExample {
                    window: b"hello".to_vec()
                })
            }
        );
    }
    #[test]
    fn test_xml_reader_element_as_optional() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            active_pane: bool,
            #[xml(element)]
            side: Option<SideExample>,
        }
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct SideExample {
            active_pane: bool,
            window: Vec<u8>,
        }

        let xml_content = r#"
        <ex active_pane="1">
            <side window="hello" active_pane="true" />
        </ex>
        "#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
        assert_eq!(
            example,
            Example {
                active_pane: true,
                side: Some(SideExample {
                    active_pane: true,
                    window: b"hello".to_vec()
                })
            }
        );

        // If not existing, it should be None
        let xml_content = r#"
        <ex active_pane="1"></ex>
        "#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
        assert_eq!(
            example,
            Example {
                active_pane: true,
                side: None,
            }
        );
    }
    #[test]
    #[should_panic(expected = "Missing required field `side`")]
    fn test_xml_reader_element_as_required() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            active_pane: bool,
            #[xml(element)]
            side: SideExample,
        }
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct SideExample {
            active_pane: bool,
            window: Vec<u8>,
        }

        let xml_content = r#"
        <ex active_pane="1"></ex>
        "#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
    }
    #[test]
    fn test_xml_reader_element_as_array() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            active_pane: bool,
            #[xml(element)]
            side: Vec<SideExample>,
        }
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct SideExample {
            active_pane: bool,
            window: Vec<u8>,
        }

        let xml_content = r#"
        <ex active_pane="1">
            <side window="hello2" active_pane="true" />
            <side window="hello3" active_pane="true" />
            <side window="hello1" active_pane="true" />
        </ex>
        "#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
        assert_eq!(
            example,
            Example {
                active_pane: true,
                side: vec![
                    SideExample {
                        active_pane: true,
                        window: b"hello2".to_vec()
                    },
                    SideExample {
                        active_pane: true,
                        window: b"hello3".to_vec()
                    },
                    SideExample {
                        active_pane: true,
                        window: b"hello1".to_vec()
                    },
                ]
            }
        );
    }
    #[test]
    fn test_xml_reader_element_as_array_with_sequence() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            active_pane: bool,
            #[xml(following_elements, sequence)]
            side: Vec<SideExample>,
            side2: Vec<SideExample>,
        }
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct SideExample {
            active_pane: bool,
            window: Vec<u8>,
        }

        let xml_content = r#"
        <ex active_pane="1">
            <side window="hello2" active_pane="true"/>
            <side window="hello3" active_pane="true"/>
            <side window="hello1" active_pane="true"/>
            <side2 window="side2 hello12" active_pane="true"/>
            <side2 window="side2 hello123" active_pane="true"/>
            <side2 window="side2 hello1" active_pane="true"/>
        </ex>
        "#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
        assert_eq!(
            example,
            Example {
                active_pane: true,
                side: vec![
                    SideExample {
                        active_pane: true,
                        window: b"hello2".to_vec()
                    },
                    SideExample {
                        active_pane: true,
                        window: b"hello3".to_vec()
                    },
                    SideExample {
                        active_pane: true,
                        window: b"hello1".to_vec()
                    },
                ],
                side2: vec![
                    SideExample {
                        active_pane: true,
                        window: b"side2 hello12".to_vec()
                    },
                    SideExample {
                        active_pane: true,
                        window: b"side2 hello123".to_vec()
                    },
                    SideExample {
                        active_pane: true,
                        window: b"side2 hello1".to_vec()
                    },
                ]
            }
        );
    }
    #[test]
    fn test_xml_reader_element_read_alter_element_tag_name() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            #[xml(element, name = "noside")]
            side: SideExample,
        }
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct SideExample {
            active_pane: bool,
            window: Vec<u8>,
        }
        let xml_content = r#"
        <ex>
            <noside window="hello" active_pane="true" />
        </ex>
        "#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
        assert_eq!(
            example,
            Example {
                side: SideExample {
                    active_pane: true,
                    window: b"hello".to_vec()
                }
            }
        );
    }
    #[test]
    fn test_xml_reader_element_all_as_elements() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            #[xml(following_elements, name = "noside")]
            side: SideExample,
            #[xml(name = "littleside")]
            lside: SideExample,
            #[xml(name = "bigside")]
            bside: SideExample,
        }
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct SideExample {
            active_pane: bool,
            window: Vec<u8>,
        }
        let xml_content = r#"
        <ex>
            <noside window="hello" active_pane="1" />
            <littleside window="very tiny" active_pane="0" />
            <bigside window="very big" active_pane="1" />
        </ex>
        "#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            ..Default::default()
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
        assert_eq!(
            example,
            Example {
                side: SideExample {
                    active_pane: true,
                    window: b"hello".to_vec()
                },
                lside: SideExample {
                    active_pane: false,
                    window: b"very tiny".to_vec()
                },
                bside: SideExample {
                    active_pane: true,
                    window: b"very big".to_vec()
                }
            }
        );
    }
    #[test]
    fn test_xml_reader_element_all_bool_valid_types() {
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct Example {
            #[xml(following_elements, name = "noside")]
            side: SideExample,
            #[xml(name = "littleside")]
            lside: SideExample,
            #[xml(name = "bigside")]
            bside: SideExample,
            #[xml(name = "noside2")]
            side2: SideExample,
            #[xml(name = "littleside2")]
            lside2: SideExample,
            #[xml(name = "bigside2")]
            bside2: SideExample,
        }
        #[derive(XmlRead, Default, PartialEq, Eq, Debug)]
        struct SideExample {
            active_pane: bool,
        }
        let xml_content = r#"
        <ex>
            <noside active_pane="1" />
            <littleside active_pane="on" />
            <bigside active_pane="true" />
            <noside2 active_pane="0" />
            <littleside2 active_pane="off" />
            <bigside2 active_pane="false" />
        </ex>
        "#;
        let mut xml = Reader::from_reader(Cursor::new(xml_content));
        let mut example = Example {
            side: SideExample { active_pane: false },
            lside: SideExample { active_pane: false },
            bside: SideExample { active_pane: false },
            side2: SideExample { active_pane: true },
            lside2: SideExample { active_pane: true },
            bside2: SideExample { active_pane: true },
        };
        example.read_xml("ex", &mut xml, "ex", &mut None).unwrap();
        assert_eq!(
            example,
            Example {
                side: SideExample { active_pane: true },
                lside: SideExample { active_pane: true },
                bside: SideExample { active_pane: true },
                side2: SideExample { active_pane: false },
                lside2: SideExample { active_pane: false },
                bside2: SideExample { active_pane: false },
            }
        );
    }
}
