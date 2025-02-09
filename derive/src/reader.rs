use proc_macro::{Span, TokenStream};
use quote::{quote, ToTokens as _};
use syn::{
    parse_macro_input, punctuated::Punctuated, spanned::Spanned as _, token::Comma, Data,
    DeriveInput, Error, Field, Fields, LitBool, LitByteStr, LitStr,
};

pub fn impl_xml_reader(input: TokenStream) -> TokenStream {
    // Parse the incoming token stream into a structured representation of the type (DeriveInput).
    let input = parse_macro_input!(input as DeriveInput);
    // Extract the identifier (name) of the type (struct or enum) that the macro is processing.
    let name = &input.ident;
    // Convert the identifier into a mutable string, allowing for later customization via attributes.
    let mut name_str = name.to_string();
    
    // Gather top-level metadata from the struct’s attributes.
    for attr in input.attrs {
        // Check and parse attributes based on their identifier.
        let result = if attr.path().is_ident("xml") {
            // If the attribute is #[xml(...)]
            attr.parse_nested_meta(|meta| {
                if meta.path.is_ident("name") {
                    name_str = meta.value()?.parse::<LitStr>()?.value();
                } else {
                    return Err(meta.error(format!(
                        "Unsupported top-level `#[xml(...)]` option `{}`",
                        meta.path.clone().into_token_stream()
                    )));
                }
                Ok(())
            })
        } else if attr.path().is_ident("doc") {
            // If the attribute is a documentation comment (#[doc]), ignore it.
            Ok(())
        } else {
            Err(Error::new(
                attr.span(),
                format!(
                    "Unsupported top-level attribute `{}` - expected `#[xml(...)]`",
                    attr.path().into_token_stream()
                ),
            ))
        };
        if let Err(e) = result {
            panic!("Failed to parse: {}", e);
        }
    }

    // Gather all struct fields
    let mut fields: &Punctuated<Field, Comma> = &Punctuated::new();
    if let Data::Struct(data_struct) = &input.data {
        match &data_struct.fields {
            Fields::Named(f) => fields = &f.named,
            _ => panic!("Only struct with named fields is supported"),
        }
    };
    // OR
    // Gather all enum variants
    let mut variants_fields = Vec::new();
    if let Data::Enum(data_enum) = &input.data {
        let mut fields = Vec::new();

        for variant in &data_enum.variants {
            match &variant.fields {
                Fields::Unnamed(u) => {
                    if u.unnamed.len() > 1 {
                        panic!("Only variants with a single unnamed field are supported")
                    } else {
                        fields.push((variant, u.unnamed.iter().last().unwrap()))
                    }
                }
                _ => panic!("Only enums with unnamed fields are supported"),
            }
        }

        variants_fields = fields
    }

    // XML serialization code
    let mut attributes = Vec::new(); // tag attributes
    let mut initial_item_attributes = Vec::new(); // tag attributes used in vec
    let mut elements = Vec::new(); // tag elements
    let mut initial_item_elements = Vec::new(); // tag elements used with a initial item
    let mut check_elements = Vec::new(); // code to verify data has been captured
    let mut init_check_elements = Vec::new(); // code to initialization for checking elements
    let mut inner = quote! {}; // the inner value

    // Gather information if enum variants were found
    // Only supports elements
    for variant in &variants_fields {
        let (variant, variant_field_type) = (variant.0, variant.1);
        // Get code enum variant definition
        let variant_name = &variant.ident;
        let mut variant_name_str = variant_name.to_string();

        for attr in &variant.attrs {
            let result = if attr.path().is_ident("xml") {
                attr.parse_nested_meta(|meta| {
                    if meta.path.is_ident("name") {
                        variant_name_str = meta.value()?.parse::<LitStr>()?.value();
                    }
                    Ok(())
                })
            } else if attr.path().is_ident("doc") {
                // Ignore `#[doc]` attributes (doc comments)
                Ok(())
            } else {
                Err(Error::new(
                    attr.span(),
                    format!(
                        "Unsupported attribute `{}` - expected `#[xml(...)]`",
                        attr.path().into_token_stream()
                    ),
                ))
            };

            if let Err(e) = result {
                panic!("Failed to parse: {}", e);
            }
        }

        elements.push(quote! {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == #variant_name_str.as_bytes() => {
                propagated_event.replace(Ok(event.unwrap().into_owned()));
                let mut choice = #variant_field_type::default();
                choice.read_xml(#variant_name_str, xml, #name_str, propagated_event)?;
                *self = #name::#variant_name(choice);
                chosen = true;
            }
        });
        initial_item_elements.push(quote! {
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == #variant_name_str.as_bytes() => {
                propagated_event.replace(Ok(event.unwrap().into_owned()));
                let mut choice = #variant_field_type::default();
                choice.read_xml(#variant_name_str, xml, #name_str, propagated_event)?;
                let choice = #name::#variant_name(choice);
                item = Some(choice);
                chosen = true;
            }
        });
    }
    // OR
    // Gather information if struct fields were found
    let mut following_elements = false;
    let mut sequence = false;
    let mut next_sequence = None;
    let mut fields = fields.iter().peekable();
    while let Some(field) = fields.next() {
        // Get peek so sequence elements are parsed correctly
        if sequence {
            if let Some(f) = fields.peek() {
                next_sequence = Some(f.ident.clone().unwrap().to_string());
            }
        }

        // Get code struct field definition
        let field_name = &field.ident.clone().unwrap();
        let mut field_name_str = field_name.to_string();

        // Gather struct fields optional metadata
        let mut element = false;
        let mut skip = false;
        let mut inner_value = false;
        for attr in &field.attrs {
            let result = if attr.path().is_ident("xml") {
                attr.parse_nested_meta(|meta| {
                    if meta.path.is_ident("default_bool") {
                        let _ = meta.value()?.parse::<LitBool>()?.value();
                    } else if meta.path.is_ident("default_bytes") {
                        let _ = meta.value()?.parse::<LitByteStr>()?;
                    } else if meta.path.is_ident("name") {
                        field_name_str = meta.value()?.parse::<LitStr>()?.value();
                    } else if meta.path.is_ident("sequence") {
                        sequence = true;
                        if let Some(f) = fields.peek() {
                            next_sequence = Some(f.ident.clone().unwrap().to_string());
                        }
                    } else if meta.path.is_ident("skip") {
                        skip = true;
                    } else if meta.path.is_ident("val") {
                        inner_value = true;
                    } else if meta.path.is_ident("following_elements") {
                        following_elements = true;
                    } else if meta.path.is_ident("element") {
                        element = true;
                    } else {
                        return Err(meta.error(format!(
                            "Unsupported `#[xml(...)]` option `{}`",
                            meta.path.clone().into_token_stream()
                        )));
                    }
                    Ok(())
                })
            } else if attr.path().is_ident("doc") {
                // Ignore `#[doc]` attributes (doc comments)
                Ok(())
            } else {
                Err(Error::new(
                    attr.span(),
                    format!(
                        "Unsupported attribute `{}` - expected `#[xml(...)]`",
                        attr.path().into_token_stream()
                    ),
                ))
            };

            if let Err(e) = result {
                panic!("Failed to parse: {}", e);
            }
        }

        // Ignore fields marked with `#[xml(skip)]`
        if skip {
            continue;
        }

        // Generate the logic for reading the field to XML attributes
        if inner_value {
            let result = match &field.ty {
                syn::Type::Path(type_path) => {
                    match &type_path.path.segments[0].arguments {
                        syn::PathArguments::AngleBracketed(inner) => {
                             match &inner.args[0] {
                                syn::GenericArgument::Type(inner_type) => {
                                    if inner_type.to_token_stream().to_string() == "u8" {
                                        Ok(())
                                    } else {
                                        Err(Error::new(
                                            inner_type.span(),
                                            "Only Vec<u8> is supported for inner value. Specify `#[xml(element)]` if you want to serialize it as an element",
                                        ))
                                    }
                                }
                                args => {
                                    Err(Error::new(
                                        args.span(),
                                        format!(
                                            "Unsupported angle bracket args `{}` for inner value",
                                            generic.into_token_stream()
                                        ),
                                    ))
                                }
                             }
                        }
                        arg => {
                            Err(Error::new(
                                arg.span(),
                                format!(
                                    "Unsupported type path args `{}` for inner value",
                                    arg.into_token_stream()
                                ),
                            )),
                        }
                    }
                }
                ty => Err(Error::new(
                    ty.span(),
                    format!(
                        "Unsupported field type path `{}`",
                        ty.into_token_stream()
                    ),
                )),
            }
            
            if Err(e) = result {
                panic!("Failed to parse: {}", e);
            }
            else {
                elements.push(quote! {
                    Ok(Event::Text(ref e)) => {
                        self.#field_name = e.as_ref().into();
                        #field_name = true;
                        break;
                    }
                });
                initial_item_elements.push(quote! {
                    Ok(Event::Text(ref e)) => {
                        item.#field_name = e.as_ref().into();
                        #field_name = true;
                        break;
                    }
                });
                check_elements.push(quote! {
                    if !#field_name {
                        panic!("Missing required inner text `{}`", #field_name_str);
                    }
                });
                init_check_elements.push(quote! {
                    let mut #field_name = false;
                });
            }
        } else if !element && !following_elements {
            let attr_read_logic = match &field.ty {
                syn::Type::Path(type_path) => {
                    let last_segment = type_path.path.segments.last().unwrap();
                    let field_name_as_bytes =
                        LitByteStr::new(field_name_str.as_bytes(), Span::call_site().into());

                    match last_segment.ident.to_string().as_str() {
                        "bool" => Ok((
                            quote! {
                                #field_name_as_bytes => self.#field_name = *a.value == *b"1" || *a.value == *b"true" || *a.value == *b"on",
                            },
                            quote! {
                                #field_name_as_bytes => item.#field_name = *a.value == *b"1" || *a.value == *b"true" || *a.value == *b"on",
                            },
                        )),
                        "Vec" => {
                            // Handle Vec<u8> fields only for attributes
                            match &type_path.path.segments[0].arguments {
                                syn::PathArguments::AngleBracketed(args) => {
                                    if let syn::GenericArgument::Type(inner_type) = &args.args[0] {
                                        if inner_type.to_token_stream().to_string() == "u8" {
                                            Ok((
                                                quote! {
                                                    #field_name_as_bytes => self.#field_name = a.value.into(),
                                                },
                                                quote! {
                                                    #field_name_as_bytes => item.#field_name = a.value.into(),
                                                },
                                            ))
                                        } else {
                                            Err(Error::new(
                                                inner_type.span(),
                                                "Only Vec<u8> is supported for attribute. Specify `#[xml(element)]` if you want to serialize it as an element",
                                            ))
                                        }
                                    } else {
                                        let generic = &args.args[0];
                                        Err(Error::new(
                                            generic.span(),
                                            format!(
                                                "Unsupported Vec inner type `{}` for attribute",
                                                generic.into_token_stream()
                                            ),
                                        ))
                                    }
                                }
                                arg => Err(Error::new(
                                    arg.span(),
                                    format!(
                                        "Unsupported Vec type `{}` for attribute",
                                        arg.into_token_stream()
                                    ),
                                )),
                            }
                        }

                        segement => Err(Error::new(
                            segement.span(),
                            format!("Unsupported struct field datatype `{}`", segement),
                        )),
                    }
                }

                r#type => Err(Error::new(
                    r#type.span(),
                    format!(
                        "Unsupported struct field type `{}`",
                        r#type.into_token_stream()
                    ),
                )),
            };

            match attr_read_logic {
                Ok(logic) => {
                    attributes.push(logic.0);
                    initial_item_attributes.push(logic.1);
                }
                Err(e) => panic!("Failed: {}", e),
            }
        } else {
            let element_read_logic = match &field.ty {
                syn::Type::Path(type_path) => {
                    let last_segment = type_path.path.segments.last().unwrap();
                    let field_type = last_segment.ident.clone();

                    match field_type.to_string().as_str() {
                        "Option" => Ok((
                            quote! {
                                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == #field_name_str.as_bytes() => {
                                    propagated_event.replace(Ok(event.unwrap().into_owned()));
                                    self.#field_name.read_xml(#field_name_str, xml, #name_str, propagated_event)?;
                                }
                            },
                            quote! {
                                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == #field_name_str.as_bytes() => {
                                    propagated_event.replace(Ok(event.unwrap().into_owned()));
                                    item.#field_name.read_xml(#field_name_str, xml, #name_str, propagated_event)?;
                                }
                            },
                            quote! {},
                            quote! {},
                        )),
                        "Vec" => {
                            // Sequence of different elements multiples can appear so we need to use the next element differentiate tag as closing
                            let closing_tag = if let Some(next_tag) = next_sequence.take() {
                                quote! {
                                    #next_tag
                                }
                            } else {
                                quote! {
                                    tag_name
                                }
                            };

                            Ok((
                                quote! {
                                    Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == #field_name_str.as_bytes() => {
                                        propagated_event.replace(Ok(event.unwrap().into_owned()));
                                        self.#field_name.read_xml(#field_name_str, xml, #closing_tag, propagated_event)?;
                                        #field_name = true;
                                    }
                                },
                                quote! {
                                    Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == #field_name_str.as_bytes() => {
                                        propagated_event.replace(Ok(event.unwrap().into_owned()));
                                        item.#field_name.read_xml(#field_name_str, xml, #closing_tag, propagated_event)?;
                                        #field_name = true;
                                    }
                                },
                                // Validating the presence of the field
                                quote! {
                                    if !#field_name {
                                        panic!("Missing required field `{}`", #field_name_str);
                                    }
                                },
                                // Intializing validation
                                quote! {
                                    let mut #field_name = false;
                                },
                            ))
                        }
                        _ => Ok((
                            quote! {
                                // no need to worry about closing tags
                                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == #field_name_str.as_bytes() => {
                                    propagated_event.replace(Ok(event.unwrap().into_owned()));
                                    self.#field_name.read_xml(#field_name_str, xml, #name_str, propagated_event)?;
                                    #field_name = true;
                                }
                            },
                            quote! {
                                // no need to worry about closing tags
                                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == #field_name_str.as_bytes() => {
                                    propagated_event.replace(Ok(event.unwrap().into_owned()));
                                    item.#field_name.read_xml(#field_name_str, xml, #name_str, propagated_event)?;
                                    #field_name = true;
                                }
                            },
                            // Validating the presence of the field
                            quote! {
                                if !#field_name {
                                    panic!("Missing required field `{}`", #field_name_str);
                                }
                            },
                            // Intializing validation
                            quote! {
                                let mut #field_name = false;
                            },
                        )),
                    }
                }

                r#type => Err(Error::new(
                    r#type.span(),
                    format!(
                        "Unsupported struct field type `{}`",
                        r#type.into_token_stream()
                    ),
                )),
            };

            match element_read_logic {
                Ok(logic) => {
                    elements.push(logic.0);
                    initial_item_elements.push(logic.1);
                    check_elements.push(logic.2);
                    init_check_elements.push(logic.3);
                }
                Err(e) => panic!("Failed: {}", e),
            }
        }
    }

    // An element needs to be init to use in Vec and Option situations
    let mut init_element = quote! {};
    // For Vec need to safely unwrap since checks are already done to gurantee
    let mut add_vec_element = quote! {};
    // For option some are already cast to a option
    let mut set_opt_element = quote! {};

    // For enum variants
    if !variants_fields.is_empty() {
        // Validating the presence of the field
        check_elements.push(quote! {
            if !chosen {
                panic!("Missing required field `{}`", tag_name);
            }
        });
        // Intializing validation
        init_check_elements.push(quote! {
            let mut chosen = false;
        });
        init_element = quote! {let mut item = None;};
        add_vec_element = quote! {self.push(item.unwrap());};
        set_opt_element = quote! { *self = item;};
    }
    // For structs
    else {
        init_element = quote! {let mut item = #name::default();};
        add_vec_element = quote! {self.push(item);};
        set_opt_element = quote! { self.replace(item);};
    }

    let expanded =
        // Generate the implementation for the `XmlReader` trait for the struct
        quote! {
            impl<B: BufRead> XmlReader<B> for Vec<#name> {
                fn read_xml<'a>(&mut self, tag_name: &'a str, xml: &'a mut Reader<B>, closing_name: &'a str, propagated_event: &'a mut Option<Result<Event<'static>, quick_xml::Error>>)
                -> Result<(), XlsxError> {
                    // Keep memory usage to a minimum
                    let mut buf = Vec::with_capacity(1024);
                    loop {
                        #init_element
                        buf.clear();
                        let event = if let Some(e) = propagated_event.take() {
                            e
                        } else {
                            xml.read_event_into(&mut buf)
                        };
                        match event {
                            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == tag_name.as_bytes() => {
                                // Read the tag attributes
                                for attr in e.attributes() {
                                    if let Ok(a) = attr {
                                        match a.key.as_ref() {
                                            #(#initial_item_attributes)*
                                            _ => (),
                                        }
                                    }
                                }
                                // Read the nested tag contents
                                if let Ok(Event::Start(_)) = event {
                                    let mut nested_buf = Vec::with_capacity(1024);
                                    #(#init_check_elements)*
                                    loop {
                                        nested_buf.clear();
                                        let event = if let Some(e) = propagated_event.take() {
                                            e
                                        } else {
                                            xml.read_event_into(&mut nested_buf)
                                        };
                                        match event {
                                            #(#initial_item_elements)*
                                            Ok(Event::End(ref e)) if e.local_name().as_ref() == tag_name.as_bytes() => {
                                                break
                                            }
                                            Ok(Event::Eof) => {
                                                return Err(XlsxError::XmlEof(tag_name.into()))
                                            }
                                            Err(e) => {
                                                return Err(XlsxError::Xml(e));
                                            }
                                            _ => (),
                                        }
                                    }
                                    #(#check_elements)*
                                }
                                #add_vec_element
                            }
                            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == closing_name.as_bytes() => {
                                propagated_event.replace(Ok(event.unwrap().into_owned()));
                                break
                            },
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == closing_name.as_bytes() => {
                                propagated_event.replace(Ok(event.unwrap().into_owned()));
                                break
                            },
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof(tag_name.into())),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => ()
                        }
                    }
                    Ok(())
                }
            }
            impl<B: BufRead> XmlReader<B> for #name {
                fn read_xml<'a>(
                    &mut self,
                    tag_name: &'a str,
                    xml: &'a mut Reader<B>,
                    closing_name: &'a str,
                    propagated_event: &'a mut Option<Result<Event<'static>, quick_xml::Error>>
                ) -> Result<(), XlsxError> {
                    // Keep memory usage to a minimum
                    let mut buf = Vec::with_capacity(1024);
                    loop {
                        buf.clear();
                        let event = if let Some(e) = propagated_event.take() {
                            e
                        } else {
                            xml.read_event_into(&mut buf)
                        };
                        match event {
                            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == tag_name.as_bytes() => {
                                // Read the tag attributes
                                for attr in e.attributes() {
                                    if let Ok(a) = attr {
                                        match a.key.as_ref() {
                                            #(#attributes)*
                                            _ => (),
                                        }
                                    }
                                }
                                // Read the nested tag contents
                                if let Ok(Event::Start(_)) = event {
                                    let mut nested_buf = Vec::with_capacity(1024);
                                    #(#init_check_elements)*
                                    loop {
                                        nested_buf.clear();
                                        let event = if let Some(e) = propagated_event.take() {
                                            e
                                        } else {
                                            xml.read_event_into(&mut nested_buf)
                                        };
                                        match event {
                                            #(#elements)*
                                            Ok(Event::End(ref e)) if e.local_name().as_ref() == tag_name.as_bytes() => {
                                                break
                                            }
                                            Ok(Event::Eof) => {
                                                return Err(XlsxError::XmlEof(tag_name.into()))
                                            }
                                            Err(e) => {
                                                return Err(XlsxError::Xml(e));
                                            }
                                            _ => (),
                                        }
                                    }
                                    #(#check_elements)*
                                }
                                break
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == tag_name.as_bytes() => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof(tag_name.into())),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    Ok(())
                }
            }
            impl<B: BufRead> XmlReader<B> for Option<#name> {
                fn read_xml<'a>(
                    &mut self,
                    tag_name: &'a str,
                    xml: &'a mut Reader<B>,
                    closing_name: &'a str,
                    propagated_event: &'a mut Option<Result<Event<'static>, quick_xml::Error>>
                ) -> Result<(), XlsxError> {
                    #init_element
                    // Keep memory usage to a minimum
                    let mut buf = Vec::with_capacity(1024);
                    loop {
                        buf.clear();
                        let event = if let Some(e) = propagated_event.take() {
                            e
                        } else {
                            xml.read_event_into(&mut buf)
                        };
                        match event {
                            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if e.local_name().as_ref() == tag_name.as_bytes() => {
                                // Read the tag attributes
                                for attr in e.attributes() {
                                    if let Ok(a) = attr {
                                        match a.key.as_ref() {
                                            #(#initial_item_attributes)*
                                            _ => (),
                                        }
                                    }
                                }

                                // Read the nested tag contents
                                if let Ok(Event::Start(_)) = event {
                                    let mut nested_buf = Vec::with_capacity(1024);
                                    #(#init_check_elements)*
                                    loop {
                                        nested_buf.clear();
                                        let event = if let Some(e) = propagated_event.take() {
                                            e
                                        } else {
                                            xml.read_event_into(&mut nested_buf)
                                        };
                                        match event {
                                            #(#initial_item_elements)*
                                            Ok(Event::End(ref e)) if e.local_name().as_ref() == tag_name.as_bytes() => {
                                                break
                                            }
                                            Ok(Event::Eof) => {
                                                return Err(XlsxError::XmlEof(tag_name.into()))
                                            }
                                            Err(e) => {
                                                return Err(XlsxError::Xml(e));
                                            }
                                            _ => (),
                                        }
                                    }
                                    #(#check_elements)*
                                }
                                #set_opt_element
                                break
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == tag_name.as_bytes() => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof(tag_name.into())),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    Ok(())
                }
            }
        };
    TokenStream::from(expanded)
}
