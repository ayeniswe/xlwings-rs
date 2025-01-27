use proc_macro::TokenStream;
use quote::{quote, ToTokens as _};
use syn::{
    parse_macro_input, spanned::Spanned as _, Data, DeriveInput, Error, Fields, LitBool,
    LitByteStr, LitStr,
};

pub fn impl_xml_writer(input: TokenStream) -> TokenStream {
    // Gather the code definition
    let input = parse_macro_input!(input as DeriveInput);
    let name = &input.ident;
    let mut name_str = None;

    // Gather top level struct metadata
    for attr in input.attrs {
        let result = if attr.path().is_ident("xml") {
            attr.parse_nested_meta(|meta| {
                if meta.path.is_ident("name") {
                    name_str = Some(meta.value()?.parse::<LitStr>()?.value());
                } else {
                    return Err(meta.error(format!(
                        "Unsupported top-level `#[xml(...)]` option `{}`",
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
    let fields = if let Data::Struct(data_struct) = &input.data {
        match &data_struct.fields {
            Fields::Named(fields) => &fields.named,
            _ => panic!("Only struct with named fields is supported"),
        }
    } else {
        panic!("Only structs are supported")
    };

    // XML serialization code
    let mut attr_writers = Vec::new(); // tag attribute writers
    let mut element_writers = Vec::new(); // tag element writers

    let mut following_elements = false;
    for field in fields {
        // Get code struct field definition
        let field_name = &field.ident.clone().unwrap();
        let mut field_name_str = field_name.to_string();

        // Gather struct fields optional metadata
        let mut default_bool = None;
        let mut default_bytes = None;
        let mut element = false;
        let mut skip = false;
        for attr in &field.attrs {
            let result = if attr.path().is_ident("xml") {
                attr.parse_nested_meta(|meta| {
                    if meta.path.is_ident("default_bool") {
                        default_bool = Some(meta.value()?.parse::<LitBool>()?.value());
                    } else if meta.path.is_ident("default_bytes") {
                        default_bytes = Some(meta.value()?.parse::<LitByteStr>()?);
                    } else if meta.path.is_ident("name") {
                        field_name_str = meta.value()?.parse::<LitStr>()?.value();
                    } else if meta.path.is_ident("skip") {
                        skip = true;
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

        // Generate the logic for writing the field to XML attributes
        if !element && !following_elements {
            let attr_write_logic = match &field.ty {
                syn::Type::Path(type_path) => {
                    let last_segment = type_path.path.segments.last().unwrap();
                    match last_segment.ident.to_string().as_str() {
                        "bool" => {
                            let logic = if let Some(default_bool) = default_bool {
                                quote! {
                                    if self.#field_name != #default_bool {
                                        let value = if self.#field_name { b"1" } else { b"0" };
                                        attrs.push((#field_name_str.as_bytes(), value.as_ref()));
                                    }
                                }
                            } else {
                                quote! {
                                    let value = if self.#field_name { b"1" } else { b"0" };
                                    attrs.push((#field_name_str.as_bytes(), value.as_ref()));
                                }
                            };
                            Ok(logic)
                        }
                        "Vec" => {
                            // Handle Vec<u8> fields only for attributes
                            let inner_type = match &type_path.path.segments[0].arguments {
                                syn::PathArguments::AngleBracketed(args) => {
                                    if let syn::GenericArgument::Type(inner_type) = &args.args[0] {
                                        if inner_type.to_token_stream().to_string() == "u8" {
                                            Ok(())
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
                            };
                            match inner_type {
                                Ok(_) => {
                                    let logic = if let Some(default_bytes) = default_bytes {
                                        quote! {
                                            if self.#field_name != #default_bytes {
                                                attrs.push((#field_name_str.as_bytes(), self.#field_name.as_ref()));;
                                            }
                                        }
                                    } else {
                                        quote! {
                                            if !self.#field_name.is_empty() {
                                                attrs.push((#field_name_str.as_bytes(), self.#field_name.as_ref()));;
                                            }
                                        }
                                    };
                                    Ok(logic)
                                }
                                Err(e) => Err(e),
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
            match attr_write_logic {
                Ok(logic) => attr_writers.push(logic),
                Err(e) => panic!("Failed: {}", e),
            }
        } else {
            let element_write_logic = match &field.ty {
                syn::Type::Path(type_path) => {
                    let last_segment = type_path.path.segments.last().unwrap();
                    match last_segment.ident.to_string().as_str() {
                        "Option" => {
                            let logic = quote! {
                                if let Some(value) = &self.#field_name {
                                    value.write_xml(writer, #field_name_str)?;
                                }
                            };
                            Ok(logic)
                        }
                        "Vec" => {
                            let logic = quote! {
                                for item in &self.#field_name {
                                    item.write_xml(writer, #field_name_str)?;
                                }
                            };
                            Ok(logic)
                        }
                        _ => {
                            let logic = quote! {
                                self.#field_name.write_xml(writer, #field_name_str)?;
                            };
                            Ok(logic)
                        }
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
            match element_write_logic {
                Ok(logic) => element_writers.push(logic),
                Err(e) => panic!("Failed: {}", e),
            }
        }
    }

    // Generate the trait
    let tag_name = if let Some(name_str) = name_str {
        quote! { #name_str }
    } else {
        quote! { tag_name }
    };
    let expanded = if element_writers.is_empty() {
        quote! {
            impl<W: Write> XmlWriter<W> for #name {
                fn write_xml<'a>(
                    &self,
                    writer: &'a mut Writer<W>,
                    tag_name: &'a str,
                ) -> Result<&'a mut Writer<W>, XlsxError> {
                    let mut attrs = Vec::new();
                    // Generated attribute writing logic
                    #(#attr_writers)*

                    writer
                        .create_element(#tag_name)
                        .with_attributes(attrs)
                        .write_empty()?;

                    Ok(writer)
                }
            }
        }
    } else {
        quote! {
                impl<W: Write> XmlWriter<W> for #name {
                    fn write_xml<'a>(
                        &self,
                        writer: &'a mut Writer<W>,
                        tag_name: &'a str,
                    ) -> Result<&'a mut Writer<W>, XlsxError> {
                        let mut attrs = Vec::new();
                        // Generated attribute writing logic
                        #(#attr_writers)*

                        writer
                        .create_element(#tag_name)
                        .with_attributes(attrs)
                        .write_inner_content::<_, XlsxError>(|writer| {
                            // Generated element writing logic
                            #(#element_writers)*
                            Ok(())
                        })?;

                        Ok(writer)
                    }
                }
        }
    };
    TokenStream::from(expanded)
}
