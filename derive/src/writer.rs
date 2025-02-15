use proc_macro::TokenStream;
use quote::{quote, ToTokens as _};
use syn::{
    parse_macro_input, punctuated::Punctuated, spanned::Spanned as _, token::Comma, Data,
    DeriveInput, Error, Field, Fields, LitBool, LitByteStr, LitStr,
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
    let mut fields = &Punctuated::new();
    if let Data::Struct(data_struct) = &input.data {
        match &data_struct.fields {
            Fields::Named(f) => fields = &f.named,
            _ => panic!("Only struct with named fields is supported"),
        }
    };
    // OR: If the input isn’t a struct, handle an enum instead.
    let mut variants_fields = Vec::new();
    if let Data::Enum(data_enum) = &input.data {
        let mut fields = Vec::new();
        for variant in &data_enum.variants {
            // Only process variants that are tuple-like (unnamed fields).
            match &variant.fields {
                Fields::Unnamed(u) => {
                    if u.unnamed.len() > 1 {
                        panic!("Only tuple-like variants with a single field are supported")
                    } else {
                        fields.push(variant)
                    }
                }
                _ => panic!("Only enums variants tuple-like are supported"),
            }
        }
        variants_fields = fields
    }

    // XML serialization code: prepare containers for various XML parsing components.
    let mut attr_writers = Vec::new(); // tag attribute writers
    let mut element_writers = Vec::new(); // tag element writers
    let mut inner_text = quote! {}; // tag element inner text

    // Gather information if enum variants were found
    // Only supports elements
    for variant in &variants_fields {
        // Retrieve the variant's identifier and convert it to a string, which will serve as the default XML tag name.
        let variant_name = &variant.ident;
        let mut variant_name_str = variant_name.to_string();
        // Process each attribute attached to the variant.
        for attr in &variant.attrs {
            let result = if attr.path().is_ident("xml") {
                // If the attribute is #[xml(...)], parse its nested metadata.
                attr.parse_nested_meta(|meta| {
                    // If a 'name' option is provided, override the default tag name with this value.
                    if meta.path.is_ident("name") {
                        variant_name_str = meta.value()?.parse::<LitStr>()?.value();
                    }
                    Ok(())
                })
            } else if attr.path().is_ident("doc") {
                // Ignore documentation attributes.
                Ok(())
            } else {
                // Any other attribute is not supported and results in an error.
                Err(Error::new(
                    attr.span(),
                    format!(
                        "Unsupported attribute `{}` - expected `#[xml(...)]`",
                        attr.path().into_token_stream()
                    ),
                ))
            };
            // If parsing fails, terminate with an error.
            if let Err(e) = result {
                panic!("Failed to parse: {}", e);
            }
        }
        // Generate the code fragment to handle XML inner writer types
        element_writers.push(quote! {
            #name::#variant_name(v) => {
                v.write_xml(writer, #variant_name_str)?
            }
        });
    }
    //
    // OR
    //
    // Optional metadata that can effect globally other fields
    let mut following_elements = false;
    let mut inner_value_found = false;
    for field in fields {
        // Get code struct field definition
        let field_name = &field.ident.clone().unwrap();
        let mut field_name_str = field_name.to_string();

        // Gather struct fields optional metadata
        let mut default_bool = None;
        let mut default_bytes = None;
        let mut element = false;
        let mut skip = false;
        let mut inner_value = false;
        for attr in &field.attrs {
            let result = if attr.path().is_ident("xml") {
                attr.parse_nested_meta(|meta| {
                    // Track if a value is found to equal default for boolean it will prevent write
                    if meta.path.is_ident("default_bool") {
                        default_bool = Some(meta.value()?.parse::<LitBool>()?.value());
                    // Track if a value is found to equal default for bytes it will prevent write.
                    } else if meta.path.is_ident("default_bytes") {
                        default_bytes = Some(meta.value()?.parse::<LitByteStr>()?);
                    // Update the XML tag name accordingly.
                    } else if meta.path.is_ident("name") {
                        field_name_str = meta.value()?.parse::<LitStr>()?.value();
                    // Mark this field to be ignored.
                    } else if meta.path.is_ident("skip") {
                        skip = true;
                    // Indicate that this field represents inner text.
                    } else if meta.path.is_ident("val") {
                        inner_value = true;
                        inner_value_found = true;
                    // Set to account for following iteration fields to act as elements.
                    } else if meta.path.is_ident("following_elements") {
                        if !inner_value_found {
                            following_elements = true;
                        } else {
                            return Err(meta.error(format!(
                                "Either a single `val` option is allowed or multiple elements `{}`",
                                meta.path.clone().into_token_stream()
                            )));
                        }
                    // Mark the field to be read as an XML element.
                    } else if meta.path.is_ident("element") {
                        if !inner_value_found {
                            element = true;
                        } else {
                            return Err(meta.error(format!(
                                "Either a single `val` option is allowed or multiple elements `{}`",
                                meta.path.clone().into_token_stream()
                            )));
                        }
                    } else if meta.path.is_ident("sequence") {
                        // ignore applies to xml reader only
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
        if inner_value {
            inner_text = quote! {self.#field_name}
        } else if !element && !following_elements {
            let attr_write_logic = match &field.ty {
                syn::Type::Path(type_path) => {
                    let last_segment = type_path.path.segments.last().unwrap();
                    match last_segment.ident.to_string().as_str() {
                        "bool" => {
                            if let Some(default_bool) = default_bool {
                                Ok(quote! {
                                    if self.#field_name != #default_bool {
                                        let value = if self.#field_name { b"1" } else { b"0" };
                                        attrs.push((#field_name_str.as_bytes(), value.as_ref()));
                                    }
                                })
                            } else {
                                Ok(quote! {
                                    let value = if self.#field_name { b"1" } else { b"0" };
                                    attrs.push((#field_name_str.as_bytes(), value.as_ref()));
                                })
                            }
                        }
                        "Vec" => {
                            // Handle Vec<u8> fields only for attributes
                            match &type_path.path.segments[0].arguments {
                                syn::PathArguments::AngleBracketed(args) => {
                                    if let syn::GenericArgument::Type(inner_type) = &args.args[0] {
                                        if inner_type.to_token_stream().to_string() == "u8" {
                                            if let Some(default_bytes) = default_bytes {
                                                Ok(quote! {
                                                    if self.#field_name != #default_bytes {
                                                        attrs.push((#field_name_str.as_bytes(), self.#field_name.as_ref()));;
                                                    }
                                                })
                                            } else {
                                                Ok(quote! {
                                                    if !self.#field_name.is_empty() {
                                                        attrs.push((#field_name_str.as_bytes(), self.#field_name.as_ref()));;
                                                    }
                                                })
                                            }
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
                                        "Unsupported path type `{}` for Vec attribute",
                                        arg.into_token_stream()
                                    ),
                                )),
                            }
                        }
                        "Option" => {
                            // Handle Vec<u8> fields only for attributes
                            match &type_path.path.segments[0].arguments {
                                syn::PathArguments::AngleBracketed(args) => {
                                    if let syn::GenericArgument::Type(inner_type) = &args.args[0] {
                                        let inner = inner_type.to_token_stream().to_string();

                                        if inner == "Vec < u8 >" {
                                            if let Some(default_bytes) = default_bytes {
                                                Ok(quote! {
                                                    if let Some(value) = &self.#field_name {
                                                        if value != #default_bytes {
                                                            attrs.push((#field_name_str.as_bytes(), value.as_ref()));
                                                        }
                                                    }
                                                })
                                            } else {
                                                Ok(quote! {
                                                    if let Some(value) = &self.#field_name {
                                                        attrs.push((#field_name_str.as_bytes(), value.as_ref()));
                                                    }
                                                })
                                            }
                                        } else if inner == "bool" {
                                            if let Some(default_bool) = default_bool {
                                                Ok(quote! {
                                                    if let Some(value) = &self.#field_name {
                                                        if value != #default_bool {
                                                            attrs.push((#field_name_str.as_bytes(), value.as_ref()));
                                                        }
                                                    }
                                                })
                                            } else {
                                                Ok(quote! {
                                                    if let Some(value) = &self.#field_name {
                                                        let value = if *value { b"1" } else { b"0" };
                                                        attrs.push((#field_name_str.as_bytes(), value.as_ref()));
                                                    }
                                                })
                                            }
                                        } else {
                                            Err(Error::new(
                                                inner_type.span(),
                                                format!("Unsupported inner type `{}` for Optional attribute, only `Vec<u8>` or `bool` is supported. Specify `#[xml(element)]` if you want to serialize it as an element",
                                            inner_type.into_token_stream()),
                                            ))
                                        }
                                    } else {
                                        let generic = &args.args[0];
                                        Err(Error::new(
                                            generic.span(),
                                            format!(
                                                "Unsupported Optional inner type `{}` for attribute",
                                                generic.into_token_stream()
                                            ),
                                        ))
                                    }
                                }
                                arg => Err(Error::new(
                                    arg.span(),
                                    format!(
                                        "Unsupported path type `{}` for optional attribute",
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

    // Some fields may have a rust like name but we
    // need a name to always match xml tag syntax
    let tag_name = if let Some(name_str) = name_str {
        quote! { #name_str }
    } else {
        quote! { tag_name }
    };

    // Generate the implementation for the `XmlWriter` trait for the struct
    let expanded = if !variants_fields.is_empty() {
        // Generated writer trait for enum data type
        quote! {
            impl<W: Write> XmlWriter<W> for #name {
                fn write_xml<'a>(
                    &self,
                    writer: &'a mut Writer<W>,
                    tag_name: &str,
                ) -> Result<&'a mut Writer<W>, XlsxError> {
                    match self {
                        #(#element_writers)*
                    };
                    Ok(writer)
                }
            }
        }
    } else {
        let writer = if inner_value_found {
            // Writes only inner text
            quote! {
                let text = String::from_utf8_lossy(&#inner_text);
                writer
                    .create_element(#tag_name)
                    .with_attributes(attrs)
                    .write_text_content(BytesText::new(&text))?;
            }
        } else if element_writers.is_empty() {
            // Allows to only write attributes
            quote! {
                writer
                    .create_element(#tag_name)
                    .with_attributes(attrs)
                    .write_empty()?;
            }
        } else {
            // Writes nested elements
            quote! {
                writer
                    .create_element(#tag_name)
                    .with_attributes(attrs)
                    .write_inner_content::<_, XlsxError>(|writer| {
                    // Generated element writing logic
                    #(#element_writers)*
                    Ok(())
                })?;
            }
        };

        quote! {
            impl<W: Write> XmlWriter<W> for #name {
                fn write_xml<'a>(
                    &self,
                    writer: &'a mut Writer<W>,
                    tag_name: &'a str,
                ) -> Result<&'a mut Writer<W>, XlsxError> {
                    let mut attrs: Vec<(&[u8], &[u8])> = Vec::new();
                    // Generated attribute writing logic
                    #(#attr_writers)*

                    #writer

                    Ok(writer)
                }
            }
        }
    };
    TokenStream::from(expanded)
}
