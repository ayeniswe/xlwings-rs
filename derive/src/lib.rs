use proc_macro::TokenStream;
use quote::{quote, ToTokens as _};
use syn::{
    parse_macro_input, spanned::Spanned as _, Data, DeriveInput, Error, Fields, LitBool,
    LitByteStr, LitStr,
};
/// Derive macro for generating XML serialization code.
///
/// This macro generates an implementation of the `XmlWrite` trait for the annotated struct.
/// The struct's fields can be customized using the `#[xml(...)]` attribute.
///
/// # Attributes
///
/// Note: This macro is limited to attributes of Vec<u8> and bool types.
///
/// The following attributes are supported:
///
/// ## `#[xml(name = "field_name")]`
/// - **Purpose**: Specifies the name of the field in the generated XML.
/// - **Usage**: Applied to struct fields.
/// - **Example**:
///   ```rust
///   #[derive(XmlWrite)]
///   struct MyStruct {
///       #[x(name = "custom_name")]
///       field: i32,
///   }
///   ```
/// - **Notes**:
///   - The value must be a string literal (e.g., `name = "field_name"`).
///   - If not provided, the field's Rust name is used as the XML name.
///   - If the field is used at the root of a struct it will override any use in composition
///
/// ## `#[xml(default_bool = true)]`
/// - **Purpose**: Specifies a default value for a bool field if it is not provided.
/// - **Usage**: Applied to struct fields.
/// - **Example**:
///   ```rust
///   #[derive(XmlWrite)]
///   struct MyStruct {
///       #[x(default = true)]
///       active: bool,
///   }
///   ```
/// - **Notes**:
///   - The value can be of a boolean (e.g., `default_bool = true`).
///   - If not provided, the field is treated as required.
///
/// ## `#[xml(default_bytes = true)]`
/// - **Purpose**: Specifies a default value for a Vec<u8> field if it is not provided.
/// - **Usage**: Applied to struct fields.
/// - **Example**:
///   ```rust
///   #[derive(XmlWrite)]
///   struct MyStruct {
///       #[xml(default_bytes = b"0")]
///       active: Vec<u8>,
///   }
///   ```
/// - **Notes**:
///   - The value can be of a byte string literal (e.g., `default_bytes = b"0"`).
///   - If not provided, the field is treated as required.
///
/// ## `#[xml(element)]`
/// - **Purpose**: Specifies a field as axml element tag.
/// - **Usage**: Applied to struct fields.
/// - **Example**:
///   ```rust
///   #[derive(XmlWrite)]
///   struct MyStruct {
///       #[xml(element)]
///       active: MySubStruct,
///   }
///   ```
///
/// ## `#[xml(following_elements)]`
/// - **Purpose**: Specifies all following fields to be used as an element.
/// - **Usage**: Applied to a single struct fields and the following fields are as if `xml(element)`` is applied to each following field.       
/// - **Example**:
///   ```rust
///   #[derive(XmlWrite)]
///   struct MyStruct {
///       #[xml(following_elements)]
///       active: MySubStruct,
///       active: MySubStruct2,
///       active: MySubStruct3,
///       active: MySubStruct4,
///       active: MySubStruct5,
///   }
///   ```
///
/// ## `#[xml(skip)]`
/// - **Purpose**: Specifies to skip the serialization of a field.
/// - **Usage**: Applied to a single struct fields.       
/// - **Example**:
///   ```rust
///   #[derive(XmlWrite)]
///   struct MyStruct {
///       #[xml(skip)]
///       extra_info: String,
///   }
///   ```
/// - **Notes**:
///   - The field ignores the other attribute's options
///
/// # Examples
///
/// Basic usage:
/// ```rust
/// #[derive(XmlWrite)]
/// struct MyStruct {
///     #[xml(name = "active_pane", default = true)]
///     active: bool,
/// }
/// ```
///
/// This will generate XML serialization code where:
/// - The `active` field is serialized as `<MyStruct active_pane = "0"/>`.
#[proc_macro_derive(XmlWrite, attributes(xml))]
pub fn derive_xml_writer(input: TokenStream) -> TokenStream {
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
