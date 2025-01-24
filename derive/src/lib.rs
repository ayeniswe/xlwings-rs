use proc_macro::TokenStream;
use quote::{format_ident, quote, ToTokens as _};
use syn::{
    meta::{parser, ParseNestedMeta},
    parenthesized, parse_macro_input,
    punctuated::Punctuated,
    Attribute, Data, DeriveInput, Fields, Lit, LitBool, LitByteStr, LitStr, Meta, MetaList,
    MetaNameValue, Token,
};

/// Defines uncommon possible default values for struct fields.
enum DefaultValue {
    Bool(bool),
    Bytes(LitByteStr),
}

/// Derive macro for generating XML serialization code.
///
/// This macro generates an implementation of the `XmlWriter` trait for the annotated struct.
/// The struct's fields can be customized using the `#[x(...)]` attribute.
///
/// # Attributes
///
/// The following attributes are supported:
///
/// ## `#[x(name = "field_name")]`
/// - **Purpose**: Specifies the name of the field in the generated XML.
/// - **Usage**: Applied to struct fields.
/// - **Example**:
///   ```rust
///   #[derive(XmlWriter)]
///   struct MyStruct {
///       #[x(name = "custom_name")]
///       field: i32,
///   }
///   ```
/// - **Notes**:
///   - The value must be a string literal (e.g., `name = "field_name"`).
///   - If not provided, the field's Rust name is used as the XML name.
///
/// ## `#[x(tag = "struct_name")]`
/// - **Purpose**: Specifies the name of the start/empty tag in the generated XML.
/// - **Usage**: Applied to structs.
/// - **Example**:
///   ```rust
///   #[derive(XmlWriter)]
///   #[x(tag = "sheet")]
///   struct MyStruct {
///       field: i32,
///   }
///   ```
/// - **Notes**:
///   - The value must be a string literal (e.g., `tag = "struct_name"`).
///   - If not provided, the struct's Rust name is used as the XML name.
///
/// ## `#[x(default_bool = true)]`
/// - **Purpose**: Specifies a default value for a bool field if it is not provided.
/// - **Usage**: Applied to struct fields.
/// - **Example**:
///   ```rust
///   #[derive(XmlWriter)]
///   struct MyStruct {
///       #[x(default = true)]
///       active: bool,
///   }
///   ```
/// - **Notes**:
///   - The value can be of a boolean (e.g., `default = true`).
///   - If not provided, the field is treated as required.
///
/// ## `#[x(default_bytes = true)]`
/// - **Purpose**: Specifies a default value for a Vec<u8> field if it is not provided.
/// - **Usage**: Applied to struct fields.
/// - **Example**:
///   ```rust
///   #[derive(XmlWriter)]
///   struct MyStruct {
///       #[x(default_bytes = b"0")]
///       active: Vec<u8>,
///   }
///   ```
/// - **Notes**:
///   - The value can be of a byte string literal (e.g., `default = b"0"`).
///   - If not provided, the field is treated as required.
///
/// # Examples
///
/// Basic usage:
/// ```rust
/// #[derive(XmlWriter)]
/// struct MyStruct {
///     #[x(name = "active_pane", default = true)]
///     active: bool,
/// }
/// ```
///
/// This will generate XML serialization code where:
/// - The `active` field is serialized as `<MyStruct active_pane = "0"/>`.
#[proc_macro_derive(XmlWriter, attributes(xml))]
pub fn derive_xml_writer(input: TokenStream) -> TokenStream {
    // Gather the code definition
    let input = parse_macro_input!(input as DeriveInput);
    let name = &input.ident;
    let mut name_str = name.to_string();

    // Gather top level metadata
    for attr in input.attrs {
        if attr.path().is_ident("xml") {
            let result = attr.parse_nested_meta(|meta| {
                if meta.path.is_ident("name") {
                    name_str = meta.value()?.parse::<LitStr>()?.value();
                } else {
                    return Err(meta.error(format!(
                        "Unsupported flag `{}`",
                        meta.path.clone().into_token_stream()
                    )));
                }
                Ok(())
            });
            if let Err(e) = result {
                panic!("Failed to parse `xml` attribute: {}", e);
            }
        } else {
            panic!("Unsupported XmlWriter attribute")
        }
    }

    // Gather all members of the struct
    let fields = if let Data::Struct(data_struct) = &input.data {
        match &data_struct.fields {
            Fields::Named(fields) => &fields.named,
            _ => panic!("Only struct with named fields is supported"),
        }
    } else {
        panic!("Only structs are supported")
    };

    let mut attr_writers = Vec::new();
    for field in fields {
        // Get actual field verbatim name
        let field_name = &field.ident.clone().unwrap();
        let mut field_name_str = field_name.to_string();

        // Gather field member optional metadata
        let mut default_bool = None;
        let mut default_bytes = None;
        for attr in &field.attrs {
            if attr.path().is_ident("xml") {
                let result = attr.parse_nested_meta(|meta| {
                    if meta.path.is_ident("default_bool") {
                        default_bool = Some(meta.value()?.parse::<LitBool>()?.value());
                    } else if meta.path.is_ident("default_bytes") {
                        default_bytes = Some(meta.value()?.parse::<LitByteStr>()?);
                    } else if meta.path.is_ident("name") {
                        field_name_str = meta.value()?.parse::<LitStr>()?.value();
                    } else {
                        return Err(meta.error(format!(
                            "Unsupported flag `{}`",
                            meta.path.clone().into_token_stream()
                        )));
                    }
                    Ok(())
                });
                if let Err(e) = result {
                    panic!("Failed to parse `xml` attribute: {}", e);
                }
            } else {
                panic!("Unsupported XmlWriter attribute")
            }
        }

        // Generate the logic for writing the field to XML attributes
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
                        logic
                    }
                    "Vec" => {
                        // Handle Vec<u8> fields
                        let inner_type = match &type_path.path.segments[0].arguments {
                            syn::PathArguments::AngleBracketed(args) => {
                                if let syn::GenericArgument::Type(inner_type) = &args.args[0] {
                                    inner_type
                                } else {
                                    panic!("Unsupported Vec inner type");
                                }
                            }
                            _ => panic!("Unsupported Vec type"),
                        };
                        if inner_type.to_token_stream().to_string() == "u8" {
                            let logic = if let Some(default_bytes) = default_bytes {
                                quote! {
                                    if self.#field_name != #default_bytes {
                                        attrs.push((#field_name_str.as_bytes(), self.#field_name.as_ref()));;
                                    }
                                }
                            } else {
                                quote! {
                                    attrs.push((#field_name_str.as_bytes(), self.#field_name.as_ref()));
                                }
                            };
                            logic
                        } else {
                            panic!("Only Vec<u8> is supported for Vec fields");
                        }
                    }
                    _ => panic!("Unsupported struct field datatype"),
                }
            }
            _ => panic!("Unsupported struct field type"),
        };
        attr_writers.push(attr_write_logic);
    }

    // Generate the struct definition with the added methods
    let expanded = quote! {
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
                    .create_element(#name_str)
                    .with_attributes(attrs)
                    .write_empty()?;

                Ok(writer)
            }
        }
    };
    TokenStream::from(expanded)
}
