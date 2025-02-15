mod enum_to_bytes;
mod reader;
mod writer;

use proc_macro::TokenStream;

/// Derive macro for generating XML serialization code.
///
/// This macro generates an implementation of the `XmlWrite` trait for the annotated struct.
/// The struct's fields can be customized using the `#[xml(...)]` attribute.
///
/// # Attributes
///
/// Note: This macro is limited to attributes and inner values of `Vec<u8>` and `bool` types.
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
///       field: Vec<u8>,
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
///       #[x(default_bool = true)]
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
/// ## `#[xml(val)]`
/// - **Purpose**: Specifies that the field will be written as a inner value.
/// - **Usage**: Applied to a single struct fields.
/// - **Example**:
///   ```rust
///   #[derive(XmlWrite)]
///   struct MyStruct {
///       #[xml(val)]
///       item: Vec<u8>,
///   }
///   ```
/// - **Notes**:
///   - Only a single field can have this attribute
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
    writer::impl_xml_writer(input)
}

/// Derive macro for deserialization code to XML.
///
/// This macro generates an implementation of the `XmlRead` trait for the annotated struct.
/// The struct's fields can be customized using the `#[xml(...)]` attribute.
///
/// # Attributes
///
/// Note: This macro is limited to attributes and inner values of `Vec<u8>` and `bool` types.
///
/// The following attributes are supported:
///
/// ## `#[xml(name = "field_name")]`
/// - **Purpose**: Specifies the name of the field in the generated XML.
/// - **Usage**: Applied to struct fields.
/// - **Example**:
///   ```rust
///   #[derive(XmlRead)]
///   struct MyStruct {
///       #[x(name = "custom_name")]
///       field: Vec<u8>,
///   }
///   ```
/// - **Notes**:
///   - The value must be a string literal (e.g., `name = "field_name"`).
///   - If not provided, the field's Rust name is used as the XML name.
///   - If the field is used at the root of a struct it will override any use in composition
///
/// ## `#[xml(element)]`
/// - **Purpose**: Specifies a field as axml element tag.
/// - **Usage**: Applied to struct fields.
/// - **Example**:
///   ```rust
///   #[derive(XmlRead)]
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
///   #[derive(XmlRead)]
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
///   #[derive(XmlRead)]
///   struct MyStruct {
///       #[xml(skip)]
///       extra_info: String,
///   }
///   ```
/// - **Notes**:
///   - The field ignores the other attribute's options
///
/// ## `#[xml(sequence)]`
/// - **Purpose**: Specifies that the following elements must follow a sequence one after the other
/// - **Usage**: Applied to a single struct fields.
/// - **Example**:
///   ```rust
///   #[derive(XmlRead)]
///   struct MyStruct {
///       #[xml(sequence)]
///       item: Vec<SomeStruct>,
///       item2: Vec<OtherStruct>,
///       item3: Vec<NewStruct>,
///   }
///   ```
/// - **Notes**:
///   - The field ignores the other attribute's options
///
/// ## `#[xml(val)]`
/// - **Purpose**: Specifies that the field will be read as a value.
/// - **Usage**: Applied to a single struct fields.
/// - **Example**:
///   ```rust
///   #[derive(XmlRead)]
///   struct MyStruct {
///       #[xml(val)]
///       item: Vec<u8>,
///   }
///   ```
/// - **Notes**:
///   - Only a single field can have this attribute
///
/// # Examples
///
/// Basic usage:
/// ```rust
/// #[derive(XmlRead)]
/// struct MyStruct {
///     #[xml(name = "active_pane", default = true)]
///     active: bool,
/// }
/// ```
///
/// This will generate XML serialization code where:
/// - The `active` field is serialized as `<MyStruct active_pane = "0"/>`.
#[proc_macro_derive(XmlRead, attributes(xml))]
pub fn derive_xml_reader(input: TokenStream) -> TokenStream {
    reader::impl_xml_reader(input)
}

/// The `EnumToBytes` macro can be used to convert an enum to/from bytes.
/// 
/// When applied, it will automatically transform the enum variants into their
/// byte representations. The top-level `camelcase` attribute will convert
/// **all variants** of the enum to camelCase.
///
/// ## `#[camelcase]`
/// - **Purpose**: Specifies that a variant or enum will use camelCase.
/// - **Usage**: Applied to a single enum variant or enum
/// - **Example**:
///   ```rust
///   #[derive(EnumToBytes)]
///   enum MyEnum {
///       #[camelCase]
///       FirstVariant,
///       SecondVariant,
///   }
///   ```
/// - **Notes**:
///   - Either `camelcase` or `name` can be used, but not both for a single variant
///
/// ## `#[name]`
/// - **Purpose**: Specifies that a variant will be renamed as specified.
/// - **Usage**: Applied to a single enum variant or enum
/// - **Example**:
///   ```rust
///   #[derive(EnumToBytes)]
///   enum MyEnum {
///       #[name = "NewVariant"]
///       FirstVariant,
///       SecondVariant,
///   }
///   ```
#[proc_macro_derive(EnumToBytes)]
pub fn derive_enum_to_bytes(input: TokenStream) -> TokenStream {
    enum_to_bytes::impl_enum_to_bytes(input)
}
