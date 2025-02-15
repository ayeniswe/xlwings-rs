use proc_macro::{Span, TokenStream};
use quote::quote;
use syn::{parse_macro_input, DeriveInput, Ident, LitByteStr};

/// Converts PascalCase enum variants to lowercase, or camelCase if needed.
fn to_case_string(ident: &Ident) -> String {
    let name = ident.to_string();
    let mut chars = name.chars();
    if let Some(first) = chars.next() {
        if first.is_uppercase() && chars.clone().any(|c| c.is_lowercase()) {
            // Convert first letter to lowercase for camelCase (e.g., AnExample -> anExample)
            let mut result = first.to_lowercase().to_string();
            result.push_str(chars.as_str());
            result
        } else {
            // Fully lowercase for standard cases (e.g., YEAR -> year)
            name.to_lowercase()
        }
    } else {
        String::new()
    }
}

fn create_lit_byte_str(value: String) -> LitByteStr {
    LitByteStr::new(value.as_bytes(), Span::call_site().into())
}

pub fn impl_enum_to_bytes(input: TokenStream) -> TokenStream {
    let input = parse_macro_input!(input as DeriveInput);
    let name = &input.ident;

    // Extract enum variants
    let data = match input.data {
        syn::Data::Enum(data) => data,
        _ => panic!("EnumToBytes can only be derived for enums"),
    };

    let variants = data.variants.iter().map(|variant| {
        let ident = &variant.ident;
        let string_value = create_lit_byte_str(to_case_string(ident));
        quote! {
            #string_value => Ok(#name::#ident),
        }
    });

    let into_vec_cases = data.variants.iter().map(|variant| {
        let ident = &variant.ident;
        let string_value = create_lit_byte_str(to_case_string(ident));
        quote! {
            #name::#ident => #string_value.to_vec(),
        }
    });

    let expanded = quote! {
        impl TryFrom<Vec<u8>> for #name {
            type Error = XlsxError;

            fn try_from(value: Vec<u8>) -> Result<Self, Self::Error> {
                match value.as_slice() {
                    #(#variants)*
                    _ => {
                        let value = String::from_utf8_lossy(&value);
                        Err(XlsxError::MissingVariant(
                            stringify!(#name).into(),
                            value.into(),
                        ))
                    }
                }
            }
        }

        impl From<#name> for Vec<u8> {
            fn from(value: #name) -> Self {
                match value {
                    #(#into_vec_cases)*
                }
            }
        }
    };

    TokenStream::from(expanded)
}
