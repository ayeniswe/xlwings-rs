use proc_macro::{Span, TokenStream};
use quote::quote;
use syn::{parse_macro_input, DeriveInput, Data, Error, Ident, LitByteStr, LitStr};

fn to_camel_case(value: String) -> String {
    let mut chars = value.chars();
    if let Some(first) = chars.next() {
        if first.is_uppercase() && chars.clone().any(|c| c.is_lowercase()) {
            // Convert first letter to lowercase for camelCase (e.g., AnExample -> anExample)
            let mut result = first.to_lowercase().to_string();
            result.push_str(chars.as_str());
            result
        } else {
            // Fully lowercase for standard cases (e.g., YEAR -> year)
            value.to_lowercase()
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
    
    // Check if all variants should be camelCase
    let mut global_camel_case = false;
    for attr in &input.attrs {
        if attr.path().is_ident("camel") {
            global_camel_case = true;
        }
    }
    
    // Extract enum variants
    let data = match input.data {
        Data::Enum(data) => data,
        _ => panic!("EnumToBytes can only be derived for enums"),
    };

    let (try_from_variants, from_variants) = data.variants.iter().map(|variant| {
        let ident = &variant.ident;
        let mut ident_str = ident.to_string();
        let mut rename = None;
        let mut camel_case = globl_camel_case;
        
        // Get metadata to transform final name
        for attr in &variant.attrs {
            if attr.path().is_ident("name") {
                rename = Some(attr.parse_args::<LitStr>().expect("expected a string for rename").value());
            } else if attr.path().is_ident("camel") {
                camel_case = true;
            }
        }
        
        // Prevent using both rename and camelcase
        if rename.is_some() && (camel_case || global_camel_case) {
            return (
                Error::new_spanned(variant, "Cannot use both 'rename' and 'camelcase' attributes").to_compile_error(),
                quote! {}
            );
        }
    
        // Apply transformation
        if let Some(rename) = rename {
            ident_str = rename;
        } else if camel_case {
            ident_str = to_camel_case(ident_str);
        }
    
        let lit_byte = create_lit_byte_str(ident_str);
        (
            quote! {
                #lit_byte => Ok(#name::#ident),
            },
            quote! {
                #name::#ident => #lit_byte.to_vec(),
            }
        )
    }).unzip::<(Vec<_>, Vec<_>)>();
    
    let expanded = quote! {
        impl TryFrom<Vec<u8>> for #name {
            type Error = XlsxError;

            fn try_from(value: Vec<u8>) -> Result<Self, Self::Error> {
                match value.as_slice() {
                    #(#try_from_variants)*
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
                    #(#from_variants)*
                }
            }
        }
    };

    TokenStream::from(expanded)
}
