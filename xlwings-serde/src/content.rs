use crate::{PreprocessNamespace, CONTENT_NAMESPACE};
use yaserde::{ser::to_string, YaDeserialize, YaSerialize};

/// Deserialize the .xlsx file `[Content_Type].xml`
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(rename = "Types")]
pub struct ContentType {
    #[yaserde(rename = "Default")]
    defaults: Vec<TypeChildren>,
    #[yaserde(rename = "Override")]
    overrides: Vec<TypeChildren>,
}
impl ToString for ContentType {
    fn to_string(&self) -> String {
        let original_namespaces = format!(r#"<Types {CONTENT_NAMESPACE}>"#);
        let cleared_namespaces = r#"<Types>"#;
        to_string(self)
            .unwrap()
            .replace(&cleared_namespaces, &original_namespaces)
    }
}
impl PreprocessNamespace for ContentType {}

#[derive(YaDeserialize, YaSerialize, Debug)]
struct TypeChildren {
    #[yaserde(rename = "ContentType", attribute = true)]
    content_type: String,
    #[yaserde(rename = "Extension", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    extension: Option<String>,
    #[yaserde(rename = "PartName", attribute = true)]
    #[yaserde(skip_serializing_if = "Option::is_none")]
    part_name: Option<String>,
}
