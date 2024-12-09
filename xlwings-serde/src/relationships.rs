use crate::{PreprocessNamespace, REL_NAMESPACE};
use yaserde::{ser::to_string, YaDeserialize, YaSerialize};

/// Deserialize the .xlsx file `_rels/.rels`
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(rename = "Relationships")]
pub struct Relationship {
    #[yaserde(rename = "Relationship")]
    relationships: Vec<RelationshipChildren>,
}
#[derive(YaDeserialize, YaSerialize, Debug)]
struct RelationshipChildren {
    #[yaserde(rename = "Id", attribute = true)]
    id: String,
    #[yaserde(rename = "Type", attribute = true)]
    r#type: String,
    #[yaserde(rename = "Target", attribute = true)]
    target: String,
}
impl ToString for Relationship {
    fn to_string(&self) -> String {
        let original_namespaces = format!(r#"<Relationships {REL_NAMESPACE}>"#);
        let cleared_namespaces = r#"<Relationships>"#;
        to_string(self)
            .unwrap()
            .replace(&cleared_namespaces, &original_namespaces)
    }
}
impl PreprocessNamespace for Relationship {}
