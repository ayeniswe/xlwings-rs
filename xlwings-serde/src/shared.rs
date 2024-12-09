use crate::{PreprocessNamespace, MAIN_NAMESPACE};
use serde::Deserialize;
use yaserde::{ser::to_string, YaDeserialize, YaSerialize};

/// Deserialize the .xlsx file `xl/sharedStrings.xml`
#[derive(YaSerialize, YaDeserialize, Deserialize, Debug)]
#[yaserde(rename = "sst")]
pub struct SharedString {
    #[yaserde(rename = "count", attribute = true)]
    pub count: String,
    #[yaserde(rename = "uniqueCount", attribute = true)]
    pub unique_count: String,
    #[yaserde(rename = "si")]
    pub strings: Vec<SharedStringItem>,
}
impl ToString for SharedString {
    fn to_string(&self) -> String {
        let original_namespaces = format!(
            r#"<sst {MAIN_NAMESPACE} count="{}" uniqueCount="{}">"#,
            self.count, self.unique_count
        );
        let cleared_namespaces = format!(
            r#"<sst count="{}" uniqueCount="{}">"#,
            self.count, self.unique_count
        );
        to_string(self)
            .unwrap()
            .replace(&cleared_namespaces, &original_namespaces)
    }
}
impl PreprocessNamespace for SharedString {}

#[derive(YaSerialize, YaDeserialize, Deserialize, Debug)]
pub struct SharedStringItem {
    #[yaserde(rename = "t")]
    pub text: Text,
}
#[derive(YaSerialize, YaDeserialize, Deserialize, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Text {
    #[yaserde(rename = "space", attribute = true, prefix = "xml")]
    pub space: Option<String>,
    #[yaserde(text = true)]
    pub value: String,
}
