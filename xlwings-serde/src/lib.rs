pub mod content;
pub mod drawing;
pub mod relationships;
pub mod shared;
pub mod sheet;
pub mod style;
pub mod theme;
pub mod workbook;
pub use content::ContentType;
pub use drawing::Drawing;
pub use relationships::Relationship;
pub use shared::SharedString;
pub use sheet::Sheet;
pub use style::Style;
pub use theme::Theme;
pub use workbook::Book;

// Common namespaces seen in first tag seen in excel files
const MAIN_NAMESPACE: &str = r#"xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#;
const CONTENT_NAMESPACE: &str =
    r#"xmlns="http://schemas.openxmlformats.org/package/2006/content-types""#;
const REL_NAMESPACE: &str =
    r#"xmlns="http://schemas.openxmlformats.org/package/2006/relationships""#;

pub trait PreprocessNamespace {
    /// Strips away the beginning namespace `xmlns=` from xml first tag to enable yaserde to function properly
    fn strip_main_namespace(data: String) -> String {
        data.replacen(MAIN_NAMESPACE, "", 1)
            .replacen(REL_NAMESPACE, "", 1)
            .replacen(CONTENT_NAMESPACE, "", 1)
    }
}
