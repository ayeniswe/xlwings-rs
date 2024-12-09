use std::io::{BufReader, Read, Write};
use std::sync::{Arc, Mutex};
use std::{
    collections::HashMap,
    sync::{Arc, Mutex},
};
use std::{fs::File, path::Path};
use thiserror::Error;
use xlwings_serde::{
    Book, ContentType, Drawing, PreprocessNamespace, Relationship, SharedString, Sheet, Style,
    Theme,
};
use yaserde::de::from_str;
use yaserde::YaDeserialize;
use zip::{ZipArchive};

/// A new type for managing a shared string table.
///
/// The `SharedStringTable` provides an efficient way to map strings
/// to their corresponding `SharedStringItem` references, enabling
/// fast lookups and shared access across multiple worksheets.
///
/// This struct is designed to support concurrent access to shared
/// string data by enabling synchronization (e.g., using `Mutex` or `RwLock`)
/// when used in multi-threaded environments.
pub(crate) struct SharedStringTable {
    pub shared_string: SharedString,
    pub table: HashMap<String, usize>,
}
impl SharedStringTable {
    pub(crate) fn new(shared_string: SharedString) -> Self {
        let mut table = HashMap::new();
        // Maps string value to respective index in shared strings array
        for item in shared_string.strings.iter().enumerate() {
            let (idx, item) = item;
            table.insert(item.text.value.clone(), idx);
        }
        SharedStringTable {
            table,
            shared_string: shared_string,
        }
    }
}

/// Opens and parses an XML file from a ZIP archive into a deserialized object.
///
/// This function is designed to work with XML files stored in a ZIP archive. It reads the content of the specified file,
/// preprocesses it to remove unnecessary namespaces, and deserializes the XML into the specified type.
///
/// # Details
/// **Memory Considerations**:
///    - The XML file is read fully into a `String` before processing. While this approach is simple, it may be memory-intensive
///      for large files. Optimization for streaming parsing may be considered in future iterations.
///
/// **Namespace Handling**:
///    - To ensure compatibility with `YaDeserialize`, the `xmlns` namespaces are stripped using the
///      `PreprocessNamespace` trait implemented by `T`.
///    - Re-adding or handling namespaces during deserialization is the responsibility of the `T` implementation.
///
/// # Error Handling
///    - If the file does not exist in the archive, an `ProcessError::OpenXMLFileNotFound` is returned.
///    - If deserialization fails, an `ProcessError::DeserializationError` is returned with the error message.
/// ```
fn open_xml_file<T: YaDeserialize + PreprocessNamespace>(
    zip_file: &mut ZipArchive<File>,
    filename: &str,
) -> Result<T, XMLError> {
    if let Ok(file) = zip_file.by_name(filename) {
        // This can be memory intesive since we will loaded a full string and ideally ii woudl like to avoid
        // this but as of now it will suffice because of the loopholes needed for xml parsing
        let mut reader = BufReader::new(file);
        let mut data = String::new();
        reader.read_to_string(&mut data).unwrap();
        // The xmlns namespaces needs to be cleared in order for yaserde to parse correctly
        // and adding it back is handle within T
        let no_namespace_data = T::strip_main_namespace(data);
        match from_str(&no_namespace_data) {
            Ok(data) => Ok(data),
            Err(e) => Err(XMLError::DeserializationError(e.to_string())),
        }
    } else {
        Err(XMLError::OpenXMLFileNotFound(filename.to_string()))
    }
}

#[derive(Error, Debug)]
pub enum XMLError {
    #[error("Open XML Format requires '{0}' but file is not found.")]
    OpenXMLFileNotFound(String),
    #[error("Failed to deserialize: {0}")]
    DeserializationError(String),
}
