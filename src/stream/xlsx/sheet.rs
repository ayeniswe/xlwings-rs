use crate::{
    errors::XcelmateError,
    stream::utils::{xml_reader, Key, Save, XmlWriter},
};
use bimap::{BiBTreeMap, BiHashMap, BiMap};
use quick_xml::{
    events::{BytesDecl, BytesStart, Event},
    name::QName,
    Reader, Writer,
};
use std::{
    collections::HashMap,
    io::{BufRead, Read, Seek, Write},
    ops::RangeInclusive,
    sync::Arc,
};
use zip::{
    write::{FileOptionExtension, FileOptions},
    ZipArchive,
};

struct Worksheet {  

}
