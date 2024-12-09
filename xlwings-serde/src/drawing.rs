use yaserde::{ser::to_string, YaDeserialize, YaSerialize};

use crate::PreprocessNamespace;

/// Deserialize the .xlsx file(s) `xl/drawing/drawing1.xml`
#[derive(YaDeserialize, YaSerialize, Debug)]
#[yaserde(prefix=
    "xdr", rename = "wsDr", namespaces = {
        "xdr"="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a"="http://schemas.openxmlformats.org/drawingml/2006/main",
        "r"="http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "c"="http://schemas.openxmlformats.org/drawingml/2006/chart",
        "cx"="http://schemas.microsoft.com/office/drawing/2014/chartex",
        "cx1"="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
        "mc"="http://schemas.openxmlformats.org/markup-compatibility/2006",
        "dgm"="http://schemas.openxmlformats.org/drawingml/2006/diagram",
        "x3Unk"="http://schemas.microsoft.com/office/drawing/2010/slicer",
        "sle15"="http://schemas.microsoft.com/office/drawing/2012/slicer"
})]
pub struct Drawing {}
impl ToString for Drawing {
    fn to_string(&self) -> String {
        to_string(self).unwrap()
    }
}
impl PreprocessNamespace for Drawing {}
