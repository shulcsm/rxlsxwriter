use xml::writer::{EmitterConfig, XmlEvent};
use std::io::{Write, Seek};
use zip::{ZipWriter, CompressionMethod, write::FileOptions};

use super::super::worksheet::WorkSheet;
use super::util::XmlWriter;

pub fn write_worksheet(index: usize, worksheet: &WorkSheet, writer: &mut ZipWriter<impl Write + Seek>) {
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);
    writer.start_file(format!("xl/worksheets/sheet{}.xml", index + 1), options).unwrap();

    let mut writer = EmitterConfig::new().perform_indent(true)
        .create_writer(writer);

    writer.write(
        XmlEvent::start_element("worksheet")
            .default_ns("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            .ns("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            .attr("xml:space", "preserve"));
    // @TODO sheetPr
    // @TODO dimension

    writer.write(XmlEvent::start_element("sheetViews"));
    writer.write(
        XmlEvent::start_element("sheetView")
            // @TODO showGridLines
            .attr("workbookViewId", "0"));

    writer.write_and_close(
        XmlEvent::start_element("selection")
            // @TODO
            .attr("activeCell", "A1")
            .attr("sqref", "A1"));


    writer.write(XmlEvent::end_element()); // sheetView
    writer.write(XmlEvent::end_element()); // sheetViews

    writer.write_and_close(
        XmlEvent::start_element("sheetFormatPr")
            .attr("defaultRowHeight", "15"));
    // @TODO cols

    // @TODO data
    // @TODO auto filter
    // @TODO hperlinks
    // @TODO charts

    writer.write(XmlEvent::end_element()); // worksheet

    // @TODO charts and rels docs
}
