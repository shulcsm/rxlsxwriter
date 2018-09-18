use xml::writer::{EventWriter, EmitterConfig, XmlEvent};
use std::io::{Write, Seek};

use super::workbook::WorkBook;
use std::path::Path;
use std::fs::File;
use zip::{ZipWriter, CompressionMethod, write::FileOptions};


fn write_content_types(workbook: &WorkBook, writer: &mut EventWriter<ZipWriter<impl Write + Seek>>) {
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);
    writer.inner_mut().start_file("[Content_Types].xml", options).unwrap();

    writer.write(XmlEvent::start_element("Types")
        .default_ns("http://schemas.openxmlformats.org/package/2006/content-types"));

    // @TODO Theme
    // @TODO Styles

    writer.write(
        XmlEvent::start_element("Default")
            .attr("Extension", "xml")
            .attr("ContentType", "application/vnd.openxmlformats-package.relationships+xml"));
    writer.write(XmlEvent::end_element());

    writer.write(
        XmlEvent::start_element("Default")
            .attr("Extension", "rels")
            .attr("ContentType", "application/xml"));
    writer.write(XmlEvent::end_element());

    writer.write(
        XmlEvent::start_element("Override")
            .attr("PartName", "/docProps/core.xml")
            .attr("ContentType", "application/vnd.openxmlformats-package.core-properties+xml"));
    writer.write(XmlEvent::end_element());

    writer.write(
        XmlEvent::start_element("Override")
            .attr("PartName", "/docProps/app.xml")
            .attr("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml"));
    writer.write(XmlEvent::end_element());

    writer.write(
        XmlEvent::start_element("Override")
            .attr("PartName", "/xl/workbook.xml")
            .attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"));
    writer.write(XmlEvent::end_element());

    // @TODO loop over sheets
    // <Override PartName="/xl/worksheets/sheet{{ index }}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />

    writer.write(XmlEvent::end_element());
}


pub fn write_document(workbook: &WorkBook, dst_path: String) {
    let path = Path::new(&dst_path);
    let file = File::create(&path).unwrap();

    let mut zip = ZipWriter::new(file);

    let mut writer = EmitterConfig::new().perform_indent(true)
        .create_writer(zip);

    write_content_types(workbook, &mut writer);
    writer.inner_mut().finish().unwrap();
}