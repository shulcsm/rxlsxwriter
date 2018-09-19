use xml::writer::{EventWriter, EmitterConfig, XmlEvent};
use std::io::{Write, Seek};

use super::workbook::WorkBook;
use std::path::Path;
use std::fs::File;
use zip::{ZipWriter, CompressionMethod, write::FileOptions};


fn write_content_types(workbook: &WorkBook, writer: &mut ZipWriter<impl Write + Seek>) {
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);
    writer.start_file("[Content_Types].xml", options).unwrap();

    let mut writer = EmitterConfig::new().perform_indent(true)
        .create_writer(writer);

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

    for (i, sheet) in workbook.sheets.iter().enumerate() {
        writer.write(
            XmlEvent::start_element("Override")
                .attr("PartName", &format!("xl/worksheets/sheet{}.xml", i + 1))
                .attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
        writer.write(XmlEvent::end_element());
    }
    writer.write(XmlEvent::end_element());
}

fn write_root_rels(workbook: &WorkBook, writer: &mut ZipWriter<impl Write + Seek>) {
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);
    writer.start_file("_rels/.rels", options).unwrap();

    // Should this whole thing just be static string/file?
    let mut writer = EmitterConfig::new().perform_indent(true)
        .create_writer(writer);

    writer.write(
        XmlEvent::start_element("Relationships")
            .default_ns("http://schemas.openxmlformats.org/package/2006/relationships"));

    writer.write(
        XmlEvent::start_element("Relationship")
            .attr("Id", "rId1")
            .attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
            .attr("Target", "xl/workbook.xml"));
    writer.write(XmlEvent::end_element());

    writer.write(
        XmlEvent::start_element("Relationship")
            .attr("Id", "rId2")
            .attr("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
            .attr("Target", "docProps/core.xml"));
    writer.write(XmlEvent::end_element());

    writer.write(
        XmlEvent::start_element("Relationship")
            .attr("Id", "rId3")
            .attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
            .attr("Target", "docProps/app.xml"));
    writer.write(XmlEvent::end_element());

    writer.write(XmlEvent::end_element());
}

fn write_workbook_rels(workbook: &WorkBook, writer: &mut ZipWriter<impl Write + Seek>) {
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);
    writer.start_file("xl/_rels/workbook.xml.rels", options).unwrap();

    let mut writer = EmitterConfig::new().perform_indent(true)
        .create_writer(writer);

    writer.write(
        XmlEvent::start_element("Relationships")
            .default_ns("http://schemas.openxmlformats.org/package/2006/relationships")
    );


    for (i, sheet) in workbook.sheets.iter().enumerate() {
        writer.write(
            XmlEvent::start_element("Relationship")
                .attr("Id", &format!("rId{}", i + 1))
                .attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
                .attr("Target", &format!("worksheets/sheet{}.xml", i +1 ))

        );
        writer.write(XmlEvent::end_element());
    }

    // @TODO theme
    // @TODO styles
    // @TODO shared strings

    writer.write(XmlEvent::end_element());
}

pub fn write_document(workbook: &WorkBook, dst_path: String) {
    let path = Path::new(&dst_path);
    let file = File::create(&path).unwrap();

    let mut zip = ZipWriter::new(file);

    write_content_types(workbook, &mut zip);
    write_root_rels(workbook, &mut zip);
    write_workbook_rels(workbook, &mut zip);

    zip.finish().unwrap();
}