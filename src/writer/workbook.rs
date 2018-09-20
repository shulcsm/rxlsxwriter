use xml::writer::{EmitterConfig, XmlEvent};
use std::io::{Write, Seek};

use super::super::workbook::WorkBook;
use zip::{ZipWriter, CompressionMethod, write::FileOptions};
use chrono::Utc;

use super::util::XmlWriter;


pub fn write_content_types(workbook: &WorkBook, writer: &mut ZipWriter<impl Write + Seek>) {
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);
    writer.start_file("[Content_Types].xml", options).unwrap();

    let mut writer = EmitterConfig::new().perform_indent(true)
        .create_writer(writer);

    writer.write(XmlEvent::start_element("Types")
        .default_ns("http://schemas.openxmlformats.org/package/2006/content-types"));

    // @TODO Theme
    // @TODO Styles

    writer.write_and_close(
        XmlEvent::start_element("Default")
            .attr("Extension", "xml")
            .attr("ContentType", "application/vnd.openxmlformats-package.relationships+xml"));

    writer.write_and_close(
        XmlEvent::start_element("Default")
            .attr("Extension", "rels")
            .attr("ContentType", "application/xml"));

    writer.write_and_close(
        XmlEvent::start_element("Override")
            .attr("PartName", "/docProps/core.xml")
            .attr("ContentType", "application/vnd.openxmlformats-package.core-properties+xml"));

    writer.write_and_close(
        XmlEvent::start_element("Override")
            .attr("PartName", "/docProps/app.xml")
            .attr("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml"));

    writer.write_and_close(
        XmlEvent::start_element("Override")
            .attr("PartName", "/xl/workbook.xml")
            .attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"));

    for (i, _sheet) in workbook.sheets.iter().enumerate() {
        writer.write_and_close(
            XmlEvent::start_element("Override")
                .attr("PartName", &format!("xl/worksheets/sheet{}.xml", i + 1))
                .attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
    }
    writer.write(XmlEvent::end_element());
}

pub fn write_root_rels(workbook: &WorkBook, writer: &mut ZipWriter<impl Write + Seek>) {
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);
    writer.start_file("_rels/.rels", options).unwrap();

    // Should this whole thing just be static string/file?
    let mut writer = EmitterConfig::new().perform_indent(true)
        .create_writer(writer);

    writer.write(
        XmlEvent::start_element("Relationships")
            .default_ns("http://schemas.openxmlformats.org/package/2006/relationships"));

    writer.write_and_close(
        XmlEvent::start_element("Relationship")
            .attr("Id", "rId1")
            .attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
            .attr("Target", "xl/workbook.xml"));

    writer.write_and_close(
        XmlEvent::start_element("Relationship")
            .attr("Id", "rId2")
            .attr("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
            .attr("Target", "docProps/core.xml"));

    writer.write_and_close(
        XmlEvent::start_element("Relationship")
            .attr("Id", "rId3")
            .attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
            .attr("Target", "docProps/app.xml"));

    writer.write(XmlEvent::end_element());
}

pub fn write_workbook_rels(workbook: &WorkBook, writer: &mut ZipWriter<impl Write + Seek>) {
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);
    writer.start_file("xl/_rels/workbook.xml.rels", options).unwrap();

    let mut writer = EmitterConfig::new().perform_indent(true)
        .create_writer(writer);

    writer.write(
        XmlEvent::start_element("Relationships")
            .default_ns("http://schemas.openxmlformats.org/package/2006/relationships")
    );

    for (i, sheet) in workbook.sheets.iter().enumerate() {
        writer.write_and_close(
            XmlEvent::start_element("Relationship")
                .attr("Id", &format!("rId{}", i + 1))
                .attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
                .attr("Target", &format!("worksheets/sheet{}.xml", i + 1))
        );
    }

    // @TODO theme
    // @TODO styles
    // @TODO shared strings

    writer.write(XmlEvent::end_element());
}

pub fn write_properties_app(workbook: &WorkBook, writer: &mut ZipWriter<impl Write + Seek>) {
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);
    writer.start_file("docProps/app.xml", options).unwrap();

    let mut writer = EmitterConfig::new().perform_indent(true)
        .create_writer(writer);

    writer.write(
        XmlEvent::start_element("Properties")
            .default_ns("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
            .ns("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
    );

    writer.write_and_close_chars(XmlEvent::start_element("Application"), "Microsoft Excel");
    writer.write_and_close_chars(XmlEvent::start_element("DocSecurity"), "0");
    writer.write_and_close_chars(XmlEvent::start_element("ScaleCrop"), "false");
    // Company???
    writer.write_and_close_chars(XmlEvent::start_element("LinksUpToDate"), "false");
    writer.write_and_close_chars(XmlEvent::start_element("SharedDoc"), "false");
    writer.write_and_close_chars(XmlEvent::start_element("HyperlinksChanged"), "false");
    // @TODO make sure this is what we want
    writer.write_and_close_chars(XmlEvent::start_element("AppVersion"), "12.0000");

    writer.write(XmlEvent::start_element("HeadingPairs"));
    writer.write(
        XmlEvent::start_element("vt:vector")
            .attr("size", "2")
            .attr("baseType", "variant")
    );

    writer.write(XmlEvent::start_element("vt:variant"));
    writer.write_and_close_chars(XmlEvent::start_element("vt:lpstr"), "Worksheets");
    writer.write(XmlEvent::end_element()); // vt:variant

    let num_sheets = &workbook.sheets.len().to_string();

    writer.write(XmlEvent::start_element("vt:variant"));
    writer.write_and_close_chars(XmlEvent::start_element("vt:i4"), num_sheets);
    writer.write(XmlEvent::end_element()); // vt:variant

    writer.write(XmlEvent::end_element()); // vt:vector
    writer.write(XmlEvent::end_element()); // HeadingPairs


    writer.write(XmlEvent::start_element("TitlesOfParts"));
    writer.write(
        XmlEvent::start_element("vt:variant")
            .attr("size", num_sheets)
            .attr("baseType", "lpstr"));

    for (i, sheet) in workbook.sheets.iter().enumerate() {
        // @TODO real sheet names when we have em
        writer.write_and_close_chars(XmlEvent::start_element("vt:lpstr"), &format!("Sheet {}", i + 1));
    }

    writer.write(XmlEvent::end_element()); // vt:variant
    writer.write(XmlEvent::end_element()); // TitlesOfParts


    writer.write(XmlEvent::end_element()); // Properties
}

pub fn write_properties_core(workbook: &WorkBook, writer: &mut ZipWriter<impl Write + Seek>) {
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);
    writer.start_file("docProps/core.xml", options).unwrap();

    let mut writer = EmitterConfig::new().perform_indent(true)
        .create_writer(writer);

    writer.write(
        XmlEvent::start_element("cp:corePropertie")
            .ns("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
            .ns("dc", "http://purl.org/dc/elements/1.1/")
            .ns("dcterms", "http://purl.org/dc/terms/")
            .ns("dcmitype", "http://purl.org/dc/dcmitype/")
            .ns("xsi", "http://www.w3.org/2001/XMLSchema-instance")
    );

    // @TODO actual props when we have tem on workbook
    writer.write_and_close_chars(XmlEvent::start_element("dc:creator"), "Creator");
    writer.write_and_close_chars(XmlEvent::start_element("cp:lastModifiedBy"), "Modifier");

    let dt = Utc::now();

    writer.write_and_close_chars(
        XmlEvent::start_element("dcterms:created")
            .attr("xsi:type", "dcterms:W3CDTF"),
        &dt.format("%Y-%m-%dT%H:%M:%SZ").to_string(),
    );

    writer.write_and_close_chars(
        XmlEvent::start_element("dcterms:modified")
            .attr("xsi:type", "dcterms:W3CDTF"),
        &dt.format("%Y-%m-%dT%H:%M:%SZ").to_string(),
    );

    writer.write(XmlEvent::end_element()); // cp:coreProperties
}

pub fn write_workbook(workbook: &WorkBook, writer: &mut ZipWriter<impl Write + Seek>) {
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);
    writer.start_file("xl/workbook.xml", options).unwrap();

    let mut writer = EmitterConfig::new().perform_indent(true)
        .create_writer(writer);

    writer.write(
        XmlEvent::start_element("workbook")
            .default_ns("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            .ns("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            .attr("xml:space", "preserve"));

    writer.write_and_close(
        XmlEvent::start_element("fileVersion")
            .attr("appName", "xl")
            .attr("lastEdited", "4")
            .attr("lowestEdited", "4")
            .attr("rupBuild", "4505"));

    writer.write_and_close(
        XmlEvent::start_element("workbookPr")
            .attr("defaultThemeVersion", "124226")
            .attr("codeName", "ThisWorkbook"));

    writer.write(XmlEvent::start_element("bookViews"));
    writer.write_and_close(
        XmlEvent::start_element("workbookView")
            .attr("activeTab", "0") // @TODO
            .attr("autoFilterDateGrouping", "1")
            .attr("firstSheet", "0")
            .attr("minimized", "0")
            .attr("showHorizontalScroll", "1")
            .attr("showSheetTabs", "1")
            .attr("showVerticalScroll", "1")
            .attr("tabRatio", "600")
            .attr("visibility", "visible")
    );
    writer.write(XmlEvent::end_element()); // bookViews

    writer.write(XmlEvent::start_element("sheets"));

    for (i, sheet) in workbook.sheets.iter().enumerate() {
        writer.write_and_close(
            XmlEvent::start_element("sheet")
                // @TODO real sheet names when we have em
                .attr("name", &format!("Sheet {}", i + 1))
                .attr("sheetId", &(i + 1).to_string())
                .attr("r:id", &format!("rId{}", i + 1))
                // @TODO sheet state
        );
    }
    writer.write(XmlEvent::end_element()); // sheets

    // @TODO named ranges
    // @TODO auto filter

    writer.write_and_close(
        XmlEvent::start_element("calcPr")
            .attr("calcId", "124519")
            .attr("calcMode", "auto")
            .attr("fullCalcOnLoad", "1"));

    writer.write(XmlEvent::end_element()); // workbook
}