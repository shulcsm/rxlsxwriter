use xml::writer::{XmlEvent};
use std::io::{Write};

use super::super::workbook::WorkBook;
use chrono::Utc;

use super::util::XmlWriter;
use xml::EventWriter;


pub fn write_content_types(mut w: EventWriter<impl Write>, workbook: &WorkBook) {
    w.write(XmlEvent::start_element("Types")
        .default_ns("http://schemas.openxmlformats.org/package/2006/content-types"));

    // @TODO Theme
    // @TODO Styles

    w.write_and_close(
        XmlEvent::start_element("Default")
            .attr("Extension", "xml")
            .attr("ContentType", "application/vnd.openxmlformats-package.relationships+xml"));

    w.write_and_close(
        XmlEvent::start_element("Default")
            .attr("Extension", "rels")
            .attr("ContentType", "application/xml"));

    w.write_and_close(
        XmlEvent::start_element("Override")
            .attr("PartName", "/docProps/core.xml")
            .attr("ContentType", "application/vnd.openxmlformats-package.core-properties+xml"));

    w.write_and_close(
        XmlEvent::start_element("Override")
            .attr("PartName", "/docProps/app.xml")
            .attr("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml"));

    w.write_and_close(
        XmlEvent::start_element("Override")
            .attr("PartName", "/xl/workbook.xml")
            .attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"));

    for (i, _sheet) in workbook.sheets.iter().enumerate() {
        w.write_and_close(
            XmlEvent::start_element("Override")
                .attr("PartName", &format!("xl/worksheets/sheet{}.xml", i + 1))
                .attr("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
    }
    w.write(XmlEvent::end_element());
}

pub fn write_root_rels(mut w: EventWriter<impl Write>, _workbook: &WorkBook) {
    w.write(
        XmlEvent::start_element("Relationships")
            .default_ns("http://schemas.openxmlformats.org/package/2006/relationships"));

    w.write_and_close(
        XmlEvent::start_element("Relationship")
            .attr("Id", "rId1")
            .attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
            .attr("Target", "xl/workbook.xml"));

    w.write_and_close(
        XmlEvent::start_element("Relationship")
            .attr("Id", "rId2")
            .attr("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
            .attr("Target", "docProps/core.xml"));

    w.write_and_close(
        XmlEvent::start_element("Relationship")
            .attr("Id", "rId3")
            .attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
            .attr("Target", "docProps/app.xml"));

    w.write(XmlEvent::end_element());
}

pub fn write_workbook_rels(mut w: EventWriter<impl Write>, workbook: &WorkBook) {
    w.write(
        XmlEvent::start_element("Relationships")
            .default_ns("http://schemas.openxmlformats.org/package/2006/relationships"));

    for (i, _sheet) in workbook.sheets.iter().enumerate() {
        w.write_and_close(
            XmlEvent::start_element("Relationship")
                .attr("Id", &format!("rId{}", i + 1))
                .attr("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
                .attr("Target", &format!("worksheets/sheet{}.xml", i + 1))
        );
    }

    // @TODO theme
    // @TODO styles
    // @TODO shared strings

    w.write(XmlEvent::end_element());
}

pub fn write_properties_app(mut w: EventWriter<impl Write>, workbook: &WorkBook) {
    w.write(
        XmlEvent::start_element("Properties")
            .default_ns("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
            .ns("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"));

    w.write_and_close_chars(XmlEvent::start_element("Application"), "Microsoft Excel");
    w.write_and_close_chars(XmlEvent::start_element("DocSecurity"), "0");
    w.write_and_close_chars(XmlEvent::start_element("ScaleCrop"), "false");
    // Company???
    w.write_and_close_chars(XmlEvent::start_element("LinksUpToDate"), "false");
    w.write_and_close_chars(XmlEvent::start_element("SharedDoc"), "false");
    w.write_and_close_chars(XmlEvent::start_element("HyperlinksChanged"), "false");
    // @TODO make sure this is what we want
    w.write_and_close_chars(XmlEvent::start_element("AppVersion"), "12.0000");

    w.write(XmlEvent::start_element("HeadingPairs"));
    w.write(
        XmlEvent::start_element("vt:vector")
            .attr("size", "2")
            .attr("baseType", "variant"));

    w.write(XmlEvent::start_element("vt:variant"));
    w.write_and_close_chars(XmlEvent::start_element("vt:lpstr"), "Worksheets");
    w.write(XmlEvent::end_element()); // vt:variant

    let num_sheets = &workbook.sheets.len().to_string();

    w.write(XmlEvent::start_element("vt:variant"));
    w.write_and_close_chars(XmlEvent::start_element("vt:i4"), num_sheets);
    w.write(XmlEvent::end_element()); // vt:variant

    w.write(XmlEvent::end_element()); // vt:vector
    w.write(XmlEvent::end_element()); // HeadingPairs


    w.write(XmlEvent::start_element("TitlesOfParts"));
    w.write(
        XmlEvent::start_element("vt:variant")
            .attr("size", num_sheets)
            .attr("baseType", "lpstr"));

    for (i, _sheet) in workbook.sheets.iter().enumerate() {
        // @TODO real sheet names when we have em
        w.write_and_close_chars(XmlEvent::start_element("vt:lpstr"), &format!("Sheet {}", i + 1));
    }

    w.write(XmlEvent::end_element()); // vt:variant
    w.write(XmlEvent::end_element()); // TitlesOfParts

    w.write(XmlEvent::end_element()); // Properties
}

pub fn write_properties_core(mut w: EventWriter<impl Write>, _workbook: &WorkBook) {
    w.write(
        XmlEvent::start_element("cp:corePropertie")
            .ns("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
            .ns("dc", "http://purl.org/dc/elements/1.1/")
            .ns("dcterms", "http://purl.org/dc/terms/")
            .ns("dcmitype", "http://purl.org/dc/dcmitype/")
            .ns("xsi", "http://www.w3.org/2001/XMLSchema-instance")
    );

    // @TODO actual props when we have tem on workbook
    w.write_and_close_chars(XmlEvent::start_element("dc:creator"), "Creator");
    w.write_and_close_chars(XmlEvent::start_element("cp:lastModifiedBy"), "Modifier");

    let dt = Utc::now();

    w.write_and_close_chars(
        XmlEvent::start_element("dcterms:created")
            .attr("xsi:type", "dcterms:W3CDTF"),
        &dt.format("%Y-%m-%dT%H:%M:%SZ").to_string());

    w.write_and_close_chars(
        XmlEvent::start_element("dcterms:modified")
            .attr("xsi:type", "dcterms:W3CDTF"),
        &dt.format("%Y-%m-%dT%H:%M:%SZ").to_string());

    w.write(XmlEvent::end_element()); // cp:coreProperties
}

pub fn write_workbook(mut w: EventWriter<impl Write>, workbook: &WorkBook) {
    w.write(
        XmlEvent::start_element("workbook")
            .default_ns("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            .ns("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            .attr("xml:space", "preserve"));

    w.write_and_close(
        XmlEvent::start_element("fileVersion")
            .attr("appName", "xl")
            .attr("lastEdited", "4")
            .attr("lowestEdited", "4")
            .attr("rupBuild", "4505"));

    w.write_and_close(
        XmlEvent::start_element("workbookPr")
            .attr("defaultThemeVersion", "124226")
            .attr("codeName", "ThisWorkbook"));

    w.write(XmlEvent::start_element("bookViews"));
    w.write_and_close(
        XmlEvent::start_element("workbookView")
            .attr("activeTab", "0") // @TODO
            .attr("autoFilterDateGrouping", "1")
            .attr("firstSheet", "0")
            .attr("minimized", "0")
            .attr("showHorizontalScroll", "1")
            .attr("showSheetTabs", "1")
            .attr("showVerticalScroll", "1")
            .attr("tabRatio", "600")
            .attr("visibility", "visible"));

    w.write(XmlEvent::end_element()); // bookViews

    w.write(XmlEvent::start_element("sheets"));

    for (i, _sheet) in workbook.sheets.iter().enumerate() {
        w.write_and_close(
            XmlEvent::start_element("sheet")
                // @TODO real sheet names when we have em
                .attr("name", &format!("Sheet {}", i + 1))
                .attr("sheetId", &(i + 1).to_string())
                .attr("r:id", &format!("rId{}", i + 1))
                // @TODO sheet state
        );
    }
    w.write(XmlEvent::end_element()); // sheets

    // @TODO named ranges
    // @TODO auto filter

    w.write_and_close(
        XmlEvent::start_element("calcPr")
            .attr("calcId", "124519")
            .attr("calcMode", "auto")
            .attr("fullCalcOnLoad", "1"));

    w.write(XmlEvent::end_element()); // workbook
}