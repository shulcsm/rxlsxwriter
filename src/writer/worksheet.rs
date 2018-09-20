use xml::writer::{ XmlEvent};
use std::io::{Write};

use super::super::worksheet::WorkSheet;
use super::util::XmlWriter;
use xml::EventWriter;


pub fn write_worksheet(mut w: EventWriter<impl Write>, ws: &WorkSheet) {
    w.write(
        XmlEvent::start_element("worksheet")
            .default_ns("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            .ns("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            .attr("xml:space", "preserve"));
    // @TODO sheetPr
    // @TODO dimension

    w.write(XmlEvent::start_element("sheetViews"));
    w.write(
        XmlEvent::start_element("sheetView")
            // @TODO showGridLines
            .attr("workbookViewId", "0"));

    w.write_and_close(
        XmlEvent::start_element("selection")
            // @TODO
            .attr("activeCell", "A1")
            .attr("sqref", "A1"));


    w.write(XmlEvent::end_element()); // sheetView
    w.write(XmlEvent::end_element()); // sheetViews

    w.write_and_close(
        XmlEvent::start_element("sheetFormatPr")
            .attr("defaultRowHeight", "15"));
    // @TODO cols

    // @TODO data
    // @TODO auto filter
    // @TODO hperlinks
    // @TODO charts

    w.write(XmlEvent::end_element()); // worksheet

    // @TODO charts and rels docs
}
