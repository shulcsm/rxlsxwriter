mod workbook;
mod worksheet;
mod util;

use super::workbook::WorkBook;
use std::path::Path;
use std::fs::File;
use zip::ZipWriter;

use self::workbook::{write_content_types, write_root_rels, write_workbook_rels, write_properties_app, write_properties_core, write_workbook};
use self::worksheet::write_worksheet;
use self::util::xml_file;


pub fn write_document(wb: &WorkBook, dst_path: String) {
    let path = Path::new(&dst_path);
    let file = File::create(&path).unwrap();

    let mut zip = ZipWriter::new(file);

    write_content_types(xml_file(&mut zip, "[Content_Types].xml"), wb);
    write_root_rels(xml_file(&mut zip, "_rels/.rels"), wb);
    write_workbook_rels(xml_file(&mut zip, "xl/_rels/workbook.xml.rels"), wb);

    write_properties_app(xml_file(&mut zip, "docProps/app.xml"), wb);
    write_properties_core(xml_file(&mut zip, "docProps/core.xml"), wb);

    // @TODO shared strings
    // @TODO theme
    // @TODO style
    write_workbook(xml_file(&mut zip, "xl/workbook.xml"), wb);

    for (i, sheet) in wb.sheets.iter().enumerate() {
        write_worksheet(xml_file(&mut zip, &format!("xl/worksheets/sheet{}.xml", i + 1)), sheet);
    }

    zip.finish().unwrap();
}