mod workbook;
mod worksheet;
mod util;

use super::workbook::WorkBook;
use std::path::Path;
use std::fs::File;
use zip::{ZipWriter};

use self::workbook::{write_content_types, write_root_rels, write_workbook_rels, write_properties_app, write_properties_core, write_workbook};
use self::worksheet::write_worksheet;


pub fn write_document(workbook: &WorkBook, dst_path: String) {
    let path = Path::new(&dst_path);
    let file = File::create(&path).unwrap();

    let mut zip = ZipWriter::new(file);

    write_content_types(workbook, &mut zip);
    write_root_rels(workbook, &mut zip);
    write_workbook_rels(workbook, &mut zip);
    write_properties_app(workbook, &mut zip);
    write_properties_core(workbook, &mut zip);

    // @TODO shared strings
    // @TODO theme
    // @TODO style
    write_workbook(workbook, &mut zip);

    for (i, sheet) in workbook.sheets.iter().enumerate() {
        write_worksheet(i, sheet, &mut zip);
    }

    zip.finish().unwrap();
}