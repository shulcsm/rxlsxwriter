extern crate xml;
extern crate zip;
extern crate chrono;

mod workbook;
mod worksheet;
mod cell;
mod writer;

pub use crate::workbook::WorkBook;
pub use crate::writer::write_document;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        let mut wb = WorkBook::new();
        let mut ws = wb.create_worksheet();
//        ws.set_value(1, 1, "A");
        write_document(&wb, format!("{}/{}", env!("CARGO_MANIFEST_DIR"), "test.zip"));
    }
}
