extern crate xml;
extern crate zip;
extern crate chrono;

mod workbook;
mod worksheet;
mod cell;
mod writer;

pub use workbook::WorkBook;
pub use writer::write_document;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        let mut wb = WorkBook::new();
        {
            let mut ws = wb.create_worksheet();
        }
        write_document(&wb, format!("{}/{}", env!("CARGO_MANIFEST_DIR"), "test.zip"));
    }
}
