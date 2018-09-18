use super::worksheet::WorkSheet;
use std::io::Write;

pub struct WorkBook {
    sheets: Vec<WorkSheet>
}

impl WorkBook {
    pub fn new() -> WorkBook {
        WorkBook {
            sheets: Vec::new()
        }
    }

    pub fn create_worksheet(&mut self) -> &mut WorkSheet {
        let ws = WorkSheet::new();
        self.sheets.push(ws);
        self.sheets.last_mut().unwrap()
    }

}