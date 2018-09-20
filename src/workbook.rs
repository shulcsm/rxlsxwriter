use super::worksheet::WorkSheet;

pub struct WorkBook {
    // @TODO document properties
    // @TODO activeTab/sheet
    pub sheets: Vec<WorkSheet>
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