use super::cell::Cell;
use std::collections::BTreeMap;

// Indexes are 0 based
// 1048576
type RowIdx = u32;


// 16384
type ColIdx = u16;

pub struct WorkSheet {
    // @TODO name
    // @TODO state
    cells: BTreeMap<(RowIdx, ColIdx), Cell>,
    max_row: RowIdx,
    max_col: ColIdx
}


impl WorkSheet {
    pub fn new() -> WorkSheet {
        WorkSheet {
            cells: BTreeMap::new(),
            max_row: 1,
            max_col: 1
        }
    }

    fn update_dimensions(&mut self, row_idx: RowIdx, col_idx: ColIdx) {
        if row_idx > self.max_row {
            self.max_row = row_idx;
        }

        if col_idx > self.max_col {
            self.max_col = col_idx;
        }
    }

    pub fn set_cell(&mut self, row_idx: RowIdx, col_idx: ColIdx, cell: Cell) {
        self.update_dimensions(row_idx, col_idx);
        self.cells.insert((row_idx, col_idx), cell);
    }

    pub fn set_value<V: Into<Cell>>(&mut self, row_idx: RowIdx, col_idx: ColIdx, value: V) {
        self.set_cell(row_idx, col_idx, value.into())
    }
}