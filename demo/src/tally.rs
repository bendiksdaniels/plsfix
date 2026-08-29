// What a sheet builder wrote: labels, numbers and formulas, counted at the
// write site so the saved file can be reconciled against the builders' own
// count instead of eyeballed.

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub struct Tally {
    pub strings: usize,
    pub numbers: usize,
    pub formulas: usize,
}

impl Tally {
    pub fn cells(&self) -> usize {
        self.strings + self.numbers + self.formulas
    }
}

impl std::ops::AddAssign for Tally {
    fn add_assign(&mut self, other: Tally) {
        self.strings += other.strings;
        self.numbers += other.numbers;
        self.formulas += other.formulas;
    }
}
