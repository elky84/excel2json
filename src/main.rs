use office::{Excel, DataType};

fn read_excel(path: &str, sheet: &str) {
    let mut workbook = Excel::open(path).unwrap();

    if let Ok(range) = workbook.worksheet_range(sheet) {
        let total_cells = range.get_size().0 * range.get_size().1;
        let non_empty_cells: usize = range.rows().map(|r| {
            r.iter().filter(|cell| cell != &&DataType::Empty).count()
        }).sum();
        println!("Found {} cells in {}, including {} non empty cells",
                 total_cells, sheet, non_empty_cells);
    }
    
    if workbook.has_vba() {
        let mut vba = workbook.vba_project().expect("Cannot find VbaProject");
        let vba = vba.to_mut();
        let module1 = vba.get_module("Module 1").unwrap();
        println!("Module 1 code:");
        println!("{}", module1);
        for r in vba.get_references() {
            if r.is_missing() {
                println!("Reference {} is broken or not accessible", r.name);
            }
        }
    }    
}

fn main() {
    read_excel("Character.xlsx", "Character")
}
