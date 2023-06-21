use calamine::{open_workbook, Error, RangeDeserializerBuilder, Reader, Xlsx};
use std::time::Instant;

fn main() {
    let now = Instant::now();
    let mut i = 0;
    //let mut excel: Xlsx<_> = open_workbook("file.xlsx").unwrap();
    let mut excel: Xlsx<_> = open_workbook("filelarge.xlsx").unwrap();
    if let Some(Ok(r)) = excel.worksheet_range("EIS-DTA") {
        for row in r.rows() {
            //println!("row[0]={:?}", row[0]);
            let test = match row[0].as_i64() {
                Some(v) => {
                    i += 1;
                    v
                }
                None => 0,
            };
            //println!("test={}", test);
        }
    }
    let elapsed = now.elapsed();

    println!("Dauer: {:?}", elapsed);
}
