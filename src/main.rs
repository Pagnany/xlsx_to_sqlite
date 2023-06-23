use calamine::{open_workbook, Reader, Xlsx};
use rusqlite::Connection;
use std::time::Instant;

fn main() {
    let mut conn = match Connection::open("./files/test.db") {
        Ok(conn) => conn,
        Err(e) => panic!("Error opening database: {:?}", e),
    };

    let mut sql_table_create =
        String::from("CREATE TABLE IF NOT EXISTS test (id INTEGER PRIMARY KEY ");

    let now = Instant::now();

    let mut i = 1;
    let mut excel: Xlsx<_> = open_workbook("filelarge.xlsx").unwrap();
    if let Some(Ok(r)) = excel.worksheet_range("EIS-DTA") {
        for cell in r.used_cells() {
            if cell.0 == 0 {
                sql_table_create.push_str(&format!(", c{} TEXT NOT NULL ", i));
                i += 1;
            }
        }
    }
    sql_table_create.push_str(")");
    println!("{}", sql_table_create);

    match conn.execute(&sql_table_create, ()) {
        Ok(_) => (),
        Err(e) => panic!("Error creating table: {:?}", e),
    }

    /*
    let tx = conn.transaction().unwrap();
    let mut stmt = tx.prepare("INSERT INTO test (name) VALUES (?1)").unwrap();
    for geb in gebinde {
        stmt.execute([geb]).unwrap();
    }
    stmt.finalize().unwrap();
    tx.commit().unwrap();
    */

    let elapsed = now.elapsed();

    println!("Dauer: {:?}", elapsed);
}
