use calamine::{open_workbook, Error, RangeDeserializerBuilder, Reader, Xlsx};
use rusqlite::Connection;
use std::time::Instant;

fn main() {
    let mut conn = match Connection::open("./files/test.db") {
        Ok(conn) => conn,
        Err(e) => panic!("Error opening database: {:?}", e),
    };
    /* 
    conn.execute(
        "CREATE TABLE test (
            id    INTEGER PRIMARY KEY,
            name  TEXT NOT NULL
        )",
        (), // empty list of parameters.
    ).unwrap();
    */

    let now = Instant::now();
    let mut gebinde = Vec::new();

    //let mut excel: Xlsx<_> = open_workbook("file.xlsx").unwrap();
    let mut excel: Xlsx<_> = open_workbook("filelarge.xlsx").unwrap();
    if let Some(Ok(r)) = excel.worksheet_range("EIS-DTA") {
        for row in r.rows() {
            gebinde.push(row[68].to_string());
            gebinde.push(row[130].to_string());
            //println!("test={}", test);
        }
    }

    let tx = conn.transaction().unwrap();
    let mut stmt = tx.prepare("INSERT INTO test (name) VALUES (?1)").unwrap();
    for geb in gebinde {
        stmt.execute([geb]).unwrap();
    }
    stmt.finalize().unwrap();
    tx.commit().unwrap();

    let mut stmt = conn.prepare("SELECT distinct name FROM test").unwrap();
    let gebinde: Vec<String> = stmt
        .query_map([], |row| row.get(0))
        .unwrap()
        .map(|x| x.unwrap())
        .collect();
    for geb in gebinde {
        println!("Gebinde: {}", geb);
    }

    let elapsed = now.elapsed();

    println!("Dauer: {:?}", elapsed);
}
