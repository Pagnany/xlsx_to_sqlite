use calamine::{open_workbook, Reader, Xlsx};
use rusqlite::{params_from_iter, Connection};
use std::time::Instant;

fn main() {
    let now = Instant::now();

    let mut conn = match Connection::open("./files/data.db") {
        Ok(conn) => conn,
        Err(e) => panic!("Error opening database: {:?}", e),
    };

    let mut sql_table_create =
        String::from("CREATE TABLE IF NOT EXISTS data (id text PRIMARY KEY ");
    let mut sql_insert_into = String::from("INSERT INTO data VALUES (");

    let mut excel: Xlsx<_> = open_workbook("EIS-DTA_4500M.xlsx").expect("Datei kann nicht geöffnet werden");
    let xlsx_range = excel.worksheet_range("EIS-DTA").expect("Arbeitsmappe nicht gefunden.").expect("Kann keine Daten aus der Arbeitsmappe lesen.");

    for i in 1..=xlsx_range.width() {
        sql_table_create.push_str(&format!(",c{} TEXT NOT NULL", i));
    }

    for i in 1..=xlsx_range.width() + 1 {
        sql_insert_into.push_str(&format!("?{},", i));
    }

    sql_table_create.push_str(")");

    match conn.execute(&sql_table_create, ()) {
        Ok(_) => (),
        Err(e) => panic!("Error creating table: {:?}", e),
    }

    sql_insert_into.pop();
    sql_insert_into.push_str(")");

    let tx = conn.transaction().expect("Fehler beim Starten der Transaktion.");
    let mut stmt = tx.prepare(sql_insert_into.as_str()).expect("Fehler beim Erstellen des Statements.");

    let mut werte: Vec<String> = Vec::new();
    let mut current_line = 0;
    werte.push(current_line.to_string());
    for cell in xlsx_range.used_cells() {
        if cell.0 == current_line {
            werte.push(cell.2.to_string());
        } else {
            stmt.execute(params_from_iter(werte.iter())).expect("Fehler beim Ausführen des Statements.");

            // erster Wert
            current_line += 1;
            werte.clear();
            werte.push(current_line.to_string());
            werte.push(cell.2.to_string());
        }
    }
    stmt.finalize().expect("Fehler beim Finalisieren des Statements.");
    tx.commit().expect("Fehler beim Commiten der Transaktion.");

    let elapsed = now.elapsed();

    println!("Dauer: {:?}", elapsed);
}