use std::thread;
use std::time::Duration;
use windows::core::GUID;
use hello_com_rust::*;

fn main() {
    println!("initializing com...");
    co_initialize().unwrap();

    let prog_id = "Excel.Application";
    let clsid: GUID = clsid_from_prog_id(prog_id).unwrap();
    println!("resolved prog_id {} -> clsid {:?}", prog_id, clsid);

    let excel = co_create_dispatch(&clsid).unwrap();
    println!("created excel: {}", excel);

    let visible_prop1 = excel.get_property("Visible").unwrap();
    println!("visible_prop1: {}", visible_prop1);

    excel.put_property("Visible", &Variant::from(true)).unwrap();
    println!("set excel to become visible");

    println!("pausing for 2 secs...");
    thread::sleep(Duration::from_secs(2));

    let workbooks = excel.get_property("Workbooks").unwrap().to_dispatch().unwrap();
    println!("workbooks: {}", workbooks);

    let new_workbook = workbooks.call_method("Add", &[]).unwrap().to_dispatch().unwrap();
    println!("new_workbook: {}", new_workbook);

    let active_sheet = new_workbook.get_property("ActiveSheet").unwrap().to_dispatch().unwrap();
    println!("active_sheet: {}", active_sheet);

    active_sheet.put_property("Name", &Variant::from("My Test Sheet!")).unwrap();
    println!("successfully set sheet name");

    println!("pausing for 5 secs");
    thread::sleep(Duration::from_secs(5));

    println!("done, exiting!");
}