use std::{thread};
use std::time::{Duration, Instant};
use windows::core::GUID;
use hello_com_rust::*;

fn main() {
    //
    // NOTE: JetBrains RustRover IDE causes a weird issue when connecting to SDO, but only in the IDE
    // and using the normal execution. If you run debug it works, as well as if you run the compiled
    // version on the command line
    //

    // sdo only works on 32-bit programs, 64-bit will return "class not registered"
    if std::env::consts::ARCH != "x86" {
        panic!("sdo only works on 32-bit x86 arch (current was {})", std::env::consts::ARCH);
    }

    let data_dir = "C:\\PROGRAMDATA\\SAGE\\ACCOUNTS\\2023\\COMPANY.000\\ACCDATA\\";

    println!("initializing com...");
    co_initialize().unwrap();

    // version 29 of sdo (other versions have other clsid)
    let clsid = GUID::from("663048C4-DAEA-4125-9F02-4F1DFB8F4666");
    println!("resolved clsid {:?}", clsid);

    let sdo_engine = co_create_dispatch(&clsid).unwrap();
    println!("sdo_engine: {}", sdo_engine);

    let sdo_workspaces = sdo_engine.get_property("Workspaces").unwrap().to_dispatch().unwrap();
    println!("sdo_workspaces: {}", sdo_workspaces);

    let sdo_workspace = sdo_workspaces.call_method("Add", &[Variant::from("Dext Commerce")]).unwrap().to_dispatch().unwrap();
    println!("sdo_workspace: {}", sdo_workspace);

    // let sageDir2 = "C:\\ProgramData\\Sage\\Accounts\\2023\\";
    // let selectCompanyVariant = call_method(&sdoEngineDispatch, "SelectCompany", &[
    //     Variant::from(sageDir2)
    // ]);
    // println!("selectCompany: {}", selectCompanyVariant);

    sdo_workspace.put_property("UI", &Variant::from(true)).unwrap();
    println!("setting workspace UI to false");

    let connected = sdo_workspace.call_method("Connect", &[
        Variant::from(data_dir), Variant::from("sdouser"), Variant::from("test"), Variant::from("Dext Commerce")
    ]).unwrap();
    println!("connected: {}", connected);

    let now = Instant::now();

    let setup_data = sdo_workspace.call_method("CreateObject", &[Variant::from("SetupData")]).unwrap().to_dispatch().unwrap();
    println!("setup_data: {}", setup_data);

    let moved = setup_data.call_method("MoveFirst", &[]).unwrap();
    println!("moved: {}", moved);

    let fields = setup_data.get_property("Fields").unwrap().to_dispatch().unwrap();
    println!("fields: {}", fields);

    let field_count = fields.get_property("Count").unwrap().to_i32().unwrap();
    println!("field_count: {}", field_count);

    for i in 0..field_count {
        let item = fields.call_method("Item", &[Variant::from(i+1)]).unwrap().to_dispatch().unwrap();
        let _name = item.get_property("Name").unwrap();
        let _value = item.get_property("Value").unwrap();
        println!("field: {} => {:?}", _name, _value);
    }

    println!("took {:.2?} to dump setupData", now.elapsed());

    println!("pausing for 5 secs");
    thread::sleep(Duration::from_secs(5));

    sdo_workspace.call_method("Disconnect", &[]).unwrap();
    println!("disconnected");

    println!("done, exiting!");
}