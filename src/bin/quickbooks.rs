use std::thread;
use std::time::Duration;
use windows::core::GUID;
use hello_com_rust::*;

fn main() {
    let data_file = "C:\\Users\\Public\\Documents\\Intuit\\QuickBooks\\Company Files\\Fizzed Consulting.qbw";

    println!("initializing com...");
    co_initialize().unwrap();

    let prog_id = "QBXMLRP2.RequestProcessor";
    let clsid: GUID = clsid_from_prog_id(prog_id).unwrap();
    println!("resolved prog_id {} -> clsid {:?}", prog_id, clsid);

    let request_processor = co_create_dispatch(&clsid).unwrap();
    println!("request_processor: {}", request_processor);

    let connected = request_processor.call_method("OpenConnection2", &[
        Variant::from(""), Variant::from("Dagger Desktop"), Variant::from(1)
    ]).unwrap();
    println!("connected: {}", connected);

    let session = request_processor.call_method("BeginSession", &[
        Variant::from(data_file), Variant::from(1)
    ]).unwrap();
    println!("session: {}", session);

    let request_xml = "<?xml version=\"1.0\" ?>\
        <?qbxml version=\"8.0\"?>\
        <QBXML>\
          <QBXMLMsgsRq onError=\"stopOnError\">\
            <HostQueryRq requestID=\"1\" />\
          </QBXMLMsgsRq>\
        </QBXML>";

    let response_xml = request_processor.call_method("ProcessRequest", &[session, Variant::from(request_xml)]).unwrap();
    println!("response_xml: {}", response_xml);

    println!("pausing for 5 secs");
    thread::sleep(Duration::from_secs(5));

    println!("done, exiting!");
}