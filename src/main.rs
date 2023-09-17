use std::io::{BufRead, BufReader, Write};
use std::net::{TcpListener, TcpStream};
use std::ptr;
use std::ptr::null;
use windows::core::{CanInto, GUID, IUnknown, Result, w};
use windows::core::PCWSTR;
use windows::Win32::System::Com::{CoCreateInstance, CoInitializeEx, CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoInitialize, CLSIDFromProgID, CLSCTX_LOCAL_SERVER, IDispatch, DISPATCH_METHOD, DISPATCH_PROPERTYGET, CLSCTX_SERVER, DISPPARAMS, CLSCTX_INPROC_SERVER, CLSCTX_REMOTE_SERVER, CLSIDFromProgIDEx, CoGetClassObject};
use windows::Win32::System::Ole::*;
use windows::Win32::System::Variant::*;

fn main() {
    unsafe {
        CoInitialize(None).unwrap();

        let clsid = CLSIDFromProgIDEx(w!("Word.Application")).unwrap();
        // let clsid = CLSIDFromProgID(w!("QBXMLRP2.RequestProcessor")).unwrap();
        // let clsid = CLSIDFromProgID(w!("WScript.Shell.1")).unwrap();

        println!("clsid: {:?}", clsid);

        // let dispatch: IUnknown = CoGetClassObject(&clsid, CLSCTX_LOCAL_SERVER|CLSCTX_INPROC_SERVER, None).unwrap();

        let dispatch: IDispatch = CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER).unwrap();

        println!("dispatch created! typeInfoCount={}", dispatch.GetTypeInfoCount().unwrap());

        // let typeInfo = dispatch.GetTypeInfo(0, 0x0400).unwrap();


        let mut dispid: i32 = 88;
        let dispid_ptr: *mut i32 = &mut dispid;
        let propname = w!("Build");
        // the first param of "GetIDsOfNames" uses something called IID_NULL
        // which turns out to be a GUID that's simply all zeroes!!!!
        let IID_NULL = GUID::default();
        dispatch.GetIDsOfNames(&IID_NULL,&propname, 1, 0x0400, dispid_ptr).unwrap();

        println!("dispid: {}", dispid);

        let mut params: DISPPARAMS = DISPPARAMS::default();
        /*params.cArgs = 0;
        params.cNamedArgs = 0;
        params.rgdispidNamedArgs = ptr::null_mut();
        params.rgvarg = ptr::null_mut();*/
        println!("dispparams: {:?}", params);

        let mut result = VARIANT::default();

        dispatch.Invoke(dispid, &IID_NULL, 0x0400, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &params,Some(&mut result), None,None).unwrap();

        println!("Invoke worked!");

        // can every variant be changed to a string?
        let mut vec: Vec<u16> = Vec::with_capacity(255);

        let pwstr = VariantToStringAlloc(&result).unwrap().to_string().unwrap();

        // VariantToString(&result, vec.as_mut_slice()).unwrap();

        //let versionFloat = VariantToDouble(&result).unwrap();

        println!("result was {:?}", pwstr);



    }

    /*unsafe {
        CoInitializeEx(ptr::null(), COINIT_APARTMENTTHREADED)
    }?;*/



    /*let listener = TcpListener::bind("127.0.0.1:7878").unwrap();

    println!("Listening on {:?}", listener.local_addr().unwrap());

    for stream in listener.incoming() {
        let stream = stream.unwrap();

        println!("Connection established! from={:?}", stream.peer_addr().unwrap());

        handle_connection(stream);
    }*/
}

fn handle_connection(mut stream: TcpStream) {
    let buf_reader = BufReader::new(&mut stream);
    let http_request: Vec<_> = buf_reader
        .lines()
        .map(|result| result.unwrap())
        .take_while(|line| !line.is_empty())
        .collect();

    println!("Request: {:#?}", http_request);

    let response = "HTTP/1.1 200 OK\r\nContent-Length: 3\r\n\r\nOK!";

    stream.write_all(response.as_bytes()).unwrap();
}
