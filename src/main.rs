mod com;

use com::*;
use std::any::{Any, TypeId};
use std::io::{BufRead, BufReader, Write};
use std::mem::ManuallyDrop;
use std::net::{TcpListener, TcpStream};
use std::{ptr, thread, time};
use std::ptr::null;
use std::time::Duration;
use windows::core::{CanInto, GUID, HSTRING, IntoParam, IUnknown, Result, w};
use windows::core::PCWSTR;
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::System::Com::{CoCreateInstance, CoInitializeEx, CLSCTX_ALL, COINIT_APARTMENTTHREADED, CoInitialize, CLSIDFromProgID, CLSCTX_LOCAL_SERVER, IDispatch, DISPATCH_METHOD, DISPATCH_PROPERTYGET, CLSCTX_SERVER, DISPPARAMS, CLSCTX_INPROC_SERVER, CLSCTX_REMOTE_SERVER, CLSIDFromProgIDEx, CoGetClassObject, DISPATCH_PROPERTYPUT};
use windows::Win32::System::Ole::*;
use windows::Win32::System::Variant::*;


// critical constant used for various com methods that turns out to be very important
static IID_NULL: GUID = GUID::zeroed();
// indicates default locale for name lookups of com methods
static DEFAULT_LOCALE_ID: u32 = 0x0400;

fn main() {
    println!("Initializing com...");
    co_initialize();

    let progId = "Word.Application".to_string();
    let clsid = clsid_from_prog_id(&progId);
    println!("Resolved progId {} -> clsid {:?}", progId, clsid);

    let dispatch = co_create_instance(&clsid);
    println!("Created dispatch");

    let visible_before = get_property(&dispatch, "Visible");
    println!("visible_before: {}", visible_before);

    put_property(&dispatch, "Visible", &Variant::from(true));
    println!("Set to visible");

    let visible_after = get_property(&dispatch, "Visible");
    println!("visible_after: {}", visible_after);

    let move_result = call_method(&dispatch, "Move", &[Variant::from(900), Variant::from(600)]);
    println!("Moved: result={}", move_result);

    thread::sleep(Duration::from_secs(5));

    let exit_result = call_method(&dispatch, "Quit", &[]);
    println!("Exit: result={}", move_result);

    println!("Done, exiting!");
}

fn main2() {
    /*set_property_any(std::ptr::null(), "test", &"test 2");
    set_property_any(std::ptr::null(), "test", &10i32);
    set_property_any(std::ptr::null(), "test", &true);*/

    unsafe {
        CoInitialize(None).unwrap();

        let clsid = CLSIDFromProgID(w!("Word.Application")).unwrap();
        // let clsid = CLSIDFromProgID(w!("QBXMLRP2.RequestProcessor")).unwrap();
        // let clsid = CLSIDFromProgID(w!("WScript.Shell.1")).unwrap();

        println!("clsid: {:?}", clsid);

        // let dispatch: IUnknown = CoGetClassObject(&clsid, CLSCTX_LOCAL_SERVER|CLSCTX_INPROC_SERVER, None).unwrap();

        let dispatch: IDispatch = CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER).unwrap();

        println!("dispatch created! typeInfoCount={}", dispatch.GetTypeInfoCount().unwrap());

        // let typeInfo = dispatch.GetTypeInfo(0, 0x0400).unwrap();


        /*let visible_var = get_property(&dispatch, "Visible");
        println!("visible: {}", debug_variant(visible_var));

        let build_var = get_property(&dispatch, "Build");
        println!("build: {}", debug_variant(build_var));

        let creator_var = get_property(&dispatch, "Creator");
        println!("creator: {}", debug_variant(creator_var));*/

        /*set_property(&dispatch, "Visible", Variant::from(true));

        // call_method(&dispatch, "Move", &[Variant::from(900i32), Variant::from(900i32)]);

        println!("Pausing 10 secs...");
        std::thread::sleep(time::Duration::from_secs(10));

        let values = &[Variant::from(100i32), Variant::from(200i32)];
        let dispid = get_ids_of_names(&dispatch, "Move");

        // setup parameters we need to pass to the com invoke
        // https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/getting-and-setting-properties
        let mut params: DISPPARAMS = DISPPARAMS::default();
        params.cNamedArgs = 0;
        //params.rgdispidNamedArgs = &mut DISPID_PROPERTYPUT;     // directly from ms c++ example
        // params.rgvarg = VARIANT::from_bool(value);

        let argsLen = values.len();
        let mut args: Vec<VARIANT> = Vec::new();
        for i in 0..argsLen {
            let mut v: &Variant = values.get(argsLen - 1 - i).unwrap();
            //args.push(v.0.clone());
            // args.push(v.0.clone());
        }

        println!("Debug");

        params.cArgs = argsLen as u32;
        params.rgvarg = args.as_mut_ptr() as *mut VARIANT;

        let mut result = VARIANT::default();

        unsafe {
            // TODO: exception handling, PUTs should never return a value i believe
            dispatch.Invoke(dispid, &IID_NULL, DEFAULT_LOCALE_ID, DISPATCH_METHOD | DISPATCH_PROPERTYPUT, &mut params, Some(&mut result), None, None).unwrap();
        }*/





       /* let d = get_ids_of_names(&dispatch, "System");

        println!("d is {}", d);



        let mut dispid: i32 = 88;
        let dispid_ptr: *mut i32 = &mut dispid;
        let propname = w!("System");
        // the first param of "GetIDsOfNames" uses something called IID_NULL
        // which turns out to be a GUID that's simply all zeroes!!!!
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

        println!("Invoke worked!");*/

        // can every variant be changed to a string?
        /*let mut vec: Vec<u16> = Vec::with_capacity(255);

        let pwstr = VariantToStringAlloc(&result).unwrap().to_string().unwrap();*/




        // VariantToString(&result, vec.as_mut_slice()).unwrap();

        //let versionFloat = VariantToDouble(&result).unwrap();

        // let v00 : VARIANT_0_0 = result.Anonymous.try_into().unwrap();

        /*let vtType = result.Anonymous.Anonymous.vt;
        println!("type was {:?}", vtType);

        if vtType == VT_I8 {
            println!("was long, value is {}", VariantToInt64(&result).unwrap());
        } else if vtType == VT_I4 {
            println!("was int, value is {}", VariantToInt32(&result).unwrap());
        } else if vtType == VT_BSTR {
            println!("was string, value is {}", VariantToStringAlloc(&result).unwrap().to_string().unwrap());
        } else if vtType == VT_DISPATCH {
            println!("was dispatch, value is {:?}", result.Anonymous.Anonymous.Anonymous.pdispVal);
            let dispatch2 = result.Anonymous.Anonymous.Anonymous.pdispVal.as_ref().unwrap();
            println!("was dispatch, value is {:?}", dispatch2);
            //println!("was dispatch!, value is {}", Variant(&result).unwrap().to_string().unwrap());
        }*/




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

/*fn get_property(dispatch: &IDispatch, name: &str) -> *mut VARIANT {
    let dispid = get_ids_of_names(dispatch, name);

    // setup parameters we need to pass to the com invoke, empty parameters should be acceptable
    let mut params: DISPPARAMS = DISPPARAMS::default();
    /*params.cArgs = 0;
    params.cNamedArgs = 0;
    params.rgdispidNamedArgs = ptr::null_mut();
    params.rgvarg = ptr::null_mut();*/
    //println!("dispparams: {:?}", params);

    let mut result = VARIANT::default();

    unsafe {
        dispatch.Invoke(dispid, &IID_NULL, DEFAULT_LOCALE_ID, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &mut params, Some(&mut result), None, None).unwrap();
    }

    return &mut result;
}*/

/*fn set_property(dispatch: *const IDispatch, name: &str, value: Variant) {
    println!("type id was {:?}", value.type_id());

    println!("was &str {}", TypeId::of::<&str>() == value.type_id());
    println!("was bool {}", TypeId::of::<bool>() == value.type_id());
    println!("was i32 {}", TypeId::of::<i32>() == value.type_id());

}*/

/*fn set_property(dispatch: &IDispatch, name: &str, mut value: Variant) {
    let dispid = get_ids_of_names(dispatch, name);

    // setup parameters we need to pass to the com invoke
    // https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/getting-and-setting-properties
    let mut params: DISPPARAMS = DISPPARAMS::default();
    params.cArgs = 1;
    params.cNamedArgs = 1;
    params.rgdispidNamedArgs = &mut DISPID_PROPERTYPUT;     // directly from ms c++ example
    // params.rgvarg = VARIANT::from_bool(value);
    params.rgvarg = &mut value.0;

    let mut result = VARIANT::default();

    unsafe {
        // TODO: exception handling, PUTs should never return a value i believe
        dispatch.Invoke(dispid, &IID_NULL, DEFAULT_LOCALE_ID, DISPATCH_METHOD | DISPATCH_PROPERTYPUT, &mut params, Some(&mut result), None, None).unwrap();
    }
}*/

/*fn set_property_ex<S: Into<String>> (dispatch: &IDispatch, name: S, value: Variant) {
    let dispid = get_ids_of_names(dispatch, name);

    // setup parameters we need to pass to the com invoke
    // https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/getting-and-setting-properties
    let mut params: DISPPARAMS = DISPPARAMS::default();
    params.cArgs = 1;
    params.cNamedArgs = 1;
    params.rgdispidNamedArgs = &mut DISPID_PROPERTYPUT;     // directly from ms c++ example
    // params.rgvarg = VARIANT::from_bool(value);
    params.rgvarg = &mut value.to_variant();

    let mut result = VARIANT::default();

    unsafe {
        // TODO: exception handling, PUTs should never return a value i believe
        dispatch.Invoke(dispid, &IID_NULL, DEFAULT_LOCALE_ID, DISPATCH_METHOD | DISPATCH_PROPERTYPUT, &mut params, Some(&mut result), None, None).unwrap();
    }
}*/



/*fn call_method(dispatch: &IDispatch, name: &str, values: &[Variant]) {
    let dispid = get_ids_of_names(dispatch, name);

    // setup parameters we need to pass to the com invoke
    // https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/getting-and-setting-properties
    let mut params: DISPPARAMS = DISPPARAMS::default();
    params.cNamedArgs = 0;
    //params.rgdispidNamedArgs = &mut DISPID_PROPERTYPUT;     // directly from ms c++ example
    // params.rgvarg = VARIANT::from_bool(value);

    let argsLen = values.len();
    let mut args: Vec<Variant> = Vec::new();
    for i in 0 .. argsLen-1 {
        let mut v: &Variant = values.get(i).unwrap();
        //args.push(v.0.clone());
        args.push(Variant::from(900));
    }

    params.cArgs = argsLen as u32;
    params.rgvarg = args.as_mut_ptr() as *mut VARIANT;

    let mut result = VARIANT::default();

    unsafe {
        // TODO: exception handling, PUTs should never return a value i believe
        dispatch.Invoke(dispid, &IID_NULL, DEFAULT_LOCALE_ID, DISPATCH_METHOD | DISPATCH_PROPERTYPUT, &mut params, Some(&mut result), None, None).unwrap();
    }
}

fn get_ids_of_names<S: Into<String>>(dispatch: *const IDispatch, name: S) -> i32 {
    let mut dispid: i32 = -1;
    // hack to get a pointer to this variable
    let dispid_ptr: *mut i32 = &mut dispid;
    // https://stackoverflow.com/questions/74173128/how-to-get-a-pcwstr-object-from-a-path-or-string
    //let propname = w!(name);
    //Convert PWSTR to PCWSTR: let s: PWSTR...; let r:PCWSTR = PCWSTR(s.0)
    //Convert String to PCWSTR: let s: String=...; let r:PCWSTR = PCWSTR(HSTRING::from(s).as_ptr())
    //Convert String-literal to PCWSTR: let r:PCWSTR = w!("example")
    let propname = PCWSTR(HSTRING::from(name.into()).as_ptr());
    unsafe {
        (*dispatch).GetIDsOfNames(&IID_NULL, &propname, 1, DEFAULT_LOCALE_ID, dispid_ptr).unwrap();
    }
    return dispid;
}*/

fn debug_variant(v: *const VARIANT) -> String {
    unsafe {
        let vt = (*v).Anonymous.Anonymous.vt;

        if vt == VT_BOOL {
            return format!("vt={:?}, type=bool, value={:?}", vt, VariantToBoolean(v).unwrap().as_bool());
        } else if vt == VT_I8 {
            return format!("vt={:?}, type=i64, value={}", vt, VariantToInt64(v).unwrap());
        } else if vt == VT_I4 {
            return format!("vt={:?}, type=i32, value={}", vt, VariantToInt32(v).unwrap());
        } else if vt == VT_BSTR {
            return format!("vt={:?}, type=string, value={}", vt, VariantToStringAlloc(v).unwrap().to_string().unwrap());
        } else if vt == VT_DISPATCH {
            // the inner dispatch is actually here
            let innerDispatch = (*v).Anonymous.Anonymous.Anonymous.pdispVal.as_ref().unwrap();
            return format!("vt={:?}, type=dispatch, value={:?}", vt, innerDispatch);
        } else {
            return format!("vt={:?}, type not handled yet", vt);
        }
    }
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
