use std::any::Any;
use std::fmt;
use std::mem::ManuallyDrop;
use windows::core::{BSTR, GUID, HSTRING, PCWSTR, w};
use windows::Win32::Foundation::{VARIANT_BOOL};
use windows::Win32::System::Com::{CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER, CLSIDFromProgID, CoCreateInstance, CoInitialize, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, IDispatch};
use windows::Win32::System::Ole::DISPID_PROPERTYPUT;
use windows::Win32::System::Variant::{VARENUM, VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0, VT_BOOL, VT_BSTR, VT_EMPTY, VT_I1, VT_I2, VT_I4, VT_I8};

// critical constant used for various com methods that turns out to be very important
static IID_NULL: GUID = GUID::zeroed();
// indicates default locale for name lookups of com methods
static DEFAULT_LOCALE_ID: u32 = 0x0400;

pub fn co_initialize() {
    unsafe {
        // TODO: how do you bubble up an error?
        CoInitialize(None).unwrap();
    }
}

pub fn clsid_from_prog_id<S: Into<String>>(prog_id: S) -> GUID {
    unsafe {
        let h_prog_id: HSTRING = HSTRING::from(prog_id.into());
        let p_prog_id: PCWSTR = PCWSTR::from_raw(h_prog_id.as_ptr());
        return CLSIDFromProgID(p_prog_id).unwrap();
        // HSTRING's get dropped here, so we're all good now!
    }
}

pub fn co_create_instance(clsid: &GUID) -> IDispatch {
    unsafe {
        return CoCreateInstance(clsid, None, CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER).unwrap();
    }
}

pub fn get_ids_of_names<S: Into<String>>(dispatch: *const IDispatch, name: S) -> i32 {
    let mut dispid: i32 = -1;
    // hack to get a pointer to this variable
    let dispid_ptr: *mut i32 = &mut dispid;
    // https://stackoverflow.com/questions/74173128/how-to-get-a-pcwstr-object-from-a-path-or-string
    let h_name: HSTRING = HSTRING::from(name.into());
    let p_name: PCWSTR = PCWSTR::from_raw(h_name.as_ptr());
    unsafe {
        (*dispatch).GetIDsOfNames(&IID_NULL, &p_name, 1, DEFAULT_LOCALE_ID, dispid_ptr).unwrap();
    }
    return dispid;
}

pub fn get_property<S: Into<String>>(dispatch: &IDispatch, name: S) -> Variant {
    let dispid = get_ids_of_names(dispatch, name);

    // setup parameters we need to pass to the com invoke, empty parameters should be acceptable
    let mut params: DISPPARAMS = DISPPARAMS::default();
    let mut result = VARIANT::default();

    unsafe {
        dispatch.Invoke(dispid, &IID_NULL, DEFAULT_LOCALE_ID, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &mut params, Some(&mut result), None, None).unwrap();
    }

    // now we buid a Variant to return

    return Variant::from(result);
}

pub fn put_property<S: Into<String>>(dispatch: &IDispatch, name: S, value: &Variant) {
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
}

pub fn call_method<S: Into<String>>(dispatch: &IDispatch, name: S, values: &[Variant]) -> Variant {
    let dispid = get_ids_of_names(dispatch, name);

    // setup parameters we need to pass to the com invoke
    // https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/getting-and-setting-properties
    let mut params: DISPPARAMS = DISPPARAMS::default();
    params.cNamedArgs = 0;
    //params.rgdispidNamedArgs = &mut DISPID_PROPERTYPUT;     // directly from ms c++ example
    // params.rgvarg = VARIANT::from_bool(value);

    let args_len: usize = values.len();
    let mut args: Vec<VARIANT> = Vec::new();
    for i in 0..args_len {
        let v: &Variant = values.get(args_len - 1 - i).unwrap();
        args.push(v.to_variant());
    }

    params.cArgs = args_len as u32;
    params.rgvarg = args.as_mut_ptr() as *mut VARIANT;

    let mut result = VARIANT::default();

    unsafe {
        // TODO: exception handling, PUTs should never return a value i believe
        dispatch.Invoke(dispid, &IID_NULL, DEFAULT_LOCALE_ID, DISPATCH_METHOD | DISPATCH_PROPERTYPUT, &mut params, Some(&mut result), None, None).unwrap();
    }

    return Variant::from(result);
}




union ValueUnion {
    bool_val: bool,                         // VT_BOOL
    u8_val: u8,                             // VT_UI1, bVal  (VT_I1 also is a u8 but is in the cVal)
    i16_val: i16,                           // VT_I2, iVal
    i32_val: i32,                           // VT_I4, lVal
    i64_val: i64                            // VT_I8, llVal
}

pub struct Variant {
    vt: VARENUM,
    string_val: Option<String>,             // VT_BSTR, do not include the string in the ComValue
    val: ValueUnion
}

impl Variant {
    pub fn empty() -> Variant {
        Variant {
            vt: VT_EMPTY,
            string_val: None,
            val: ValueUnion { i32_val: 0 }
        }
    }

    pub fn to_variant(&self) -> VARIANT {
        unsafe {
            let contents = match self.vt {
                VT_BOOL => VARIANT_0_0_0 { boolVal: VARIANT_BOOL::from(self.val.bool_val) },
                VT_BSTR => VARIANT_0_0_0 { bstrVal: ManuallyDrop::new(BSTR::from(self.string_val.as_ref().unwrap())) },
                VT_I1 => VARIANT_0_0_0 { bVal: self.val.u8_val },
                VT_I2 => VARIANT_0_0_0 { iVal: self.val.i16_val },
                VT_I4 => VARIANT_0_0_0 { lVal: self.val.i32_val },
                VT_I8 => VARIANT_0_0_0 { llVal: self.val.i64_val },
                _ => todo!()
            };

            VARIANT {
                Anonymous: VARIANT_0 {
                    Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                        vt: self.vt,
                        wReserved1: 0,
                        wReserved2: 0,
                        wReserved3: 0,
                        Anonymous: contents,
                    }),
                }
            }
        }
    }
}

impl From<VARIANT> for Variant {
    fn from(value: VARIANT) -> Variant {
        unsafe {
            let vt = value.Anonymous.Anonymous.vt;
            println!("Trying to build Variant from vt {:?}", vt);
            let holder = &value.Anonymous.Anonymous.Anonymous;
            let rv = match vt {
                VT_EMPTY => Variant::empty(),
                VT_BOOL => Variant::from(holder.boolVal.as_bool()),
                // VT_BSTR => VARIANT_0_0_0 { bstrVal: ManuallyDrop::new(BSTR::from(self.string_val.as_ref().unwrap())) },
                // VT_I1 => VARIANT_0_0_0 { bVal: self.val.u8_val },
                // VT_I2 => VARIANT_0_0_0 { iVal: self.val.i16_val },
                // VT_I4 => VARIANT_0_0_0 { lVal: self.val.i32_val },
                // VT_I8 => VARIANT_0_0_0 { llVal: self.val.i64_val },
                _ => todo!()
            };

            return rv;
        }
    }
}

impl Drop for Variant {
    fn drop(&mut self) {
        println!("Dropping variant");
        /*match VARENUM(unsafe { self.0.Anonymous.Anonymous.vt.0 }) {
            VT_BSTR => unsafe {
                drop(&mut &self.0.Anonymous.Anonymous.Anonymous.bstrVal)
            }
            _ => {}
        }
        unsafe { drop(&mut self.0.Anonymous.Anonymous) }*/
    }
}

impl fmt::Display for Variant {
    // This trait requires `fmt` with this exact signature.
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        unsafe {
            let s = match self.vt {
                VT_EMPTY => format!("vt=VT_EMPTY, value=nil"),
                VT_BOOL => format!("vt=VT_BOOL, value={}", self.val.bool_val),
                VT_BSTR => format!("vt=VT_BSTR, value={}", self.string_val.as_ref().unwrap()),
                VT_I1 => format!("vt=VT_I1, value={}", self.val.u8_val),
                VT_I2 => format!("vt=VT_I2, value={}", self.val.i16_val),
                VT_I4 => format!("vt=VT_I4, value={}", self.val.i32_val),
                VT_I8 => format!("vt=VT_I8, value={}", self.val.i64_val),
                _ => todo!()
            };
            write!(f, "{}", s)
        }
    }
}

impl From<bool> for Variant {
    fn from(value: bool) -> Variant {
        Variant {
            vt: VT_BOOL,
            string_val: None,
            val: ValueUnion {
                bool_val: value
            }
        }
    }
}

impl From<&str> for Variant {
    fn from(value: &str) -> Variant {
        Variant {
            vt: VT_BSTR,
            string_val: Option::from(value.to_string()),
            val: ValueUnion { i16_val: 0 },
        }
    }
}

impl From<String> for Variant {
    fn from(value: String) -> Variant {
        Variant {
            vt: VT_BSTR,
            string_val: Option::from(value),
            val: ValueUnion { i16_val: 0 }
        }
    }
}

impl From<u8> for Variant {
    fn from(value: u8) -> Variant {
        Variant {
            vt: VT_I1,
            string_val: None,
            val: ValueUnion {
                u8_val: value
            }
        }
    }
}

impl From<i16> for Variant {
    fn from(value: i16) -> Variant {
        Variant {
            vt: VT_I2,
            string_val: None,
            val: ValueUnion {
                i16_val: value
            }
        }
    }
}

impl From<i32> for Variant {
    fn from(value: i32) -> Variant {
        Variant {
            vt: VT_I4,
            string_val: None,
            val: ValueUnion {
                i32_val: value
            }
        }
    }
}

impl From<i64> for Variant {
    fn from(value: i64) -> Variant {
        Variant {
            vt: VT_I8,
            string_val: None,
            val: ValueUnion {
                i64_val: value
            }
        }
    }
}




/*pub struct Variant(pub(crate) VARIANT);
impl Variant {
    pub fn new(num: VARENUM, contents: VARIANT_0_0_0) -> Variant {
        Variant {
            0: VARIANT {
                Anonymous: VARIANT_0 {
                    Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                        vt: num,
                        wReserved1: 0,
                        wReserved2: 0,
                        wReserved3: 0,
                        Anonymous: contents,
                    }),
                },
            },
        }
    }

    pub fn as_mut_variant(&mut self) -> *mut VARIANT {
        return &mut self.0;
    }
}

impl From<String> for Variant {
    fn from(value: String) -> Variant { Variant::new(VT_BSTR, VARIANT_0_0_0 { bstrVal: ManuallyDrop::new(BSTR::from(value)) }) }
}
impl From<&str> for Variant {
    fn from(value: &str) -> Variant { Variant::from(value.to_string()) }
}
impl From<bool> for Variant {
    fn from(value: bool) -> Variant { Variant::new(VT_BOOL, VARIANT_0_0_0 { boolVal: VARIANT_BOOL::from(value) }) }
}
impl From<i32> for Variant {
    fn from(value: i32) -> Variant { Variant::new(VT_I4, VARIANT_0_0_0 { lVal: value }) }
}

impl Drop for Variant {
    fn drop(&mut self) {
        println!("Dropping Variant!");
        match VARENUM(unsafe { self.0.Anonymous.Anonymous.vt.0 }) {
            VT_BSTR => unsafe {
                drop(&mut &self.0.Anonymous.Anonymous.Anonymous.bstrVal)
            }
            _ => {}
        }
        unsafe { drop(&mut self.0.Anonymous.Anonymous) }
    }
}*/
