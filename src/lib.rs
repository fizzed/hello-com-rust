use std::fmt;
use std::fmt::{Display, Formatter};
use std::result::Result;
use std::mem::ManuallyDrop;
use windows::core::{BSTR, GUID, HSTRING, Interface, PCWSTR};
use windows::Win32::Foundation::{VARIANT_BOOL};
use windows::Win32::System::Com::{CLSCTX_SERVER, CLSIDFromProgID, CoCreateInstance, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED, COINIT_SPEED_OVER_MEMORY, CoInitializeEx, DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO, IDispatch};
use windows::Win32::System::Ole::DISPID_PROPERTYPUT;
use windows::Win32::System::Variant::{VARENUM, VARIANT, VARIANT_0, VARIANT_0_0, VARIANT_0_0_0, VT_BOOL, VT_BSTR, VT_NULL, VT_DISPATCH, VT_EMPTY, VT_I1, VT_I2, VT_I4, VT_I8, VariantClear, VT_R4, VT_R8, VT_DATE};

// critical constant used for various com methods that turns out to be very important
static IID_NULL: GUID = GUID::zeroed();
// indicates default locale for name lookups of com methods
static DEFAULT_LOCALE_ID: u32 = 0x0400;

pub fn co_initialize() -> Result<(),windows::core::Error> {
    unsafe {
        // return CoInitializeEx(None, COINIT_MULTITHREADED | COINIT_SPEED_OVER_MEMORY);
        // return CoInitializeEx(None, COINIT_APARTMENTTHREADED | COINIT_SPEED_OVER_MEMORY);
        return CoInitializeEx(None, COINIT_APARTMENTTHREADED );
    }
}

pub fn clsid_from_prog_id<S: Into<String>>(prog_id: S) -> Result<GUID,windows::core::Error> {
    unsafe {
        let h_prog_id: HSTRING = HSTRING::from(prog_id.into());
        let p_prog_id: PCWSTR = PCWSTR::from_raw(h_prog_id.as_ptr());
        return CLSIDFromProgID(p_prog_id);
        // HSTRING's get dropped here, so we're all good now!
    }
}

pub fn co_create_instance(clsid: &GUID) -> IDispatch {
    unsafe {
        return CoCreateInstance(clsid, None, CLSCTX_SERVER).unwrap();
    }
}

pub fn co_create_dispatch(clsid: &GUID) -> Result<Dispatch,windows::core::Error> {
    unsafe {
        let v = CoCreateInstance(clsid, None, CLSCTX_SERVER)?;
        return Ok(Dispatch::new_with_dispatch(v));
    }
}

pub fn get_ids_of_names<S: Into<String>>(dispatch: *const IDispatch, name: S) -> Result<i32,windows::core::Error> {
    let mut dispid: i32 = -1;
    // hack to get a pointer to this variable
    let dispid_ptr: *mut i32 = &mut dispid;

    // https://stackoverflow.com/questions/74173128/how-to-get-a-pcwstr-object-from-a-path-or-string
    let h_name: HSTRING = HSTRING::from(name.into());
    let p_name: PCWSTR = PCWSTR::from_raw(h_name.as_ptr());

    unsafe {
        (*dispatch).GetIDsOfNames(&IID_NULL, &p_name, 1, DEFAULT_LOCALE_ID, dispid_ptr)?;
    }

    return Ok(dispid);
}

pub fn get_property<S: Into<String>>(dispatch: *const IDispatch, name: S) -> Result<Variant,windows::core::Error> {
    let dispid = get_ids_of_names(dispatch, name)?;

    // setup parameters we need to pass to the com invoke, empty parameters should be acceptable
    let mut params: DISPPARAMS = DISPPARAMS::default();
    let mut result = VARIANT::default();

    let wflags: DISPATCH_FLAGS = DISPATCH_METHOD | DISPATCH_PROPERTYGET;

    unsafe {
        // TODO: on exception we need to cleanup result
        (*dispatch).Invoke(dispid, &IID_NULL, DEFAULT_LOCALE_ID, wflags, &mut params, Some(&mut result), None, None)?;
    }

    // convert to our variant (and it'll VariantClear if a non-dispatch)
    return Ok(Variant::from(result));
}

pub fn put_property<S: Into<String>>(dispatch: *const IDispatch, name: S, value: &Variant) -> Result<(),windows::core::Error> {
    let dispid = get_ids_of_names(dispatch, name)?;

    // setup parameters we need to pass to the com invoke
    // https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/getting-and-setting-properties
    let mut params: DISPPARAMS = DISPPARAMS::default();
    params.cArgs = 1;
    params.cNamedArgs = 1;
    let mut dispid_named = DISPID_PROPERTYPUT;
    params.rgdispidNamedArgs = &mut dispid_named;     // directly from ms c++ example
    let mut rvv = value.to_variant();
    params.rgvarg = &mut rvv;

    let wflags: DISPATCH_FLAGS = DISPATCH_PROPERTYPUT;

    unsafe {
        // TODO: on exception we need to cleanup result
        (*dispatch).Invoke(dispid, &IID_NULL, DEFAULT_LOCALE_ID, wflags, &mut params, None, None, None)?;

        // safe to clear the variant we created in this method
        //println!("Clearing 1 VARIANT(s)");
        drop_variant_we_created(&rvv);
        VariantClear(&mut rvv).unwrap();
    }

    return Ok(())
}

pub fn call_method<S: Into<String>>(dispatch: *const IDispatch, name: S, values: &[Variant]) -> Result<Variant,windows::core::Error> {
    let dispid = get_ids_of_names(dispatch, name)?;

    // setup parameters we need to pass to the com invoke
    // https://learn.microsoft.com/en-us/previous-versions/windows/desktop/automat/getting-and-setting-properties
    let mut params: DISPPARAMS = DISPPARAMS::default();
    // no dispid for named args on calling a method
    params.cNamedArgs = 0;
    params.rgdispidNamedArgs = std::ptr::null_mut();

    // build array of variant arguments in reverse order (no idea why the COM api wants them reversed)
    let args_len: usize = values.len();
    let mut args: Vec<VARIANT> = Vec::new();
    for i in 0..args_len {
        // we build the variant arguments in reverse order
        let v: &Variant = values.get(args_len - 1 - i).unwrap();
        args.push(v.to_variant());
    }

    params.cArgs = args_len as u32;
    params.rgvarg = args.as_mut_ptr() as *mut VARIANT;

    let wflags: DISPATCH_FLAGS = DISPATCH_METHOD | DISPATCH_PROPERTYGET;
    let mut result = VARIANT::default();
    let mut except_info: EXCEPINFO = EXCEPINFO::default();

    unsafe {
        // TODO: on exception we need to cleanup result
        (*dispatch).Invoke(dispid, &IID_NULL, DEFAULT_LOCALE_ID, wflags, &mut params, Some(&mut result), Some(&mut except_info), None)?;

        // safe to clear the variant(s) we created in this method
        //println!("Clearing {} VARIANT(s)", args_len);
        while !args.is_empty() {
            let mut rvv = args.pop().unwrap();
            drop_variant_we_created(&rvv);
            VariantClear(&mut rvv).unwrap();
        }
    }

    return Ok(Variant::from(result));
}

#[derive(Debug)]
pub struct Error {
    message: String
}

impl Error {
    pub fn result<S: Into<String>>(message: S) -> Error {
        Error {
            message: message.into()
        }
    }
}

impl std::error::Error for Error { }

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "{}", self.message)
    }
}

pub struct Dispatch {
    dispatch: Option<IDispatch>,
    variant: Option<VARIANT>               // for some types like dispatch where we need to keep a reference to the original VARIANT
}

impl Display for Dispatch {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "vt_dispatch={:?}", self.get_dispatch().as_raw())
    }
}

// useful for debugging when dispatches are dropped
/*impl Drop for Dispatch {
    fn drop(&mut self) {
        println!("dropping {}", self);
    }
}*/

impl Dispatch {
    fn new_with_dispatch(dispatch: IDispatch) -> Dispatch {
        Dispatch {
            dispatch: Some(dispatch),
            variant: None
        }
    }

    fn new_with_variant(variant: VARIANT) -> Dispatch {
        Dispatch {
            dispatch: None,
            variant: Some(variant)
        }
    }

    fn get_dispatch(&self) -> &IDispatch {
        if self.dispatch.is_some() {
            return self.dispatch.as_ref().unwrap();
        } else {
            unsafe {
                return self.variant.as_ref().unwrap().Anonymous.Anonymous.Anonymous.pdispVal.as_ref().unwrap();
            }
        }
    }

    pub fn get_property<S: Into<String>>(&self, name: S) -> Result<Variant,windows::core::Error> {
        let dispatch = self.get_dispatch();
        return get_property(dispatch, name);
    }

    pub fn put_property<S: Into<String>>(&self, name: S, value: &Variant) -> Result<(),windows::core::Error> {
        let dispatch = self.get_dispatch();
        return put_property(dispatch, name, value);
    }

    pub fn call_method<S: Into<String>>(&self, name: S, values: &[Variant]) -> Result<Variant,windows::core::Error> {
        let dispatch = self.get_dispatch();
        return call_method(dispatch, name, values);
    }

}


union UnionedValue {
    bool_val: bool,                         // VT_BOOL
    u8_val: u8,                             // VT_UI1, bVal  (VT_I1 also is a u8 but is in the cVal)
    i16_val: i16,                           // VT_I2, iVal
    i32_val: i32,                           // VT_I4, lVal
    i64_val: i64,                           // VT_I8, llVal
    f32_val: f32,
    f64_val: f64
}

impl UnionedValue {
    pub fn default() -> UnionedValue {
        UnionedValue {
            u8_val: 0
        }
    }
}

pub struct Variant {
    vt: VARENUM,
    str: Option<String>,
    unioned: UnionedValue,
    variant: Option<VARIANT>               // for some types like dispatch where we need to keep a reference to the original VARIANT
}

impl Variant {
    pub fn empty() -> Variant {
        Variant {
            vt: VT_EMPTY,
            str: None,
            unioned: UnionedValue::default(),
            variant: None
        }
    }

    fn new_with_string(str: String) -> Variant {
        Variant {
            vt: VT_BSTR,
            str: Some(str),
            unioned: UnionedValue::default(),
            variant: None
        }
    }

    fn new_with_variant(variant: VARIANT) -> Variant {
        unsafe {
            Variant {
                vt: variant.Anonymous.Anonymous.vt,
                str: None,
                unioned: UnionedValue::default(),
                variant: Some(variant)
            }
        }
    }

    fn new_with_unioned(vt: VARENUM, value: UnionedValue) -> Variant {
        Variant {
            vt,
            str: None,
            unioned: value,
            variant: None
        }
    }

    pub fn to_dispatch(&self) -> Result<Dispatch,Error> {
        if self.vt != VT_DISPATCH {
            return Err(Error::result("variant is not a dispatch"));
        }
        // TODO: how can we move ownership of the variant to the dispatch?
        let variant = self.variant.to_owned().unwrap();
        let dispatch = Dispatch::new_with_variant(variant);
        return Ok(dispatch);
    }

    pub fn get_raw_idispatch(&self) -> *const IDispatch {
        unsafe {
            return self.variant.as_ref().unwrap().Anonymous.Anonymous.Anonymous.pdispVal.as_ref().unwrap();
        }
    }

    pub fn to_i32(&self) -> Result<i32,Error> {
        unsafe {
            match self.vt {
                VT_I1 => Ok(self.unioned.u8_val as i32),
                VT_I2 => Ok(self.unioned.i16_val as i32),
                VT_I4 => Ok(self.unioned.i32_val),
                VT_I8 => Ok(self.unioned.i16_val as i32),
                _ => Err(Error::result("variant is not a numeric type convertible to i32"))
            }
        }
    }

    pub fn to_variant(&self) -> VARIANT {
        unsafe {
            // generate new contents based on type
            let contents = match self.vt {
                VT_BOOL => VARIANT_0_0_0 { boolVal: VARIANT_BOOL::from(self.unioned.bool_val) },
                VT_BSTR => VARIANT_0_0_0 { bstrVal: ManuallyDrop::new(BSTR::from(self.str.as_ref().unwrap())) },
                VT_I1 => VARIANT_0_0_0 { bVal: self.unioned.u8_val },
                VT_I2 => VARIANT_0_0_0 { iVal: self.unioned.i16_val },
                VT_I4 => VARIANT_0_0_0 { lVal: self.unioned.i32_val },
                VT_I8 => VARIANT_0_0_0 { llVal: self.unioned.i64_val },
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
    fn from(mut value: VARIANT) -> Variant {
        unsafe {
            let vt = value.Anonymous.Anonymous.vt;
            let holder = &value.Anonymous.Anonymous.Anonymous;
            if vt == VT_DISPATCH {
                Variant::new_with_variant(value)
            } else {
                let rv = match vt {
                    VT_EMPTY | VT_NULL => Variant::empty(),
                    VT_BOOL => Variant::from(holder.boolVal.as_bool()),
                    VT_BSTR => Variant::from(holder.bstrVal.to_string()),
                    VT_I1 => Variant::from(holder.bVal),
                    VT_I2 => Variant::from(holder.iVal),
                    VT_I4 => Variant::from(holder.lVal),
                    VT_I8 => Variant::from(holder.llVal),
                    VT_R4 => Variant::from(holder.fltVal),
                    VT_R8 => Variant::from(holder.dblVal),
                    // TODO: convert to epoch millis
                    VT_DATE => {
                        let mut v = Variant::from(holder.date);
                        v.vt = VT_DATE;
                        return v;
                    },
                    _ => panic!("unsupported type {:?}", vt)
                };
                // safe to clear the variant now
                VariantClear(&mut value).unwrap();
                rv
            }

        }
    }
}

// useful for debugging when dispatches are dropped
/*impl Drop for Variant {
    fn drop(&mut self) {
        println!("dropping {}", self);
    }
}*/

fn drop_variant_we_created(variant: &VARIANT) {
    unsafe {
        match variant.Anonymous.Anonymous.vt {
            VT_BSTR => {
                drop(&variant.Anonymous.Anonymous.Anonymous.bstrVal);
            }
            _ => {}
        }
        drop(&variant.Anonymous.Anonymous)
    }
}

impl fmt::Debug for Variant {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        unsafe {
            let s = match self.vt {
                VT_EMPTY => format!("<vt_empty>"),
                VT_NULL => format!("<vt_null>"),
                VT_DISPATCH => format!("(vt_dispatch {:?})", self.get_raw_idispatch()),
                VT_BOOL => format!("(vt_bool {})", self.unioned.bool_val),
                VT_BSTR => format!("(vt_bstr {})", self.str.as_ref().unwrap()),
                VT_I1 => format!("(vt_i1 {})", self.unioned.u8_val),
                VT_I2 => format!("(vt_i2 {})", self.unioned.i16_val),
                VT_I4 => format!("(vt_i4 {})", self.unioned.i32_val),
                VT_I8 => format!("(vt_i8 {})", self.unioned.i64_val),
                VT_R4 => format!("(vt_r4 {})", self.unioned.f32_val),
                VT_R8 => format!("(vt_r8 {})", self.unioned.f64_val),
                // TODO: use epoch millis maybe?
                VT_DATE => format!("(vt_date {})", self.unioned.f64_val),
                _ => todo!()
            };
            write!(f, "{}", s)
        }
    }
}

impl fmt::Display for Variant {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        unsafe {
            let s = match self.vt {
                VT_EMPTY => format!("<empty>"),
                VT_NULL => format!("<null>"),
                VT_DISPATCH => format!("<dispatch {:?}>", self.get_raw_idispatch()),
                VT_BOOL => format!("{}", self.unioned.bool_val),
                VT_BSTR => format!("{}", self.str.as_ref().unwrap()),
                VT_I1 => format!("{}", self.unioned.u8_val),
                VT_I2 => format!("{}", self.unioned.i16_val),
                VT_I4 => format!("{}", self.unioned.i32_val),
                VT_I8 => format!("{}", self.unioned.i64_val),
                VT_R4 => format!("{}", self.unioned.f32_val),
                VT_R8 => format!("{}", self.unioned.f64_val),
                VT_DATE => format!("{}", self.unioned.f64_val),
                _ => todo!()
            };
            write!(f, "{}", s)
        }
    }
}

impl From<bool> for Variant {
    fn from(value: bool) -> Variant {
        Variant::new_with_unioned(VT_BOOL, UnionedValue { bool_val: value })
    }
}

impl From<&str> for Variant {
    fn from(value: &str) -> Variant {
        Variant::new_with_string(value.to_string())
    }
}

impl From<String> for Variant {
    fn from(value: String) -> Variant {
        Variant::new_with_string(value)
    }
}

impl From<u8> for Variant {
    fn from(value: u8) -> Variant {
        Variant::new_with_unioned(VT_I1, UnionedValue { u8_val: value })
    }
}

impl From<i16> for Variant {
    fn from(value: i16) -> Variant {
        Variant::new_with_unioned(VT_I2, UnionedValue { i16_val: value })
    }
}

impl From<i32> for Variant {
    fn from(value: i32) -> Variant {
        Variant::new_with_unioned(VT_I4, UnionedValue { i32_val: value })
    }
}

impl From<i64> for Variant {
    fn from(value: i64) -> Variant {
        Variant::new_with_unioned(VT_I8, UnionedValue { i64_val: value })
    }
}

impl From<f32> for Variant {
    fn from(value: f32) -> Variant {
        Variant::new_with_unioned(VT_R4, UnionedValue { f32_val: value })
    }
}

impl From<f64> for Variant {
    fn from(value: f64) -> Variant {
        Variant::new_with_unioned(VT_R8, UnionedValue { f64_val: value })
    }
}