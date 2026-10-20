//! COM 值类型：`ComValue`（传给 COM）与 `WordValue`（从 COM 读出）。

use std::mem::ManuallyDrop;

use windows::core::BSTR;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::{
    VARIANT, VT_BOOL, VT_BSTR, VT_DATE, VT_DISPATCH, VT_EMPTY, VT_ERROR, VT_I2, VT_I4, VT_I8,
    VT_NULL, VT_R4, VT_R8,
};

/// 传给 COM 的参数值。实现常用类型的 `From`，调用处直接 `"文本".into()` / `12.into()`。
#[derive(Clone)]
pub enum ComValue {
    Str(String),
    Bool(bool),
    Int(i32),
    Float(f64),
    /// VT_EMPTY：空值（多数可选参数收它就行）。
    Empty,
    /// 真正「省略该可选参数」：VT_ERROR + DISP_E_PARAMNOTFOUND（VBA 里 `, ,` 那个空档）。
    /// 与 `Empty`（空**值**）不同，引擎不会去解析它。
    Missing,
    /// COM 对象参数（如把 Range/Cell 传给 `Merge`、`Add`）。
    Obj(IDispatch),
}

impl std::fmt::Debug for ComValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Str(s) => write!(f, "Str({s:?})"),
            Self::Bool(b) => write!(f, "Bool({b})"),
            Self::Int(i) => write!(f, "Int({i})"),
            Self::Float(v) => write!(f, "Float({v})"),
            Self::Empty => write!(f, "Empty"),
            Self::Missing => write!(f, "Missing"),
            Self::Obj(_) => write!(f, "Obj(<IDispatch>)"),
        }
    }
}

impl From<&str> for ComValue {
    fn from(s: &str) -> Self {
        Self::Str(s.to_string())
    }
}
impl From<String> for ComValue {
    fn from(s: String) -> Self {
        Self::Str(s)
    }
}
impl From<bool> for ComValue {
    fn from(b: bool) -> Self {
        Self::Bool(b)
    }
}
impl From<i32> for ComValue {
    fn from(i: i32) -> Self {
        Self::Int(i)
    }
}
impl From<u32> for ComValue {
    fn from(i: u32) -> Self {
        Self::Int(i as i32)
    }
}
impl From<f64> for ComValue {
    fn from(f: f64) -> Self {
        Self::Float(f)
    }
}
impl From<f32> for ComValue {
    fn from(f: f32) -> Self {
        Self::Float(f as f64)
    }
}
impl From<IDispatch> for ComValue {
    fn from(d: IDispatch) -> Self {
        Self::Obj(d)
    }
}

impl From<ComValue> for VARIANT {
    fn from(v: ComValue) -> VARIANT {
        match v {
            ComValue::Str(s) => VARIANT::from(BSTR::from(s.as_str())),
            ComValue::Bool(b) => VARIANT::from(b),
            ComValue::Int(i) => VARIANT::from(i),
            ComValue::Float(f) => VARIANT::from(f),
            ComValue::Empty => VARIANT::default(),
            ComValue::Missing => {
                // DISP_E_PARAMNOTFOUND：可选参数「省略」的官方编码。
                let mut var = VARIANT::default();
                unsafe {
                    (*var.Anonymous.Anonymous).vt = VT_ERROR;
                    (*var.Anonymous.Anonymous).Anonymous.scode = 0x8002_0004u32 as i32;
                }
                var
            }
            ComValue::Obj(d) => {
                let mut var = VARIANT::default();
                unsafe {
                    (*var.Anonymous.Anonymous).vt = VT_DISPATCH;
                    (*var.Anonymous.Anonymous).Anonymous.pdispVal = ManuallyDrop::new(Some(d));
                }
                var
            }
        }
    }
}

/// 从 COM 读出的值。
#[derive(Debug, Clone, PartialEq)]
pub enum WordValue {
    Empty,
    Text(String),
    Number(f64),
    Bool(bool),
    Error(String),
}

impl std::fmt::Display for WordValue {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Empty => Ok(()),
            Self::Text(s) => f.write_str(s),
            Self::Number(n) => {
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    write!(f, "{}", *n as i64)
                } else {
                    write!(f, "{n}")
                }
            }
            Self::Bool(b) => write!(f, "{b}"),
            Self::Error(e) => write!(f, "#{e}"),
        }
    }
}

fn error_text(scode: u32) -> String {
    match scode & 0xFFFF {
        2000 => "NULL!".into(),
        2007 => "DIV/0!".into(),
        2015 => "VALUE!".into(),
        2023 => "REF!".into(),
        2029 => "NAME?".into(),
        2036 => "NUM!".into(),
        2042 => "N/A".into(),
        _ => format!("0x{scode:08X}"),
    }
}

/// VARIANT → WordValue（严格按 vt 分派，杜绝直读 union 拿到垃圾值）。
pub(crate) fn variant_to_wordvalue(v: &VARIANT) -> WordValue {
    let vt: u16 = unsafe { v.Anonymous.Anonymous.vt }.0;
    unsafe {
        let u = &v.Anonymous.Anonymous.Anonymous;
        if vt == VT_EMPTY.0 || vt == VT_NULL.0 {
            WordValue::Empty
        } else if vt == VT_BSTR.0 {
            WordValue::Text(u.bstrVal.to_string())
        } else if vt == VT_I4.0 {
            WordValue::Number(u.lVal as f64)
        } else if vt == VT_I2.0 {
            WordValue::Number(u.iVal as f64)
        } else if vt == VT_I8.0 {
            WordValue::Number(u.llVal as f64)
        } else if vt == VT_R4.0 {
            WordValue::Number(u.fltVal as f64)
        } else if vt == VT_R8.0 || vt == VT_DATE.0 {
            WordValue::Number(u.dblVal)
        } else if vt == VT_BOOL.0 {
            WordValue::Bool(bool::from(u.boolVal))
        } else if vt == VT_ERROR.0 {
            WordValue::Error(error_text(u.scode as u32))
        } else {
            WordValue::Text(format!("<未支持的 VARIANT 类型 0x{vt:04X}>"))
        }
    }
}
