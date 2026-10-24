//! IDispatch 晚绑定薄封装 —— 整个库的 COM 底座。
//!
//! 设计要点：
//! - 所有对象包装（App/Document/Range/...）内部持有一个 `IDispatchWrapper`；
//! - `get`/`call` 自动在 PROPERTYGET 与 METHOD 间回退，兼容 WPS 把属性暴露成方法等怪癖；
//! - 每个包装自带对象路径 `ctx`，错误自动带完整定位（见 [crate::error::WordError]）；
//! - **扩展性核心**：库未封装的成员用 `.raw().get/call/call_named/put` 一行直调。

use anyhow::Result;
use windows::core::{BSTR, GUID, PCWSTR};
use windows::Win32::Globalization::LOCALE_USER_DEFAULT;
use windows::Win32::System::Com::{
    IDispatch, DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT,
    DISPPARAMS, EXCEPINFO,
};
use windows::Win32::System::Ole::DISPID_PROPERTYPUT;
use windows::Win32::System::Variant::{VARIANT, VT_DISPATCH};

use crate::com::variant::{variant_to_wordvalue, ComValue, WordValue};
use crate::error::WordError;

const DISP_E_MEMBERNOTFOUND: u32 = 0x8002_0003;

/// 该 HRESULT 是否属于「应用不认这组具名参数」——WPS 的 `Invoke` 对 `rgdispidNamedArgs`
/// 支持不全，常见返回 0x0602xxxx「无此命令」。
fn rejects_named_args(hr: u32) -> bool {
    (hr & 0xFFFF_0000) == 0x0602_0000
}

/// 常见 HRESULT → 修复提示。
fn common_hint(hr: u32) -> &'static str {
    if rejects_named_args(hr) {
        return "应用不认这组参数/成员（WPS 常见于具名参数）：改用位置参数或核对成员名";
    }
    match hr {
        0x8002_0003 => "成员不存在：检查成员名拼写（可用 raw 命令探测）",
        0x8002_0005 => "类型不匹配：检查参数类型",
        0x8002_0009 => "COM 调用抛出异常：常见于操作被应用弹窗取消、序号越界或参数不被接受",
        0x800A_03EC => "应用层错误：常见于非法 Range、越界、对象不存在",
        0x800A_01A8 => "对象下标越界：检查集合索引或名称",
        _ => "",
    }
}

/// COM 错误 → 带上下文的 WordError。
pub(crate) fn com_err(kind: &str, member: &str, ctx: &str, e: windows::core::Error) -> WordError {
    let hr = e.code().0 as u32;
    WordError::new(
        format!("{kind}('{member}')"),
        ctx,
        format!("0x{hr:08X} {e}"),
    )
    .with_hint(common_hint(hr))
}

/// 错误是否为 DISP_E_MEMBERNOTFOUND（成员不存在，用于 SaveAs2→SaveAs 等回退）。
pub(crate) fn is_member_notfound(e: &anyhow::Error) -> bool {
    e.to_string().contains("0x80020003")
}

/// 纯逻辑：按 COM `DISPPARAMS` 规则排序参数 —— **具名在前（保持原顺序），位置参数逆序在后**。
/// 与 COM 无关，可单测（见本文件末尾测试）。
pub(crate) fn order_args(pos: Vec<ComValue>, named: Vec<ComValue>) -> Vec<ComValue> {
    let mut out = named;
    let mut p = pos;
    p.reverse();
    out.extend(p);
    out
}

/// BSTR → String，空指针按空串处理。
fn bstr_text(value: &BSTR) -> String {
    if value.is_empty() {
        String::new()
    } else {
        value.to_string()
    }
}

/// 调用失败时把应用抛出的 `EXCEPINFO` 拼进错误信息。
///
/// `0x80020009 发生意外` 只是外壳，真正原因（参数个数不对、插入位置无效、表格嵌套非法…）
/// 在 `bstrSource` / `bstrDescription` 里。吞掉它等于让调用方对着「发生意外」猜。
fn enrich(e: windows::core::Error, excep: &EXCEPINFO) -> windows::core::Error {
    let desc = bstr_text(&excep.bstrDescription);
    if desc.is_empty() {
        return e;
    }
    let source = bstr_text(&excep.bstrSource);
    let message = if source.is_empty() {
        desc
    } else {
        format!("{source}: {desc}")
    };
    windows::core::Error::new(e.code(), message)
}

/// 包住一个 `IDispatch` 指针 + 对象路径。
#[derive(Clone)]
pub struct IDispatchWrapper {
    pub disp: IDispatch,
    /// 对象路径，如 `WPS > report.docx > Range[120,128]`，错误定位用。
    pub ctx: String,
}

impl IDispatchWrapper {
    /// 根对象（Application）。`engine_label` 如 "WPS Writer"。
    pub fn root(disp: IDispatch, engine_label: &str) -> Self {
        Self {
            disp,
            ctx: engine_label.to_string(),
        }
    }

    /// 子对象：继承父路径并追加自身标签。
    pub fn child(&self, disp: IDispatch, label: &str) -> Self {
        Self {
            disp,
            ctx: format!("{} > {}", self.ctx, label),
        }
    }

    /// 同一对象、更深一层的上下文（如 Range → Font）。
    pub fn deeper(&self, label: &str) -> Self {
        Self {
            disp: self.disp.clone(),
            ctx: format!("{} > {}", self.ctx, label),
        }
    }

    /// get 属性并转为子对象（ctx 自动追加 label）。
    pub fn get_obj(&self, name: &str, label: &str) -> Result<IDispatchWrapper> {
        self.get(name)?.idispatch_child(label)
    }

    /// call 方法并转为子对象（ctx 自动追加 label）。
    pub fn call_obj(
        &self,
        name: &str,
        label: &str,
        args: Vec<ComValue>,
    ) -> Result<IDispatchWrapper> {
        self.call(name, args)?.idispatch_child(label)
    }

    /// 成员是否存在（`GetIDsOfNames` 探测，不调用）。用于 capability 探测。
    pub fn has_member(&self, name: &str) -> bool {
        self.get_dispid(name).is_ok()
    }

    fn get_dispid(&self, name: &str) -> windows::core::Result<i32> {
        let bstr = BSTR::from(name);
        let names = [PCWSTR(bstr.as_ptr())];
        let mut dispid: i32 = 0;
        unsafe {
            self.disp.GetIDsOfNames(
                &GUID::default(),
                names.as_ptr(),
                names.len() as u32,
                LOCALE_USER_DEFAULT,
                &mut dispid,
            )?;
        }
        Ok(dispid)
    }

    unsafe fn invoke_raw(
        &self,
        dispid: i32,
        flags: DISPATCH_FLAGS,
        params: &DISPPARAMS,
    ) -> windows::core::Result<VARIANT> {
        let mut result = VARIANT::default();
        // EXCEPINFO 一定要传实地址：传 null 时应用不会填，错误就只剩「发生意外」四个字。
        let mut excep: EXCEPINFO = std::mem::zeroed();
        let mut arg_err: u32 = 0;
        let invoked = self.disp.Invoke(
            dispid,
            &GUID::default(),
            LOCALE_USER_DEFAULT,
            flags,
            params,
            Some(&mut result),
            Some(&mut excep),
            Some(&mut arg_err),
        );
        match invoked {
            Ok(()) => Ok(result),
            Err(e) => Err(enrich(e, &excep)),
        }
    }

    /// 读取属性，如 `get("Name")` / `get("Content")`。
    /// WPS 个别成员是「方法」形式：属性读取报 MEMBERNOTFOUND 时自动退化为 METHOD 再试。
    pub fn get(&self, name: &str) -> Result<DispatchResult> {
        let dispid = self
            .get_dispid(name)
            .map_err(|e| com_err("dispid", name, &self.ctx, e))?;
        let params = DISPPARAMS {
            rgvarg: std::ptr::null_mut(),
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: 0,
            cNamedArgs: 0,
        };
        match unsafe { self.invoke_raw(dispid, DISPATCH_PROPERTYGET, &params) } {
            Ok(v) => Ok(DispatchResult::new(v, &self.ctx)),
            Err(e) if e.code().0 as u32 == DISP_E_MEMBERNOTFOUND => unsafe {
                self.invoke_raw(dispid, DISPATCH_METHOD, &params)
                    .map(|v| DispatchResult::new(v, &self.ctx))
                    .map_err(|e| com_err("get", name, &self.ctx, e).into())
            },
            Err(e) => Err(com_err("get", name, &self.ctx, e).into()),
        }
    }

    /// 调用方法，如 `call("Delete", vec![])` / `call("Range", vec![start, end])`。
    /// `Range(Start,End)`/`Item(i)` 这类参数化属性用 METHOD 报 MEMBERNOTFOUND 时自动退化为 PROPERTYGET 再试。
    pub fn call(&self, name: &str, args: Vec<ComValue>) -> Result<DispatchResult> {
        let dispid = self
            .get_dispid(name)
            .map_err(|e| com_err("dispid", name, &self.ctx, e))?;
        let ordered = order_args(args, Vec::new());
        let mut variants: Vec<VARIANT> = ordered.into_iter().map(VARIANT::from).collect();
        let params = DISPPARAMS {
            rgvarg: variants.as_mut_ptr(),
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: variants.len() as u32,
            cNamedArgs: 0,
        };
        match unsafe { self.invoke_raw(dispid, DISPATCH_METHOD, &params) } {
            Ok(v) => Ok(DispatchResult::new(v, &self.ctx)),
            Err(e) if e.code().0 as u32 == DISP_E_MEMBERNOTFOUND => unsafe {
                self.invoke_raw(dispid, DISPATCH_PROPERTYGET, &params)
                    .map(|v| DispatchResult::new(v, &self.ctx))
                    .map_err(|e| com_err("call", name, &self.ctx, e).into())
            },
            Err(e) => Err(com_err("call", name, &self.ctx, e).into()),
        }
    }

    /// 带命名参数的方法调用，如 `call_named("Add", vec![], &[("Text", v)])`。
    ///
    /// **顺序很重要**：COM 的 DISPPARAMS 规定 `rgvarg` 里**具名参数在前**（与 `rgdispidNamedArgs`
    /// 一一对应），位置参数排在后面且**逆序**。写反会导致同类型参数被静默对调、类型不同时报
    /// `0x80020005 类型不匹配`。
    pub fn call_named(
        &self,
        name: &str,
        args: Vec<ComValue>,
        named: &[(&str, ComValue)],
    ) -> Result<DispatchResult> {
        let mut names: Vec<BSTR> = Vec::with_capacity(1 + named.len());
        names.push(BSTR::from(name));
        names.extend(named.iter().map(|(n, _)| BSTR::from(*n)));
        let ptrs: Vec<PCWSTR> = names.iter().map(|b| PCWSTR(b.as_ptr())).collect();
        let mut dispids: Vec<i32> = vec![0; names.len()];
        unsafe {
            self.disp.GetIDsOfNames(
                &GUID::default(),
                ptrs.as_ptr(),
                ptrs.len() as u32,
                LOCALE_USER_DEFAULT,
                dispids.as_mut_ptr(),
            )
        }
        .map_err(|e| com_err("dispid", name, &self.ctx, e))?;

        let named_vals: Vec<ComValue> = named.iter().map(|(_, v)| v.clone()).collect();
        let ordered = order_args(args.clone(), named_vals.clone());
        let variants: Vec<VARIANT> = ordered.into_iter().map(VARIANT::from).collect();
        let mut named_ids: Vec<i32> = dispids[1..].to_vec();
        let params = DISPPARAMS {
            rgvarg: variants.as_ptr() as *mut _,
            rgdispidNamedArgs: named_ids.as_mut_ptr(),
            cArgs: variants.len() as u32,
            cNamedArgs: named_ids.len() as u32,
        };
        match unsafe { self.invoke_raw(dispids[0], DISPATCH_METHOD, &params) } {
            Ok(v) => Ok(DispatchResult::new(v, &self.ctx)),
            // 调用点的具名参数都按签名顺序列出，被拒时按同一顺序降级为纯位置参数重试一次。
            Err(e) if rejects_named_args(e.code().0 as u32) => {
                let all: Vec<ComValue> = named_vals.into_iter().chain(args).collect();
                self.call(name, all)
            }
            Err(e) => Err(com_err("call", name, &self.ctx, e).into()),
        }
    }

    /// 写属性，如 `put("Text", vec!["新文本".into()])`。
    pub fn put(&self, name: &str, args: Vec<ComValue>) -> Result<()> {
        let mut variants: Vec<VARIANT> = args.into_iter().map(VARIANT::from).collect();
        let dispid = self
            .get_dispid(name)
            .map_err(|e| com_err("dispid", name, &self.ctx, e))?;
        variants.reverse();
        let mut dispid_named: i32 = DISPID_PROPERTYPUT;
        let params = DISPPARAMS {
            rgvarg: variants.as_mut_ptr(),
            rgdispidNamedArgs: &mut dispid_named,
            cArgs: variants.len() as u32,
            cNamedArgs: 1,
        };
        unsafe { self.invoke_raw(dispid, DISPATCH_PROPERTYPUT, &params) }
            .map_err(|e| com_err("put", name, &self.ctx, e).into())
            .map(|_| ())
    }
}

/// 方法调用 / 属性读取的返回值，携带调用处对象路径便于错误定位。
pub struct DispatchResult {
    pub v: VARIANT,
    ctx: String,
}

impl DispatchResult {
    pub(crate) fn new(v: VARIANT, ctx: &str) -> Self {
        Self {
            v,
            ctx: ctx.to_string(),
        }
    }

    pub fn is_object(&self) -> bool {
        unsafe { self.v.Anonymous.Anonymous.vt }.0 == VT_DISPATCH.0
    }

    /// 按 vt 转 [WordValue]。
    pub fn value(self) -> WordValue {
        variant_to_wordvalue(&self.v)
    }

    pub fn string(self) -> Result<String> {
        Ok(self.value().to_string())
    }

    pub fn int(self) -> Result<i32> {
        let Self { v, ctx } = self;
        match variant_to_wordvalue(&v) {
            WordValue::Number(f) => Ok(f as i32),
            WordValue::Bool(b) => Ok(b as i32),
            other => {
                Err(WordError::new("to_int", ctx, format!("期望整数，实际为 {other:?}")).into())
            }
        }
    }

    pub fn float(self) -> Result<f64> {
        let Self { v, ctx } = self;
        match variant_to_wordvalue(&v) {
            WordValue::Number(f) => Ok(f),
            other => {
                Err(WordError::new("to_float", ctx, format!("期望数字，实际为 {other:?}")).into())
            }
        }
    }

    pub fn bool(self) -> Result<bool> {
        let Self { v, ctx } = self;
        match variant_to_wordvalue(&v) {
            WordValue::Bool(b) => Ok(b),
            WordValue::Number(f) => Ok(f != 0.0),
            other => {
                Err(WordError::new("to_bool", ctx, format!("期望布尔值，实际为 {other:?}")).into())
            }
        }
    }

    /// 转为 COM 对象包装（ctx 继承调用处对象路径）。
    pub fn idispatch(self) -> Result<IDispatchWrapper> {
        self.idispatch_child("")
    }

    /// 转为 COM 对象包装，ctx 追加子对象标签。
    pub fn idispatch_child(self, label: &str) -> Result<IDispatchWrapper> {
        let ctx = if label.is_empty() {
            self.ctx
        } else {
            format!("{} > {}", self.ctx, label)
        };
        let disp: Option<IDispatch> =
            unsafe { (*self.v.Anonymous.Anonymous.Anonymous.pdispVal).clone() };
        match disp {
            Some(d) => Ok(IDispatchWrapper { disp: d, ctx }),
            None => Err(WordError::new(
                "to_object",
                ctx,
                "返回值不是 COM 对象（可能为空/标量/数组）",
            )
            .into()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn strs(v: &[ComValue]) -> Vec<String> {
        v.iter()
            .map(|c| match c {
                ComValue::Str(s) => s.clone(),
                other => format!("{other:?}"),
            })
            .collect()
    }

    #[test]
    fn positional_args_are_reversed() {
        let got = order_args(vec!["a".into(), "b".into(), "c".into()], vec![]);
        assert_eq!(strs(&got), ["c", "b", "a"]);
    }

    #[test]
    fn named_comes_first_then_positional_reversed() {
        // 具名参数在前（顺序保持），位置参数逆序在后。
        let got = order_args(
            vec!["p1".into(), "p2".into()],
            vec!["n1".into(), "n2".into()],
        );
        assert_eq!(strs(&got), ["n1", "n2", "p2", "p1"]);
    }
}
