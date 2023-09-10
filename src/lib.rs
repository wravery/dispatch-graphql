#![allow(non_snake_case)]

use std::{
    cell::UnsafeCell,
    collections::BTreeMap,
    iter, mem,
    sync::{
        atomic::{AtomicI32, Ordering},
        mpsc, Arc, Mutex, Once, Weak,
    },
    thread,
};

use windows::{
    core::*,
    Win32::{
        Foundation::*,
        System::{Com::*, LibraryLoader::*, Ole::*, Variant::*},
        UI::WindowsAndMessaging::{
            self, CREATESTRUCTW, GWLP_USERDATA, WINDOW_EX_STYLE, WINDOW_STYLE, WNDCLASSEXW,
        },
    },
};
use windows_implement::implement;
use windows_interface::interface;

use serde_json::Value;

use gqlmapi_rs::*;

macro_rules! impl_dispatch {
    ($type:ident, $interface:ident) => {
        impl IDispatch_Impl for $type {
            fn GetTypeInfoCount(&self) -> windows::core::Result<u32> {
                Ok(1)
            }

            fn GetTypeInfo(&self, itinfo: u32, _lcid: u32) -> windows::core::Result<ITypeInfo> {
                if itinfo != 0 {
                    TYPE_E_ELEMENTNOTFOUND.ok()?;
                }

                unsafe {
                    let type_lib = match &mut *self.type_lib.get() {
                        Some(type_lib) => type_lib.clone(),
                        None => {
                            let type_lib = load_type_lib()?;
                            *self.type_lib.get() = Some(type_lib.clone());
                            type_lib
                        }
                    };

                    type_lib.GetTypeInfoOfGuid(&<$interface as ComInterface>::IID)
                }
            }

            fn GetIDsOfNames(
                &self,
                _riid: *const windows::core::GUID,
                rgsznames: *const windows::core::PCWSTR,
                cnames: u32,
                lcid: u32,
                rgdispid: *mut i32,
            ) -> windows::core::Result<()> {
                let type_info = self.GetTypeInfo(0, lcid)?;
                unsafe { type_info.GetIDsOfNames(rgsznames, cnames, rgdispid) }
            }

            fn Invoke(
                &self,
                dispidmember: i32,
                _riid: *const windows::core::GUID,
                lcid: u32,
                wflags: DISPATCH_FLAGS,
                pdispparams: *const DISPPARAMS,
                pvarresult: *mut VARIANT,
                pexcepinfo: *mut EXCEPINFO,
                puargerr: *mut u32,
            ) -> windows::core::Result<()> {
                let type_info = self.GetTypeInfo(0, lcid)?;
                unsafe {
                    let this: IGraphQLService = self.cast()?;
                    type_info.Invoke(
                        this.as_raw(),
                        dispidmember,
                        wflags,
                        pdispparams as *mut _,
                        pvarresult,
                        pexcepinfo,
                        puargerr,
                    )
                }
            }
        }
    };
}

#[interface("E7706FBE-117E-4F1C-AD0F-DC058C6F867B")]
unsafe trait IResultPayload: IDispatch {
    fn results(&self, payload: *mut BSTR) -> HRESULT;
}

#[implement(IResultPayload)]
struct ResultPayload {
    type_lib: UnsafeCell<Option<ITypeLib>>,
    results: Value,
}

impl_dispatch!(ResultPayload, IResultPayload);

impl IResultPayload_Impl for ResultPayload {
    unsafe fn results(&self, payload: *mut BSTR) -> HRESULT {
        let Some(payload) = payload.as_mut() else {
            return E_INVALIDARG;
        };
        let Ok(results) = serde_json::to_string(&self.results) else {
            return E_UNEXPECTED;
        };
        let results: Vec<_> = results
            .as_str()
            .encode_utf16()
            .chain(iter::once(0_u16))
            .collect();
        *payload = SysAllocStringLen(Some(&results));
        S_OK
    }
}

#[interface("ABE787E3-CE4E-4B08-9780-E15076CD6045")]
unsafe trait IPendingPayload: IDispatch {
    fn pending(&self, key: *mut i32) -> HRESULT;
}

#[implement(IPendingPayload)]
struct PendingPayload {
    type_lib: UnsafeCell<Option<ITypeLib>>,
    pending: i32,
}

impl_dispatch!(PendingPayload, IPendingPayload);

impl IPendingPayload_Impl for PendingPayload {
    unsafe fn pending(&self, key: *mut i32) -> HRESULT {
        let Some(key) = key.as_mut() else {
            return E_INVALIDARG;
        };
        *key = self.pending;
        S_OK
    }
}

#[interface("A662B860-E098-43D3-A433-BE57DCBC15C3")]
unsafe trait INextPayload: IDispatch {
    fn next(&self, payload: *mut BSTR) -> HRESULT;
    fn subscription(&self, key: *mut i32) -> HRESULT;
}

#[implement(INextPayload)]
struct NextPayload {
    type_lib: UnsafeCell<Option<ITypeLib>>,
    next: Value,
    subscription: i32,
}

impl_dispatch!(NextPayload, INextPayload);

impl INextPayload_Impl for NextPayload {
    unsafe fn next(&self, payload: *mut BSTR) -> HRESULT {
        let Some(payload) = payload.as_mut() else {
            return E_INVALIDARG;
        };
        let Ok(results) = serde_json::to_string(&self.next) else {
            return E_UNEXPECTED;
        };
        let results: Vec<_> = results
            .as_str()
            .encode_utf16()
            .chain(iter::once(0_u16))
            .collect();
        *payload = SysAllocStringLen(Some(&results));
        S_OK
    }

    unsafe fn subscription(&self, key: *mut i32) -> HRESULT {
        let Some(key) = key.as_mut() else {
            return E_INVALIDARG;
        };
        *key = self.subscription;
        S_OK
    }
}

#[interface("FA294686-DB83-4268-A84F-157012D56033")]
unsafe trait IGraphQLService: IDispatch {
    fn fetchQuery(
        &self,
        query: BSTR,
        operation_name: BSTR,
        variables: BSTR,
        next_callback: *mut IDispatch,
        result: *mut *mut IDispatch,
    ) -> HRESULT;
    fn unsubscribe(&self, key: i32) -> HRESULT;
}

#[implement(IGraphQLService)]
pub struct GraphQLService {
    type_lib: UnsafeCell<Option<ITypeLib>>,
    gqlmapi: MAPIGraphQL,
    next_subscription: AtomicI32,
    subscriptions: Arc<Mutex<BTreeMap<i32, Mutex<Subscription>>>>,
    dispatch_queue: DeferCallbackQueue,
}

impl GraphQLService {
    pub fn new() -> Self {
        Self {
            type_lib: UnsafeCell::new(None),
            gqlmapi: MAPIGraphQL::new(true),
            next_subscription: AtomicI32::new(1),
            subscriptions: Arc::new(Mutex::new(BTreeMap::new())),
            dispatch_queue: DeferCallbackQueue::new(),
        }
    }
}

impl_dispatch!(GraphQLService, IGraphQLService);

impl IGraphQLService_Impl for GraphQLService {
    unsafe fn fetchQuery(
        &self,
        query: BSTR,
        operation_name: BSTR,
        variables: BSTR,
        next_callback: *mut IDispatch,
        result: *mut *mut IDispatch,
    ) -> HRESULT {
        let (Ok(query), Ok(operation_name), Ok(variables), Some(next_callback)) = (
            String::from_utf16(query.as_wide()),
            String::from_utf16(operation_name.as_wide()),
            String::from_utf16(variables.as_wide()),
            next_callback.as_ref(),
        ) else {
            return E_INVALIDARG;
        };
        let Ok(parsed_query) = self.gqlmapi.parse_query(&query) else {
            return E_INVALIDARG;
        };

        let (tx_next, rx_next) = mpsc::channel();
        let (tx_complete, rx_complete) = mpsc::channel();

        let (key, subscriptions) = {
            let subscription =
                self.gqlmapi
                    .subscribe(parsed_query.clone(), &operation_name, &variables);

            {
                let Ok(mut locked_subscription) = subscription.lock() else {
                    return E_UNEXPECTED;
                };
                if locked_subscription.listen(tx_next, tx_complete).is_err() {
                    return E_UNEXPECTED;
                }
            }

            let Ok(mut subscriptions) = self.subscriptions.lock() else {
                return E_UNEXPECTED;
            };
            let key: i32 = self.next_subscription.fetch_add(1, Ordering::Relaxed);
            subscriptions.insert(key, subscription);

            (key, self.subscriptions.clone())
        };

        match rx_complete.try_recv() {
            Ok(()) => {
                let Ok(results) = rx_next.recv() else {
                    return E_UNEXPECTED;
                };
                let Ok(results) = serde_json::from_str(&results) else {
                    return E_UNEXPECTED;
                };
                let payload: IResultPayload = ResultPayload {
                    type_lib: UnsafeCell::new(None),
                    results,
                }
                .into();
                *result = payload.as_raw() as *mut _;
                drop_subscription(key, &subscriptions)
            }
            Err(_) => {
                let Some(dispatcher) = self.dispatch_queue.get_dispatcher() else {
                    return E_UNEXPECTED;
                };
                self.dispatch_queue
                    .add_subscription(next_callback.clone(), key);
                thread::spawn::<_, HRESULT>(move || {
                    loop {
                        match rx_next.recv() {
                            Ok(next) => {
                                let Ok(next) = serde_json::from_str(&next) else {
                                    return E_UNEXPECTED;
                                };
                                dispatcher.dispatch(next, key);
                            }
                            Err(_) => {
                                break;
                            }
                        }
                    }
                    drop_subscription(key, &subscriptions)
                });

                let payload: IPendingPayload = PendingPayload {
                    type_lib: UnsafeCell::new(None),
                    pending: key,
                }
                .into();
                *result = payload.as_raw() as *mut _;
                S_OK
            }
        }
    }

    unsafe fn unsubscribe(&self, key: i32) -> HRESULT {
        drop_subscription(key, &self.subscriptions)
    }
}

unsafe fn load_type_lib() -> windows::core::Result<ITypeLib> {
    let mut buffer: mem::MaybeUninit<[u16; MAX_PATH as usize]> = mem::MaybeUninit::uninit();
    let count =
        GetModuleFileNameW(GetModuleHandleW(PCWSTR::null())?, &mut *buffer.as_mut_ptr()) as usize;
    let buffer = buffer.assume_init();
    if count >= buffer.len() {
        return Err(ERROR_INSUFFICIENT_BUFFER.into());
    }
    let Ok(file_name) = String::from_utf16(&buffer[0..count]) else {
        return Err(E_UNEXPECTED.into());
    };
    let resource_name: Vec<_> = format!("{file_name}\\1")
        .as_str()
        .encode_utf16()
        .chain(iter::once(0_u16))
        .collect();
    LoadTypeLibEx(PCWSTR(resource_name.as_ptr()), REGKIND_NONE)
}

fn drop_subscription(
    key: i32,
    subscriptions: &Mutex<BTreeMap<i32, Mutex<Subscription>>>,
) -> HRESULT {
    let Ok(mut subscriptions) = subscriptions.lock() else {
        return E_UNEXPECTED;
    };
    subscriptions.remove(&key);
    S_OK
}

const MODULE_NAME: PCWSTR = w!(r#"dispatch_graphql.dll"#);

fn get_module_handle() -> HMODULE {
    unsafe { GetModuleHandleW(MODULE_NAME) }.unwrap_or_default()
}

#[derive(Default)]
struct UniqueHwnd(Option<HWND>);

impl Drop for UniqueHwnd {
    fn drop(&mut self) {
        if let Some(window) = self.0.take() {
            unsafe {
                let _ = WindowsAndMessaging::DestroyWindow(window);
            }
        }
    }
}

const CALLBACK_WINDOW_CLASS_NAME: PCWSTR = w!("NextCallback");
const DISPATCH_CALLBACKS: u32 = WindowsAndMessaging::WM_USER;

struct DeferCallbackDispatcher {
    window: Weak<UniqueHwnd>,
    tx: mpsc::Sender<(Value, i32)>,
}

impl DeferCallbackDispatcher {
    fn dispatch(&self, next: Value, subscription: i32) {
        if let Some(window) = self.window.upgrade() {
            if let Some(window) = window.0 {
                if let Ok(()) = self.tx.send((next, subscription)) {
                    unsafe {
                        let _ = WindowsAndMessaging::PostMessageW(
                            window,
                            DISPATCH_CALLBACKS,
                            WPARAM(0),
                            LPARAM(0),
                        );
                    }
                }
            }
        }
    }
}

struct NextCallbacks {
    rx: mpsc::Receiver<(Value, i32)>,
    next_callbacks: BTreeMap<i32, IDispatch>,
}

struct DeferCallbackQueue {
    window: Arc<UniqueHwnd>,
    tx: Mutex<mpsc::Sender<(Value, i32)>>,
}

impl DeferCallbackQueue {
    fn new() -> Self {
        Self::register_window_class();

        let (tx, rx) = mpsc::channel();

        Self {
            window: Arc::new(UniqueHwnd(Some(unsafe {
                WindowsAndMessaging::CreateWindowExW(
                    WINDOW_EX_STYLE(0),
                    CALLBACK_WINDOW_CLASS_NAME,
                    None,
                    WINDOW_STYLE(0),
                    0,
                    0,
                    0,
                    0,
                    HWND::default(),
                    None,
                    get_module_handle(),
                    Some(Box::into_raw(Box::new(NextCallbacks {
                        rx,
                        next_callbacks: BTreeMap::new(),
                    })) as *const _),
                )
            }))),
            tx: Mutex::new(tx),
        }
    }

    unsafe fn add_subscription(&self, next_callback: IDispatch, subscription: i32) {
        if let Some(window) = &self.window.0 {
            let callbacks: *mut NextCallbacks =
                WindowsAndMessaging::GetWindowLongPtrW(*window, GWLP_USERDATA) as *mut _;
            if !callbacks.is_null() {
                let callbacks = Box::leak(Box::from_raw(callbacks));
                let _ = callbacks.next_callbacks.insert(subscription, next_callback);
            }
        }
    }

    fn get_dispatcher(&self) -> Option<DeferCallbackDispatcher> {
        self.tx.lock().ok().map(|tx| DeferCallbackDispatcher {
            window: Arc::downgrade(&self.window),
            tx: tx.clone(),
        })
    }

    fn register_window_class() {
        static REGISTER: Once = Once::new();
        REGISTER.call_once(|| {
            let wnd_class = WNDCLASSEXW {
                cbSize: mem::size_of::<WNDCLASSEXW>() as u32,
                lpfnWndProc: Some(Self::window_proc),
                hInstance: HINSTANCE(get_module_handle().0),
                lpszClassName: CALLBACK_WINDOW_CLASS_NAME,
                ..Default::default()
            };

            unsafe {
                WindowsAndMessaging::RegisterClassExW(&wnd_class);
            }
        })
    }

    unsafe extern "system" fn window_proc(
        window: HWND,
        message: u32,
        wparam: WPARAM,
        lparam: LPARAM,
    ) -> LRESULT {
        match message {
            WindowsAndMessaging::WM_CREATE => {
                let create_struct: *const CREATESTRUCTW = lparam.0 as *const _;
                if !create_struct.is_null() && !(*create_struct).lpCreateParams.is_null() {
                    WindowsAndMessaging::SetWindowLongPtrW(
                        window,
                        GWLP_USERDATA,
                        (*create_struct).lpCreateParams as _,
                    );

                    LRESULT(0)
                } else {
                    LRESULT(-1)
                }
            }

            DISPATCH_CALLBACKS => {
                let callbacks: *mut NextCallbacks =
                    WindowsAndMessaging::GetWindowLongPtrW(window, GWLP_USERDATA) as *mut _;
                if !callbacks.is_null() {
                    let callbacks = Box::leak(Box::from_raw(callbacks));
                    while let Ok((next, subscription)) = callbacks.rx.try_recv() {
                        let payload: INextPayload = NextPayload {
                            type_lib: UnsafeCell::new(None),
                            next,
                            subscription,
                        }
                        .into();
                        if let (Some(next_callback), Ok(payload)) =
                            (callbacks.next_callbacks.get(&subscription), payload.cast())
                        {
                            let mut rgvarg = [VariantInit(); 1];
                            (*rgvarg[0].Anonymous.Anonymous).vt = VT_DISPATCH;
                            *(*rgvarg[0].Anonymous.Anonymous).Anonymous.pdispVal = Some(payload);
                            let params = DISPPARAMS {
                                rgvarg: rgvarg.as_mut_ptr(),
                                cArgs: 1,
                                ..Default::default()
                            };
                            let _ = next_callback.Invoke(
                                DISPID_UNKNOWN,
                                &GUID::default(),
                                0,
                                DISPATCH_METHOD,
                                &params as *const _,
                                None,
                                None,
                                None,
                            );
                            let _ = mem::take(
                                &mut *(*rgvarg[0].Anonymous.Anonymous).Anonymous.pdispVal,
                            );
                        }
                    }
                }
                LRESULT(0)
            }

            WindowsAndMessaging::WM_DESTROY => {
                let callbacks: *mut NextCallbacks =
                    WindowsAndMessaging::SetWindowLongPtrW(window, GWLP_USERDATA, 0) as *mut _;
                if !callbacks.is_null() {
                    let _ = Box::from_raw(callbacks);
                }
                LRESULT(0)
            }

            _ => WindowsAndMessaging::DefWindowProcW(window, message, wparam, lparam),
        }
    }
}

pub fn add(left: usize, right: usize) -> usize {
    left + right
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        let result = add(2, 2);
        assert_eq!(result, 4);
    }
}
