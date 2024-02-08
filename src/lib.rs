#![allow(non_snake_case)]

use std::{
    cell::UnsafeCell,
    collections::BTreeMap,
    ffi::c_void,
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
            self, PeekMessageW, CREATESTRUCTW, GWLP_USERDATA, MSG, PM_NOREMOVE, WINDOW_EX_STYLE,
            WINDOW_STYLE, WM_USER, WNDCLASSEXW,
        },
    },
};
use windows_implement::implement;
use windows_interface::interface;

use serde::Serialize;
use serde_json::Value;

use gqlmapi_rs::*;

macro_rules! impl_dispatch {
    ($type:ident, $interface:ident) => {
        #[allow(clippy::not_unsafe_ptr_arg_deref)]
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
                    let this: $interface = self.cast()?;
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

/// External entry point to create an IDispatch object for the service. If this function succeeds,
/// the caller is responsible for calling `Release` on the `IDispatch` interface pointer returned
/// through the `result` out-param. The initial value that `result` points to must be `null`.
///
/// # Safety
///
/// The `result` parameter must be a valid pointer.
#[no_mangle]
pub unsafe extern "C" fn CreateService(result: *mut *mut c_void) -> HRESULT {
    // Parameter validation
    let Some(result) = result.as_mut() else {
        return E_POINTER;
    };
    if !result.is_null() {
        return E_INVALIDARG;
    }

    let service: IGraphQLService = GraphQLService::new().into();
    let Ok(service) = service.cast::<IDispatch>() else {
        return E_NOINTERFACE;
    };
    *result = service.into_raw();
    S_OK
}

#[derive(Serialize)]
struct ResultPayload {
    results: Value,
}

#[derive(Serialize)]
struct PendingPayload {
    pending: i32,
}

#[derive(Serialize)]
struct NextPayload {
    next: Value,
    subscription: i32,
}

fn serialize_results<T: Serialize>(payload: T) -> BSTR {
    serde_json::to_string(&payload)
        .map_err(|_| ())
        .and_then(|payload| {
            let payload: Vec<_> = payload
                .as_str()
                .encode_utf16()
                .chain(iter::once(0_u16))
                .collect();
            BSTR::from_wide(&payload).map_err(|_| ())
        })
        .unwrap_or_default()
}

#[interface("FA294686-DB83-4268-A84F-157012D56033")]
pub unsafe trait IGraphQLService: IDispatch {
    fn fetchQuery(
        &self,
        query: BSTR,
        operation_name: BSTR,
        variables: BSTR,
        next_callback: *mut c_void,
        result: *mut BSTR,
    ) -> HRESULT;
    fn unsubscribe(&self, key: i32) -> HRESULT;
}

#[implement(IGraphQLService, IDispatch)]
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

impl Default for GraphQLService {
    fn default() -> Self {
        Self::new()
    }
}

impl_dispatch!(GraphQLService, IGraphQLService);

impl IGraphQLService_Impl for GraphQLService {
    unsafe fn fetchQuery(
        &self,
        query: BSTR,
        operation_name: BSTR,
        variables: BSTR,
        next_callback: *mut c_void,
        result: *mut BSTR,
    ) -> HRESULT {
        // The caller (WebView2) retains ownership of these BSTRs, so suppress the drop destructor
        // on the windows::core::BSTR arguments constructed by the generated IGraphQLService impl.
        let (query, operation_name, variables) = (
            mem::ManuallyDrop::new(query),
            mem::ManuallyDrop::new(operation_name),
            mem::ManuallyDrop::new(variables),
        );
        let (Ok(query), Ok(operation_name), Ok(variables)) = (
            String::from_utf16(query.as_wide()),
            String::from_utf16(operation_name.as_wide()),
            String::from_utf16(variables.as_wide()),
        ) else {
            return E_INVALIDARG;
        };
        if next_callback.is_null() {
            return E_INVALIDARG;
        }
        let raw = IDispatch::from_raw(next_callback);
        let next_callback = raw.clone();
        mem::forget(raw);
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
                *result = serialize_results(&ResultPayload { results });
                drop_subscription(key, &subscriptions)
            }
            Err(_) => {
                let Some(dispatcher) = self.dispatch_queue.get_dispatcher() else {
                    return E_UNEXPECTED;
                };
                self.dispatch_queue.add_subscription(next_callback, key);
                thread::spawn::<_, HRESULT>(move || {
                    while let Ok(next) = rx_next.recv() {
                        let Ok(next) = serde_json::from_str(&next) else {
                            return E_UNEXPECTED;
                        };
                        dispatcher.dispatch(next, key);
                    }
                    dispatcher.remove_callback(key);
                    drop_subscription(key, &subscriptions)
                });

                *result = serialize_results(PendingPayload { pending: key });
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
    let count = GetModuleFileNameW(get_module_handle(), &mut *buffer.as_mut_ptr()) as usize;
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

const MODULE_NAME: PCSTR =
    PCSTR::from_raw(concat!(env!("CARGO_CRATE_NAME"), ".dll", '\0').as_ptr());

fn get_module_handle() -> HMODULE {
    unsafe { GetModuleHandleA(MODULE_NAME) }.unwrap_or_default()
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
const REMOVE_CALLBACK: u32 = WindowsAndMessaging::WM_USER + 1;

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

    fn remove_callback(&self, subscription: i32) {
        if let (Ok(subscription), Some(window)) =
            (isize::try_from(subscription), self.window.upgrade())
        {
            if let Some(window) = window.0 {
                unsafe {
                    let _ = WindowsAndMessaging::PostMessageW(
                        window,
                        REMOVE_CALLBACK,
                        WPARAM(0),
                        LPARAM(subscription),
                    );
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
        Self::ensure_message_queue();

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

    fn ensure_message_queue() {
        let mut msg = MSG::default();
        let hwnd = HWND::default();
        unsafe { PeekMessageW(&mut msg, hwnd, WM_USER, WM_USER, PM_NOREMOVE) };
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
                        let payload = serialize_results(&NextPayload { next, subscription });
                        if let Some(next_callback) = callbacks.next_callbacks.get(&subscription) {
                            let mut rgvarg = [VariantInit(); 1];
                            #[allow(clippy::explicit_auto_deref)]
                            {
                                (*rgvarg[0].Anonymous.Anonymous).vt = VT_BSTR;
                                *(*rgvarg[0].Anonymous.Anonymous).Anonymous.bstrVal = payload;
                            }
                            let params = DISPPARAMS {
                                rgvarg: rgvarg.as_mut_ptr(),
                                cArgs: 1,
                                ..Default::default()
                            };
                            const LOCALE_USER_DEFAULT: u32 = 0x400;
                            let _ = next_callback.Invoke(
                                DISPID_UNKNOWN,
                                &GUID::default(),
                                LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD,
                                &params as *const _,
                                None,
                                None,
                                None,
                            );
                            ClearVariantArray(&mut rgvarg);
                        }
                    }
                }
                LRESULT(0)
            }

            REMOVE_CALLBACK => {
                if let Ok(key) = i32::try_from(lparam.0) {
                    let callbacks: *mut NextCallbacks =
                        WindowsAndMessaging::GetWindowLongPtrW(window, GWLP_USERDATA) as *mut _;
                    if !callbacks.is_null() {
                        let callbacks = Box::leak(Box::from_raw(callbacks));
                        callbacks.next_callbacks.remove(&key);
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
