#![windows_subsystem = "windows"]

extern crate webview2_com;
extern crate windows;

use std::{ffi::c_void, fmt, mem, ptr, rc::Rc, sync::mpsc};

use windows::{
    core::*,
    Win32::{
        Foundation::{E_POINTER, HWND, LPARAM, LRESULT, RECT, SIZE, WPARAM},
        Graphics::Gdi,
        System::{Com::*, LibraryLoader, Threading, Variant::*, WinRT::EventRegistrationToken},
        UI::{
            HiDpi,
            Input::KeyboardAndMouse,
            WindowsAndMessaging::{self, MSG, WINDOW_LONG_PTR_INDEX, WNDCLASSW},
        },
    },
};

use webview2_com::{Microsoft::Web::WebView2::Win32::*, *};

fn main() -> Result<()> {
    unsafe {
        CoInitializeEx(None, COINIT_APARTMENTTHREADED)?;
    }
    set_process_dpi_awareness()?;

    let mut webview = WebView::create(None, true)?;

    #[link(name = "dispatch_graphql", kind = "raw-dylib")]
    extern "C" {
        fn CreateService(result: *mut *mut c_void) -> HRESULT;
    }

    unsafe {
        let mut service = ptr::null_mut();
        if CreateService(&mut service).is_ok() && !service.is_null() {
            let service = IDispatch::from_raw(service);
            let mut host_object = VariantInit();
            #[allow(clippy::explicit_auto_deref)]
            {
                (*host_object.Anonymous.Anonymous).vt = VT_DISPATCH;
                *(*host_object.Anonymous.Anonymous).Anonymous.pdispVal = Some(service);
            }
            let _ = webview
                .webview
                .AddHostObjectToScript(w!("graphql"), &mut host_object);
            let _ = VariantClear(&mut host_object);
        }
    }

    // Configure the target URL and add an init script to output the default store and inbox IDs.
    webview
        .set_title("webview2-com example (crates/webview2-com/examples)")?
        .init(include_str!("sample.js"))?
        .navigate("https://github.com/wravery/dispatch-graphql")?;

    // Off we go....
    webview.run()
}

#[derive(Debug)]
pub enum Error {
    WebView2Error(webview2_com::Error),
    WindowsError(windows::core::Error),
    LockError,
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "{self:?}")
    }
}

impl From<webview2_com::Error> for Error {
    fn from(err: webview2_com::Error) -> Self {
        Self::WebView2Error(err)
    }
}

impl From<windows::core::Error> for Error {
    fn from(err: windows::core::Error) -> Self {
        Self::WindowsError(err)
    }
}

impl From<HRESULT> for Error {
    fn from(err: HRESULT) -> Self {
        Self::WindowsError(windows::core::Error::from(err))
    }
}

impl<'a, T: 'a> From<std::sync::PoisonError<T>> for Error {
    fn from(_: std::sync::PoisonError<T>) -> Self {
        Self::LockError
    }
}

impl<'a, T: 'a> From<std::sync::TryLockError<T>> for Error {
    fn from(_: std::sync::TryLockError<T>) -> Self {
        Self::LockError
    }
}

type Result<T> = std::result::Result<T, Error>;

struct Window(HWND);

impl Drop for Window {
    fn drop(&mut self) {
        unsafe {
            let _ = WindowsAndMessaging::DestroyWindow(self.0);
        }
    }
}

#[derive(Clone)]
pub struct FrameWindow {
    window: HWND,
    size: SIZE,
}

impl FrameWindow {
    fn new() -> Self {
        let hwnd = {
            let window_class = WNDCLASSW {
                lpfnWndProc: Some(window_proc),
                lpszClassName: w!("WebView"),
                ..Default::default()
            };

            unsafe {
                WindowsAndMessaging::RegisterClassW(&window_class);

                WindowsAndMessaging::CreateWindowExW(
                    Default::default(),
                    w!("WebView"),
                    w!("WebView"),
                    WindowsAndMessaging::WS_OVERLAPPEDWINDOW,
                    WindowsAndMessaging::CW_USEDEFAULT,
                    WindowsAndMessaging::CW_USEDEFAULT,
                    WindowsAndMessaging::CW_USEDEFAULT,
                    WindowsAndMessaging::CW_USEDEFAULT,
                    None,
                    None,
                    LibraryLoader::GetModuleHandleW(None).unwrap_or_default(),
                    None,
                )
            }
        };

        FrameWindow {
            window: hwnd,
            size: SIZE { cx: 0, cy: 0 },
        }
    }
}

struct WebViewController(ICoreWebView2Controller);

type WebViewSender = mpsc::Sender<Box<dyn FnOnce(WebView) + Send>>;
type WebViewReceiver = mpsc::Receiver<Box<dyn FnOnce(WebView) + Send>>;

#[derive(Clone)]
pub struct WebView {
    controller: Rc<WebViewController>,
    webview: Rc<ICoreWebView2>,
    tx: WebViewSender,
    rx: Rc<WebViewReceiver>,
    thread_id: u32,
    frame: Option<FrameWindow>,
    parent: HWND,
    url: String,
}

impl Drop for WebViewController {
    fn drop(&mut self) {
        unsafe { self.0.Close() }.unwrap();
    }
}

impl WebView {
    pub fn create(parent: Option<HWND>, debug: bool) -> Result<WebView> {
        let (parent, mut frame) = match parent {
            Some(hwnd) => (hwnd, None),
            None => {
                let frame = FrameWindow::new();
                (frame.window, Some(frame))
            }
        };

        let environment = {
            let (tx, rx) = mpsc::channel();

            CreateCoreWebView2EnvironmentCompletedHandler::wait_for_async_operation(
                Box::new(|environmentcreatedhandler| unsafe {
                    CreateCoreWebView2Environment(&environmentcreatedhandler)
                        .map_err(webview2_com::Error::WindowsError)
                }),
                Box::new(move |error_code, environment| {
                    error_code?;
                    tx.send(environment.ok_or_else(|| windows::core::Error::from(E_POINTER)))
                        .expect("send over mpsc channel");
                    Ok(())
                }),
            )?;

            rx.recv()
                .map_err(|_| Error::WebView2Error(webview2_com::Error::SendError))?
        }?;

        let controller = {
            let (tx, rx) = mpsc::channel();

            CreateCoreWebView2ControllerCompletedHandler::wait_for_async_operation(
                Box::new(move |handler| unsafe {
                    environment
                        .CreateCoreWebView2Controller(parent, &handler)
                        .map_err(webview2_com::Error::WindowsError)
                }),
                Box::new(move |error_code, controller| {
                    error_code?;
                    tx.send(controller.ok_or_else(|| windows::core::Error::from(E_POINTER)))
                        .expect("send over mpsc channel");
                    Ok(())
                }),
            )?;

            rx.recv()
                .map_err(|_| Error::WebView2Error(webview2_com::Error::SendError))?
        }?;

        let size = get_window_size(parent);
        let mut client_rect = RECT::default();
        unsafe {
            let _ = WindowsAndMessaging::GetClientRect(parent, &mut client_rect as *mut _);
            controller.SetBounds(RECT {
                left: 0,
                top: 0,
                right: size.cx,
                bottom: size.cy,
            })?;
            controller.SetIsVisible(true)?;
        }

        let webview = unsafe { controller.CoreWebView2()? };

        if !debug {
            unsafe {
                let settings = webview.Settings()?;
                settings.SetAreDefaultContextMenusEnabled(false)?;
                settings.SetAreDevToolsEnabled(false)?;
            }
        }

        if let Some(frame) = frame.as_mut() {
            frame.size = size;
        }

        let (tx, rx) = mpsc::channel();
        let rx = Rc::new(rx);
        let thread_id = unsafe { Threading::GetCurrentThreadId() };

        let webview = WebView {
            controller: Rc::new(WebViewController(controller)),
            webview: Rc::new(webview),
            tx,
            rx,
            thread_id,
            frame,
            parent,
            url: String::new(),
        };

        if webview.frame.is_some() {
            WebView::set_window_webview(parent, Some(Box::new(webview.clone())));
        }

        Ok(webview)
    }

    pub fn run(self) -> Result<()> {
        let webview = self.webview.as_ref();
        let url = self.url.clone();
        let (tx, rx) = mpsc::channel();

        if !url.is_empty() {
            let handler =
                NavigationCompletedEventHandler::create(Box::new(move |_sender, _args| {
                    tx.send(()).expect("send over mpsc channel");
                    Ok(())
                }));
            let mut token = EventRegistrationToken::default();
            unsafe {
                webview.add_NavigationCompleted(&handler, &mut token)?;
                let url = CoTaskMemPWSTR::from(url.as_str());
                webview.Navigate(*url.as_ref().as_pcwstr())?;
                let result = webview2_com::wait_with_pump(rx);
                webview.remove_NavigationCompleted(token)?;
                result?;
            }
        }

        if let Some(frame) = self.frame.as_ref() {
            let hwnd = frame.window;
            unsafe {
                WindowsAndMessaging::ShowWindow(hwnd, WindowsAndMessaging::SW_SHOW);
                Gdi::UpdateWindow(hwnd);
                KeyboardAndMouse::SetFocus(hwnd);
            }
        }

        let mut msg = MSG::default();
        let h_wnd = HWND::default();

        loop {
            while let Ok(f) = self.rx.try_recv() {
                (f)(self.clone());
            }

            unsafe {
                let result = WindowsAndMessaging::GetMessageW(&mut msg, h_wnd, 0, 0).0;

                match result {
                    -1 => break Err(windows::core::Error::from_win32().into()),
                    0 => break Ok(()),
                    _ => match msg.message {
                        WindowsAndMessaging::WM_APP => (),
                        _ => {
                            WindowsAndMessaging::TranslateMessage(&msg);
                            WindowsAndMessaging::DispatchMessageW(&msg);
                        }
                    },
                }
            }
        }
    }

    pub fn terminate(self) -> Result<()> {
        self.dispatch(|_webview| unsafe {
            WindowsAndMessaging::PostQuitMessage(0);
        })?;

        if self.frame.is_some() {
            WebView::set_window_webview(self.get_window(), None);
        }

        Ok(())
    }

    pub fn set_title(&mut self, title: &str) -> Result<&mut Self> {
        if let Some(frame) = self.frame.as_ref() {
            let title = CoTaskMemPWSTR::from(title);
            unsafe {
                let _ =
                    WindowsAndMessaging::SetWindowTextW(frame.window, *title.as_ref().as_pcwstr());
            }
        }
        Ok(self)
    }

    pub fn set_size(&mut self, width: i32, height: i32) -> Result<&mut Self> {
        if let Some(frame) = self.frame.as_mut() {
            frame.size = SIZE {
                cx: width,
                cy: height,
            };
            unsafe {
                self.controller.0.SetBounds(RECT {
                    left: 0,
                    top: 0,
                    right: width,
                    bottom: height,
                })?;

                let _ = WindowsAndMessaging::SetWindowPos(
                    frame.window,
                    None,
                    0,
                    0,
                    width,
                    height,
                    WindowsAndMessaging::SWP_NOACTIVATE
                        | WindowsAndMessaging::SWP_NOZORDER
                        | WindowsAndMessaging::SWP_NOMOVE,
                );
            }
        }
        Ok(self)
    }

    pub fn get_window(&self) -> HWND {
        self.parent
    }

    pub fn navigate(&mut self, url: &str) -> Result<&mut Self> {
        self.url = url.into();
        Ok(self)
    }

    pub fn init(&mut self, js: &str) -> Result<&mut Self> {
        let webview = self.webview.clone();
        let js = String::from(js);
        AddScriptToExecuteOnDocumentCreatedCompletedHandler::wait_for_async_operation(
            Box::new(move |handler| unsafe {
                let js = CoTaskMemPWSTR::from(js.as_str());
                webview
                    .AddScriptToExecuteOnDocumentCreated(*js.as_ref().as_pcwstr(), &handler)
                    .map_err(webview2_com::Error::WindowsError)
            }),
            Box::new(|error_code, _id| error_code),
        )?;

        Ok(self)
    }

    pub fn eval(&self, js: &str) -> Result<&Self> {
        let webview = self.webview.clone();
        let js = String::from(js);
        ExecuteScriptCompletedHandler::wait_for_async_operation(
            Box::new(move |handler| unsafe {
                let js = CoTaskMemPWSTR::from(js.as_str());
                webview
                    .ExecuteScript(*js.as_ref().as_pcwstr(), &handler)
                    .map_err(webview2_com::Error::WindowsError)
            }),
            Box::new(|error_code, _result| error_code),
        )?;
        Ok(self)
    }

    pub fn dispatch<F>(&self, f: F) -> Result<&Self>
    where
        F: FnOnce(WebView) + Send + 'static,
    {
        self.tx.send(Box::new(f)).expect("send the fn");

        unsafe {
            let _ = WindowsAndMessaging::PostThreadMessageW(
                self.thread_id,
                WindowsAndMessaging::WM_APP,
                WPARAM::default(),
                LPARAM::default(),
            );
        }
        Ok(self)
    }

    fn set_window_webview(hwnd: HWND, webview: Option<Box<WebView>>) -> Option<Box<WebView>> {
        unsafe {
            match SetWindowLong(
                hwnd,
                WindowsAndMessaging::GWLP_USERDATA,
                match webview {
                    Some(webview) => Box::into_raw(webview) as _,
                    None => 0_isize,
                },
            ) {
                0 => None,
                ptr => Some(Box::from_raw(ptr as *mut _)),
            }
        }
    }

    fn get_window_webview(hwnd: HWND) -> Option<Box<WebView>> {
        unsafe {
            let data = GetWindowLong(hwnd, WindowsAndMessaging::GWLP_USERDATA);

            match data {
                0 => None,
                _ => {
                    let webview_ptr = data as *mut WebView;
                    let raw = Box::from_raw(webview_ptr);
                    let webview = raw.clone();
                    mem::forget(raw);

                    Some(webview)
                }
            }
        }
    }
}

fn set_process_dpi_awareness() -> Result<()> {
    unsafe { HiDpi::SetProcessDpiAwareness(HiDpi::PROCESS_PER_MONITOR_DPI_AWARE)? };
    Ok(())
}

extern "system" fn window_proc(hwnd: HWND, msg: u32, w_param: WPARAM, l_param: LPARAM) -> LRESULT {
    let mut webview = match WebView::get_window_webview(hwnd) {
        Some(webview) => webview,
        None => return unsafe { WindowsAndMessaging::DefWindowProcW(hwnd, msg, w_param, l_param) },
    };

    let frame = webview
        .frame
        .as_mut()
        .expect("should only be called for owned windows");

    match msg {
        WindowsAndMessaging::WM_SIZE => {
            let size = get_window_size(hwnd);
            unsafe {
                webview
                    .controller
                    .0
                    .SetBounds(RECT {
                        left: 0,
                        top: 0,
                        right: size.cx,
                        bottom: size.cy,
                    })
                    .unwrap();
            }
            frame.size = size;
            LRESULT::default()
        }

        WindowsAndMessaging::WM_CLOSE => {
            unsafe {
                let _ = WindowsAndMessaging::DestroyWindow(hwnd);
            }
            LRESULT::default()
        }

        WindowsAndMessaging::WM_DESTROY => {
            webview.terminate().expect("window is gone");
            LRESULT::default()
        }

        _ => unsafe { WindowsAndMessaging::DefWindowProcW(hwnd, msg, w_param, l_param) },
    }
}

fn get_window_size(hwnd: HWND) -> SIZE {
    let mut client_rect = RECT::default();
    let _ = unsafe { WindowsAndMessaging::GetClientRect(hwnd, &mut client_rect as *mut _) };
    SIZE {
        cx: client_rect.right - client_rect.left,
        cy: client_rect.bottom - client_rect.top,
    }
}

#[allow(non_snake_case)]
#[cfg(target_pointer_width = "32")]
unsafe fn SetWindowLong(window: HWND, index: WINDOW_LONG_PTR_INDEX, value: isize) -> isize {
    WindowsAndMessaging::SetWindowLongW(window, index, value as _) as _
}

#[allow(non_snake_case)]
#[cfg(target_pointer_width = "64")]
unsafe fn SetWindowLong(window: HWND, index: WINDOW_LONG_PTR_INDEX, value: isize) -> isize {
    WindowsAndMessaging::SetWindowLongPtrW(window, index, value)
}

#[allow(non_snake_case)]
#[cfg(target_pointer_width = "32")]
unsafe fn GetWindowLong(window: HWND, index: WINDOW_LONG_PTR_INDEX) -> isize {
    WindowsAndMessaging::GetWindowLongW(window, index) as _
}

#[allow(non_snake_case)]
#[cfg(target_pointer_width = "64")]
unsafe fn GetWindowLong(window: HWND, index: WINDOW_LONG_PTR_INDEX) -> isize {
    WindowsAndMessaging::GetWindowLongPtrW(window, index)
}
