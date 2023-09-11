# dispatch-graphql

This sample demonstrates how to use a [native host object with](https://learn.microsoft.com/en-us/microsoft-edge/webview2/how-to/hostobject?tabs=win32) with
WebView2's COM interfaces exposed by [webview2-com](https://crates.io/crates/webview2-com). It wraps the [GqlMAPI](https://github.com/microsoft/gqlmapi)
bindings in [gqlmapi-rs](https://crates.io/crates/gqlmapi-rs) in an [IDispatch](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nn-oaidl-idispatch)
interface. This is exposed from the DLL with the `CreateService` function.

The [sample.rs](./examples/sample.rs) example is based on the example from [webview2-com](https://crates.io/crates/webview2-com), with some customizations.
Besides calling `CreateService` and `AddHostObjectToScript`, it also injects an initial script contained in [sample.js](./examples/sample.js) which
demonstrates invoking the `fetchQuery` method on the host object to get the store and folder ID of the user's default Inbox in Outlook, and then subscribe to
async item updates on the 10 most recent items in the Inbox. As subscription events are delivered (e.g. marking items as read/unread), they should show up
in the JavaScript console in the Dev Tools/Inspect window.
