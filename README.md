# libTridentUI

**Want to develop a browser-based desktop application for legacy operating systems? Try libTridentUI!**

**-- An extremely lightweight Chrome Embedded Framework alternative for legacy operating systems.**

**-- Works perfectly on Microsoft Windows XP SP3!**

## Why use libTridentUI?

1. **Extremely lightweight**
   
   Modern browser based UI frameworks like Electron, CEF, and WebView2 are powerful but extremely heavy.
   
   They ship a full Chromium browser with your app which would obviously occupy a huge amount of storage space and require modern OS versions.
   
   If you're building a lightweight desktop utility that needs a rich UI but also needs to run on legacy systems, your options are very limited.
   
   This is where **libTridentUI** can help you!
   
   Every copy of Windows since Windows 95 ships with mshtml.dll (the Trident/IE engine) that is capable of rendering HTML.
   
   Not only it's extremely tiny, but it also ships together with Windows so you don't even have to install it additionally! It's already sitting there inside **EVERY** single Windows machine around the world just waiting to be called.
   
   And that's exactly what **libTridentUI** is doing: An extremely lightweight **single header** library that calls mshtml.dll to render the UI of your software with zero DLL dependencies beyond what Windows already provides!

2. **Maximum compatibility assurance**
   
   **libTridentUI** is built on Microsoft Internet Explorer’s Trident engine, capable of offering absolute top-tier compatibility!!!
   
   * [x] Windows 2000 SP4 (⚠ Untested but expected to work)
   * [x] Windows XP SP3
   * [x] Windows Vista SP1
   * [x] Windows 7 SP1
   * [x] Windows 8
   * [x] Windows 8.1
   * [x] Windows 10 2004
   * [x] Windows 11 21H2
   * [ ] Linux 5+ with WineHQ (Will be supported in near future)
   * [ ] MacOS 10+ with CrossOver (Will be supported in near future)
   * [ ] Android 11+ with Winlator (Will be supported in near future)
   
   You can now create browser based programs that works across all versions of Windows spanning a period of 20 years! Isn't that beautiful?

3. **Easy to use**
   
   **libTridentUI** wrapped all of the complex COM interfaces needed to host mshtml.dll into a clean and flat C-Style API that you can quickly learn in just a few hours!
   
   No need to create any class or struct, no need to implement any COM interface yourself, **libTridentUI** has already done all of the heavy lifting for you.

4. **Multifunctional**
   
   Other libraries might only provide you a web browser window that can only run the slow as hell JavaScript and nothing else, but **libTridentUI** provides you a fully working IDispatch implementation that allows you to call C++ functions from JS and call JS functions from C++ with zero effort!
   
   In addition, **libTridentUI** also supports embedding custom ActiveX controls into your webpage! Instead of the 0.1 fps JS Canvas, you now have access to a GDI compatible hWnd that you can render complex graphics with lightning speed using C++ via GDI Plus, DirectX, or OpenGL.
   
   The ActiveX component registration is done completely in RAM, no registry writes, no uninstallation leftover!

## Supported features

### Window hosting
- [x] Create standalone Trident windows (WS_OVERLAPPEDWINDOW)
- [x] Create child Trident windows (WS_CHILD, embed inside other windows)
- [x] Create popup Trident windows (WS_POPUP)
- [x] Custom extended window styles (WS_EX_*)
- [x] Custom WndProc for host window messages
- [x] Per-window user data storage
- [x] Multiple simultaneous windows
- [x] Proper window lifecycle (WM_NCDESTROY cleanup)

### Navigation
- [x] Navigate to URL (http://, https://, file://)
- [x] Navigate to embedded resource (res:// protocol)
- [x] Navigate to HTML string (document.write injection)
- [x] Access IWebBrowser2 directly for advanced control

### JavaScript ↔ C++ bridge
- [x] Call C++ functions from JS (window.external.MyFunc)
- [x] Call JS functions from C++ (ExecScript)
- [x] Pass arguments: int, double, string, bool
- [x] Return values from C++ to JS
- [x] Return values from JS to C++
- [x] Async callbacks from C++ to JS
- [ ] Create & parse JS arrays (comming soon)
- [ ] Create & parse JS objects (comming soon)
- [ ] Create & parse JS anonymous functions (comming soon)

### DOM manipulation from C++
- [x] Get element by ID
- [x] Get/set innerHTML
- [x] Get/set innerText
- [x] Set outerHTML
- [x] Get/set attributes
- [x] Set CSS class
- [x] Set inline style
- [x] Convenience functions (SetElementHTML, SetElementText by ID)
- [x] Create/remove elements from C++
- [ ] Query elements by class/tag/CSS selector
- [ ] Event listener management from C++

### ActiveX control creation
- [x] Register custom control classes (in-memory, no registry writes)
- [x] Deterministic CLSID generation from name (MD5 hash)
- [x] Embed controls in HTML via `<object>` tag
- [x] Custom WndProc for control messages (WM_PAINT, WM_CREATE, etc.)
- [x] Per-control-class and per-instance user data
- [x] Automatic HIMETRIC ↔ pixel size conversion

### ActiveX IDispatch (JS ↔ control methods)
- [x] Bind methods callable from JS (canvas1.Add(3,7))
- [x] Bind readable/writable properties (canvas1.Counter)
- [x] Getter-only and setter-only properties
- [x] IDispatch argument passing (int, string, double, bool, IDispatch)
- [ ] Optional/default parameter values
- [ ] Automatic cyclecount conversion between JS and C++ types

### ActiveX event system
- [x] Register named events with parameter count
- [x] Fire events from C++ to JS with typed arguments
- [x] Fire events by DISPID or by name
- [x] Native IE attachEvent() support via ITypeInfo/IProvideClassInfo2
- [x] Coclass + dispatch dual-layer type information
- [x] Query expected argument count per event
- [x] Re-entrant safe event firing (sink snapshot with AddRef)
- [ ] detachEvent() from C++ side
- [ ] Custom event object (IE uses plain DISPPARAMS, not DOM Event)

### OLE Drag & Drop
- [x] Drop target: accept dragged data onto controls
- [x] Drop source: initiate drag from controls
- [x] DragEnter / DragOver / DragLeave / Drop callbacks
- [x] Text data transfer (CF_UNICODETEXT)
- [x] File drop from Explorer (CF_HDROP)
- [x] CreateTextDataObject helper
- [x] StartDragDrop helper
- [ ] Custom clipboard formats
- [ ] Drag image / visual feedback customization
- [ ] IDropTargetHelper integration

### IDocHostUIHandler customization
- [x] Context menu control (block or allow default IE menu)
- [x] Host UI flags (3D border, DPI awareness, themes)
- [x] External object (window.external) injection
- [x] Keyboard accelerator forwarding
- [ ] Custom context menu implementation from C++
- [ ] Custom error page handling (IDocHostShowUI)
- [ ] Download manager integration (IDownloadManager)
- [ ] Custom security manager (IInternetSecurityManager)

### Architecture
- [x] Single header library (header-only, no .lib/.dll needed)
- [x] C-style flat API (no classes/inheritance needed by user)
- [x] Thread safety annotations (STA thread requirement documented)
- [x] Re-entrancy protection (critical sections on IDispatch maps)
- [x] Reference-counted COM objects with InterlockedIncrement/Decrement
- [x] Doubly-linked window list for O(1) removal
- [x] Separate message loop modes (full loop vs. per-message processing)
- [ ] Thread-safe API (all calls must be from STA thread)
- [ ] Multiple STA thread support (one Trident instance per thread)
