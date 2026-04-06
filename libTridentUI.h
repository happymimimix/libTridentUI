#pragma once
/*
libTridentUI - v0.4.1 Beta
An extremely lightweight Chrome Embedded Framework alternative for legacy operating systems.
Works perfectly on Microsoft Windows XP SP3!
Made possible by:
Happy_mimimix
Claude Opus 4.6
ChatGPT 5.2 Thinking
Gemini Pro

Thread safety notice:
All UI calls must be made from the same STA thread that called TridentInit().
This is a COM/OLE requirement, not a library limitation.
*/

/*
HOW THIS WORKS (the 30-second version):
Every Windows machine since Windows 95 has an HTML rendering engine built in: mshtml.dll (aka "Trident",
the engine behind Internet Explorer). This library creates a Win32 window and stuffs an instance
of that engine inside it, like putting a browser tab into your app.

The trick is that IE's engine is an "ActiveX control" - a type of COM object that can be embedded
inside another window. To host it, we need to implement a bunch of COM interfaces that the control
calls back into. Think of it like a plugin system: our app is the host, mshtml is the plugin.
The plugin says "I need a window to draw in" and we respond with our HWND. It says "give me a
frame for my toolbar" and we give it a dummy one. And so on, for about 20 different interfaces.

The key interfaces WE implement (as the "container" / "host"):
	IOleClientSite		- "I'm the app hosting you, here's how to talk to me"
	IOleInPlaceSiteEx	- "Here's the window you can draw in"
	IOleControlSite		- "I can handle your keyboard accelerators"
	IDocHostUIHandler	- "I control the right-click menu, scrollbars, etc."
	IDispatch			- "I handle navigation events (like page load complete)"
	IOleContainer		- "I can enumerate the objects I'm hosting"

The key interfaces IE GIVES US (that we call into):
	IWebBrowser2		- "Navigate to URLs, get the document, etc."
	IOleObject			- "Activate me, close me, get my status"
	IOleInPlaceObject	- "Here's how to resize and position me"
	IHTMLDocument2/3	- "Here's the DOM of the loaded page"

The JS <-> C++ bridge works through "window.external":
	When JavaScript calls window.external.MyFunc(1, 2), IE looks up our IDocHostUIHandler,
	calls GetExternal() to get an IDispatch, calls GetIDsOfNames("MyFunc") to get a numeric ID,
	then calls Invoke(id, args) which runs our C++ callback.

The ActiveX control creation works the other way around:
	We register a class factory in memory (no registry), so when IE encounters
	<object classid="clsid:...">, it calls our factory, which creates a COM object that
	implements IOleObject/IDispatch/etc. IE embeds it the same way apps embed IE itself.
*/

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <windowsx.h>
#include <wincrypt.h>
#include <ole2.h>
#include <oleidl.h>
#include <oaidl.h>
#include <oleauto.h>
#include <exdisp.h>
#include <exdispid.h>
#include <mshtml.h>
#include <mshtmhst.h>
#include <mshtmdid.h>
#include <functional>
#include <map>
#include <string>
#include <stdexcept>
#include <ocidl.h>
#include <objsafe.h>
#include <vector>
#include <algorithm>
#include <intrin.h>
#define LONG_MAX_PATH 0x0FFF

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "uuid.lib")
#pragma comment(lib, "user32.lib")

// Public APIs
struct TridentWindowData_;
typedef TridentWindowData_* hTrident;
typedef IHTMLElement* hElement;
typedef IHTMLDocument2* hDocument;
typedef IHTMLDocument3* hDocument2;
struct ControlInstance_;
typedef ControlInstance_* hControl;
struct ControlReg_;
typedef ControlReg_* hControlClass;
enum InsertLocation {BEFORE_BEGIN,AFTER_BEGIN,BEFORE_END,AFTER_END};

// Callback signatures - both receive raw DISPPARAMS* (COM's reverse-ordered argument array).
// 
// DISPPARAMS argument convention:
// 
// COM passes function arguments in REVERSE order in DISPPARAMS::rgvarg:
//  	rgvarg[cArgs-1] = first argument
//  	rgvarg[cArgs-2] = second argument
//  	rgvarg[0]       = last argument
// 
// This is the same convention in BOTH directions:
//  	JS -> C++ (method calls): rgvarg is reverse-ordered, your callback reads them
//  	C++ -> JS (FireEvent): you pass args in reverse order, JS receives them correctly
// 
// Example - JS calls window.external.Add(10, 20):
//  	dp->cArgs = 2
//  	dp->rgvarg[0].lVal = 20  (last arg)
//  	dp->rgvarg[1].lVal = 10  (first arg)
// 
// Example - C++ fires event with args (42, "hello"):
//  	VARIANT args[2];
//  	args[0].vt = VT_BSTR; args[0].bstrVal = SysAllocString(L"hello"); // last arg first
//  	args[1].vt = VT_I4; args[1].lVal = 42; // first arg last
//  	FireEvent(c, evtId, args);
using MethodCallback = std::function<HRESULT(DISPPARAMS*, VARIANT*)>;
using ControlMethodCallback = std::function<HRESULT(hControl, DISPPARAMS*, VARIANT*)>;

typedef LRESULT(*TridentWndProc)(hTrident h, HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool* implemented);
typedef LRESULT(*ControlWndProc)(hControl c, HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool* implemented);
typedef HRESULT(*PropertyGetter)(hControl c, VARIANT* result);
typedef HRESULT(*PropertySetter)(hControl c, VARIANT value);
typedef HRESULT(*DropEnterCallback)(hControl c, IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);
typedef HRESULT(*DropOverCallback)(hControl c, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);
typedef HRESULT(*DropLeaveCallback)(hControl c);
typedef HRESULT(*DropCallback)(hControl c, IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);

inline void TridentInit();
inline void TridentShutdown();
inline bool IsTridentHostWindow(HWND hwnd);
inline hTrident FindNearestTridentHostFromHWND(HWND hwnd);
inline void TridentRunMessageLoop();
inline void TridentProcessMessage(MSG& msg, bool* is_eaten);
inline hTrident NewTridentWindow(const wchar_t* title, int x, int y, int w, int h, HWND owner = NULL, DWORD style = WS_OVERLAPPEDWINDOW, DWORD styleEX = NULL);
inline void CloseTridentWindow(hTrident h);
inline HWND GetTridentHWND(hTrident h);
inline IWebBrowser2* GetTridentBrowser(hTrident h);
inline void BindFunction(hTrident h, const wchar_t* name, MethodCallback cb);
inline void SetWndProc(hTrident h, TridentWndProc proc);
inline void SetUserData(hTrident h, void* data);
inline void* GetUserData(hTrident h);
inline void NavigateTo(hTrident h, const wchar_t* url);
inline void NavigateToRes(hTrident h, const wchar_t* resName);
inline void NavigateToHTML(hTrident h, const wchar_t* html);
inline hDocument GetDocument(hTrident h);
inline hDocument2 GetDocument2(hDocument h);

inline VARIANT EvaluateExpression(hTrident h, const wchar_t* expr);
inline VARIANT InvokeGlobalFunction(hTrident h, const wchar_t* jsfunc, VARIANT* args = NULL, unsigned int argc = 0);
inline VARIANT InvokeFunctionByPath(hTrident h, wchar_t** path, unsigned int pathlen, VARIANT* args = NULL, unsigned int argc = 0);
inline void ExecuteScript(hTrident h, const wchar_t* jscript);
inline HRESULT GetIDsOfNamesEx(IDispatch* disp, BSTR name, DISPID* dispid, bool create = false);
inline VARIANT DispatchGetNamedProperty(IDispatch* disp, const wchar_t* name);
inline void DispatchSetNamedProperty(IDispatch* disp, const wchar_t* name, const VARIANT* value);
inline VARIANT CreateObject(hTrident h, const wchar_t* classname, VARIANT* args = NULL, unsigned int argc = 0);
inline VARIANT CreateEmptyArray(hTrident h);
inline VARIANT CreateEmptyObject(hTrident h);
inline VARIANT GetDispatchIndex(VARIANT* arr, unsigned int index);
inline void SetDispatchIndex(VARIANT* arr, unsigned int index, const VARIANT* value);
inline VARIANT GetDispatchSize(VARIANT* arr);
inline void DispatchResize(VARIANT* arr, unsigned int newsize);
inline VARIANT GetDispatchProperty(VARIANT* obj, const wchar_t* prop);
inline void SetDispatchProperty(VARIANT* obj, const wchar_t* prop, const VARIANT* value);
inline VARIANT CallDispatchFunction(VARIANT* func, VARIANT* args = NULL, unsigned int argc = 0);
inline VARIANT GetGlobalFunction(hTrident h, const wchar_t* funcName);
inline VARIANT GetExternalFunction(hTrident h, const wchar_t* funcName);
inline VARIANT GetFunctionByPath(hTrident h, wchar_t** path, unsigned int pathlen);
inline VARIANT CreateFunction(hTrident h, wchar_t** args, unsigned int argc, const wchar_t* impl);

inline hElement GetElement(hTrident h, const wchar_t* id);
inline void SetHTML(hElement e, const wchar_t* html);
inline std::wstring GetHTML(hElement e);
inline void SetText(hElement e, const wchar_t* text);
inline std::wstring GetText(hElement e);
inline void SetOuterHTML(hElement e, const wchar_t* html);
inline void SetClass(hElement e, const wchar_t* cls);
inline void SetStyle(hElement e, const wchar_t* css);
inline void SetAttr(hElement e, const wchar_t* name, const wchar_t* val);
inline std::wstring GetAttr(hElement e, const wchar_t* name);
inline void SetElementHTML(hTrident h, const wchar_t* id, const wchar_t* html);
inline std::wstring GetElementHTML(hTrident h, const wchar_t* id);
inline void SetElementText(hTrident h, const wchar_t* id, const wchar_t* text);
inline std::wstring GetElementText(hTrident h, const wchar_t* id);
inline void InsertHTML(hTrident h, const wchar_t* id, InsertLocation where, const wchar_t* html);
inline void InsertHTMLAtBody(hTrident h, InsertLocation where, const wchar_t* html);
inline void RemoveElement(hTrident h, const wchar_t* id);

inline CLSID CLSIDFromName(const wchar_t* name);
inline std::wstring CLSIDToString(const CLSID& clsid);
inline std::wstring GetControlClassId(hControlClass reg);
inline CLSID GetControlCLSID(hControlClass reg);
inline hControlClass RegisterControl(const wchar_t* name, ControlWndProc proc, void* userData = NULL, DWORD style = NULL);
inline void UnregisterControl(hControlClass reg);
inline HWND GetControlHWND(hControl c);
inline void* GetControlUserData(hControl c);
inline void SetControlUserData(hControl c, void* data);
inline void InvalidateControl(hControl c);
inline hControlClass GetControlClass(hControl c);
inline void BindClassMethod(hControlClass reg, const wchar_t* name, ControlMethodCallback cb);
inline void BindClassProperty(hControlClass reg, const wchar_t* name, PropertyGetter getter, PropertySetter setter);
inline DISPID RegisterEvent(hControlClass reg, const wchar_t* name, unsigned int paramCount = 0);
inline unsigned int GetEventExpectedArgc(hControlClass reg, DISPID eventId);
inline unsigned int GetEventExpectedArgcByName(hControlClass reg, const wchar_t* name);
inline void FireEvent(hControl c, DISPID eventId, VARIANT* args = NULL);
inline void FireEventByName(hControl c, const wchar_t* name, VARIANT* args = NULL);
inline hTrident FindHostFromControl(hControl c);
inline void EnableDragDrop(hControl c, DropCallback onDrop, DropEnterCallback onDragEnter = NULL, DropOverCallback onDragOver = NULL, DropLeaveCallback onDragLeave = NULL);
inline void DisableDragDrop(hControl c);
inline HRESULT StartDragDrop(hControl c, IDataObject* data, DWORD allowedEffects = DROPEFFECT_COPY, DWORD* outEffect = NULL);

// COM implementations
inline const wchar_t* g_trident_wndclass = L"libTridentUI";
inline TridentWindowData_* g_trident_head = NULL;

// ExternalDispatch - this is what JavaScript sees as "window.external"
// 
// When JS calls window.external.Add(1, 2), here's what happens behind the scenes:
//  	1. IE calls GetIDsOfNames(["Add"]) -> we return a numeric DISPID
//  	2. IE calls Invoke(dispid, args=[1,2]) -> we look up the C++ callback and run it
// 
// IDispatch is COM's version of "dynamic method lookup" - like a dictionary of
// function pointers, keyed by name. It's how scripting languages talk to COM objects
// without knowing their C++ class layout at compile time.
// 
// DISPIDs count DOWN from INT32_MAX. This avoids conflicts with well-known DISPIDs
// like CYCLECOUNT_VALUE (0), CYCLECOUNT_CYCLECOUNT (-1), CYCLECOUNT_CYCLETIMER (-5), etc.
// We'll never run out - you'd need 2 billion bindings.
// 
// NOTE: This object lives as a member of OleSite (not heap-allocated separately),
// so its Release() intentionally does NOT delete this. Its lifetime is tied to OleSite.
class ExternalDispatch : public IDispatch {
	LONG m_ref = 0;
	CRITICAL_SECTION m_cs;
	DISPID m_nextId = INT32_MAX;
	std::map<std::wstring, DISPID> m_name2id = {};
	std::map<DISPID, MethodCallback> m_id2func = {};

public:
	ExternalDispatch() { InitializeCriticalSection(&m_cs); }
	~ExternalDispatch() { DeleteCriticalSection(&m_cs); }

	void Bind(const wchar_t* name, MethodCallback cb) {
		EnterCriticalSection(&m_cs);
		if (m_name2id.find(name) != m_name2id.end()) {
			LeaveCriticalSection(&m_cs);
			throw std::runtime_error("Function already exists");
		}
		if (m_nextId <= 0) {
			LeaveCriticalSection(&m_cs);
			throw std::runtime_error("Cannot allocate more function IDs");
		}
		DISPID id = m_nextId--;
		m_name2id[name] = id;
		m_id2func[id] = cb;
		LeaveCriticalSection(&m_cs);
	}

	STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
		if (riid == IID_IUnknown || riid == IID_IDispatch) {
			*ppv = static_cast<IDispatch*>(this);
			AddRef();
			return S_OK;
		}
		*ppv = NULL;
		return E_NOINTERFACE;
	}
	STDMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&m_ref); }
	STDMETHODIMP_(ULONG) Release() override { return InterlockedDecrement(&m_ref); }

	STDMETHODIMP GetTypeInfoCount(UINT* p) override { if (!p) return E_POINTER; *p = 0; return S_OK; }
	STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }

	// GetIDsOfNames - IE calls this to convert "Add" -> numeric DISPID.
	// When JS does window.external.Add(1,2), IE first asks us "what's the ID for 'Add'?"
	STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR* names, UINT cnt, LCID, DISPID* ids) override {
		HRESULT hr = S_OK;
		EnterCriticalSection(&m_cs);
		for (UINT i = 0; i < cnt; i++) {
			auto it = m_name2id.find(names[i]);
			if (it != m_name2id.end()) {
				ids[i] = it->second;
			}
			else {
				ids[i] = DISPID_UNKNOWN;
				hr = DISP_E_UNKNOWNNAME;
			}
		}
		LeaveCriticalSection(&m_cs);
		return hr;
	}

	// Invoke - IE calls this with the DISPID from GetIDsOfNames + the JS arguments.
	// We look up the C++ callback and call it. Note: we copy the callback THEN unlock
	// the critical section BEFORE calling the user's code. This prevents deadlocks if
	// the user's callback tries to bind more functions.
	STDMETHODIMP Invoke(DISPID id, REFIID, LCID, WORD wf, DISPPARAMS* dp, VARIANT* res, EXCEPINFO*, UINT*) override {
		if (!(wf & DISPATCH_METHOD)) return DISP_E_MEMBERNOTFOUND;
		EnterCriticalSection(&m_cs);
		auto it = m_id2func.find(id);
		if (it == m_id2func.end()) { LeaveCriticalSection(&m_cs); return DISP_E_MEMBERNOTFOUND; }
		auto fn = it->second;
		LeaveCriticalSection(&m_cs);
		if (res) VariantInit(res);
		return fn(dp, res);
	}
};

// OleSite - the "container" that hosts the IE browser control.
// 
// This is the biggest and most complex class in the library. It implements 7 COM interfaces
// that IE's engine needs from its host. Think of it as the "other half of a conversation":
// IE says "I need X" and OleSite responds with the appropriate answer.
// 
// Here's what each interface does:
// 
//  	IOleClientSite			The most basic hosting interface. IE calls this to ask
//  							"who's hosting me?" and "where should I save my data?"
//  							We mostly return E_NOTIMPL because we don't save anything.
// 
//  	IOleInPlaceSiteEx		Handles the visual embedding. IE asks for the HWND to draw
//  							in, and reports activation/deactivation state changes.
// 
//  	IOleControlSite			Keyboard and focus management. IE asks us to translate
//  							accelerator keys and track focus changes.
// 
//  	IObjectWithSite			A generic "set my parent" interface. IE uses this to
//  							establish a bidirectional link between itself and its host.
// 
//  	IOleContainer			IE asks "what other objects are you hosting?" We return
//  							a list containing just the one browser instance.
// 
//  	IDocHostUIHandler		The UI customization interface. This is where we:
//  							- Block the right-click context menu (ShowContextMenu -> S_OK)
//  							- Set DPI aware (GetHostInfo -> DOCHOSTUIFLAG_DPI_AWARE)
//  							- Provide the window.external object (GetExternal -> &ext)
// 
//  	IDispatch				We implement IDispatch ourselves to receive navigation
//  							events from the browser (CYCLECOUNT_CYCLETIMER, DISPID_DOCUMENTCOMPLETE).
//  							When the page finishes loading, we hook up our UI handler.
class OleSite :
	public IOleClientSite,
	public IOleInPlaceSiteEx,
	public IOleControlSite,
	public IObjectWithSite,
	public IOleContainer,
	public IDocHostUIHandler,
	public IDispatch
{
public:
	ExternalDispatch ext;

	OleSite() : m_hwnd(NULL), m_pUnkSite(NULL), m_pView(NULL), m_pInPlace(NULL), m_pOle(NULL), m_pBrowser(NULL), m_cookie(NULL)
	{
		InterlockedIncrement(&m_ref);
	}

	~OleSite() {
		Destroy();
		if (m_pUnkSite) m_pUnkSite->Release();
		if (m_pView) m_pView->Release();
		if (m_pInPlace) m_pInPlace->Release();
	}

	IWebBrowser2* GetBrowser() { return m_pBrowser; }

	// SetupUIHandler - hooks our IDocHostUIHandler into the loaded document.
	// We have to do this AFTER the page loads because the document object changes on
	// every navigation. The ICustomDoc interface is MSHTML's way of saying "let me swap
	// in a different UI handler for this document".
	void SetupUIHandler() {
		if (!m_pBrowser) return;
		IDispatch* d = NULL;
		m_pBrowser->get_Document(&d);
		if (!d) return;
		ICustomDoc* cd = NULL;
		if (SUCCEEDED(d->QueryInterface(IID_ICustomDoc, (void**)&cd))) {
			cd->SetUIHandler(static_cast<IDocHostUIHandler*>(this));
			cd->Release();
		}
		d->Release();
	}

	// Create - the big initialization sequence. This is where we actually create the
	// IE browser control and wire everything together. Here's the step-by-step:
	bool Create(HWND hwnd) {
		m_hwnd = hwnd;

		// Step 1: Create the WebBrowser COM object. This is the IE engine itself.
		// CLSID_WebBrowser is {8856F961-340A-11D0-A96B-00C04FD705A2} - it lives in
		// ieframe.dll and has been there since IE3.
		HRESULT hr = CoCreateInstance(CLSID_WebBrowser, NULL, CLSCTX_ALL, IID_IOleObject, (void**)&m_pOle);
		if (FAILED(hr)) return false;

		// Step 2: Tell IE "I'm your host". This lets IE call back into us for
		// window management, keyboard handling, etc.
		m_pOle->SetClientSite(static_cast<IOleClientSite*>(this));

		// Step 3: Initialize IE's persistent state from scratch (empty, no saved data).
		// Without this, IE might try to load from a stream that doesn't exist.
		IPersistStreamInit* psi = NULL;
		m_pOle->QueryInterface(IID_IPersistStreamInit, (void**)&psi);
		if (psi) {
			psi->InitNew();
			psi->Release();
		}

		// Step 4: Get IE's drawing interface for ShowObject() redraws.
		m_pOle->QueryInterface(IID_IViewObject, (void**)&m_pView);
		m_pOle->SetHostNames(OLESTR("TridentUI"), NULL);

		// Step 5: Activate IE "in place" - this makes it create its child windows inside
		// our HWND and start rendering. OLEIVERB_INPLACEACTIVATE is the command that says
		// "go live inside this rectangle".
		RECT rc;
		GetClientRect(hwnd, &rc);
		hr = m_pOle->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, static_cast<IOleClientSite*>(this), 0, hwnd, &rc);
		if (FAILED(hr)) { Destroy(); return false; }

		// Step 6: Establish a bidirectional site link. Some COM objects use IObjectWithSite
		// as an alternative to IOleClientSite for the "who's my parent?" question.
		IObjectWithSite* pSite = NULL;
		m_pOle->QueryInterface(IID_IObjectWithSite, (void**)&pSite);
		if (pSite) {
			pSite->SetSite(static_cast<IOleClientSite*>(this));
			pSite->Release();
		}

		// Step 7: Get the IWebBrowser2 interface - this is our main handle to the browser.
		// Through it we can navigate to URLs, get the document object, execute script, etc.
		m_pOle->QueryInterface(IID_IWebBrowser2, (void**)&m_pBrowser);
		if (!m_pBrowser) { Destroy(); return false; }

		// Step 8: Subscribe to browser events (like CYCLECOUNT_CYCLETIMER)
		// so we know when navigation completes and can hook up our UI handler.
		ConnectEvents(TRUE);
		return true;
	}

	// Destroy - disconnects everything and releases COM references.
	// This must be called before the host window is destroyed, because IE's child windows
	// are parented to our HWND. Calling Close(OLECLOSE_NOSAVE) tells IE to deactivate
	// and tear down its UI. SetClientSite(NULL) breaks the back-pointer so IE won't
	// try to call us after we're gone.
	void Destroy() {
		ConnectEvents(FALSE);
		if (m_pBrowser) { m_pBrowser->Release(); m_pBrowser = NULL; }
		if (m_pOle) {
			IObjectWithSite* pSite = NULL;
			m_pOle->QueryInterface(IID_IObjectWithSite, (void**)&pSite);
			if (pSite) { pSite->SetSite(NULL); pSite->Release(); }
			m_pOle->Close(OLECLOSE_NOSAVE);
			m_pOle->SetClientSite(NULL);
			m_pOle->Release();
			m_pOle = NULL;
		}
	}

	void Resize(int cx, int cy) {
		if (m_pInPlace) {
			RECT rc = { 0, 0, cx, cy };
			m_pInPlace->SetObjectRects(&rc, &rc);
		}
	}

	// QueryInterface - COM's way of asking "do you support interface X?"
	// IE calls this constantly during initialization to discover what we can do.
	// Each interface we return tells IE about a different capability we offer.
	STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
		*ppv = NULL;
		if (riid == IID_IUnknown) *ppv = static_cast<IOleWindow*>(this);
		else if (riid == IID_IOleClientSite) *ppv = static_cast<IOleClientSite*>(this);
		else if (riid == IID_IOleInPlaceSiteEx) *ppv = static_cast<IOleInPlaceSiteEx*>(this);
		else if (riid == IID_IOleInPlaceSite) *ppv = static_cast<IOleInPlaceSite*>(this);
		else if (riid == IID_IOleWindow) *ppv = static_cast<IOleWindow*>(this);
		else if (riid == IID_IOleControlSite) *ppv = static_cast<IOleControlSite*>(this);
		else if (riid == IID_IOleContainer) *ppv = static_cast<IOleContainer*>(this);
		else if (riid == IID_IObjectWithSite) *ppv = static_cast<IObjectWithSite*>(this);
		else if (riid == IID_IDocHostUIHandler) *ppv = static_cast<IDocHostUIHandler*>(this);
		else if (riid == IID_IDispatch) *ppv = static_cast<IDispatch*>(this);
		if (*ppv) { AddRef(); return S_OK; }
		return E_NOINTERFACE;
	}
	STDMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&m_ref); }
	STDMETHODIMP_(ULONG) Release() override { LONG r = InterlockedDecrement(&m_ref); if (!r) delete this; return r; }

	// ---- IOleClientSite ----
	// These methods let IE ask us about its hosting environment.
	// Most are stubs because we're a simple host - we don't save documents,
	// don't support monikers (persistent object names), etc.
	STDMETHODIMP SaveObject() override { return E_NOTIMPL; }
	STDMETHODIMP GetMoniker(DWORD, DWORD, IMoniker** p) override { if (!p) return E_POINTER; *p = NULL; return E_NOTIMPL; }
	STDMETHODIMP GetContainer(IOleContainer** p) override {
		return QueryInterface(IID_IOleContainer, (void**)p);
	}
	// ShowObject - IE says "I've just activated, you might want to redraw me"
	STDMETHODIMP ShowObject() override {
		if (!m_hwnd || !m_pView) return E_FAIL;
		HDC hDC = ::GetDC(m_hwnd);
		RECT rc; GetClientRect(m_hwnd, &rc);
		m_pView->Draw(DVASPECT_CONTENT, -1, NULL, NULL, NULL, hDC, (RECTL*)&rc, (RECTL*)&rc, NULL, NULL);
		::ReleaseDC(m_hwnd, hDC);
		return S_OK;
	}
	STDMETHODIMP OnShowWindow(BOOL) override { return E_NOTIMPL; }
	STDMETHODIMP RequestNewObjectLayout() override { return E_NOTIMPL; }

	// ---- IObjectWithSite ----
	// A generic parent-child link. IE uses this to find us later.
	STDMETHODIMP SetSite(IUnknown* p) override {
		if (m_pUnkSite) m_pUnkSite->Release();
		m_pUnkSite = p;
		if (p) p->AddRef();
		return S_OK;
	}
	STDMETHODIMP GetSite(REFIID riid, void** pp) override {
		if (!m_pUnkSite) { *pp = NULL; return E_FAIL; }
		return m_pUnkSite->QueryInterface(riid, pp);
	}

	// ---- IOleInPlaceSite ----
	// This is where the magic happens: IE asks "can I activate inside your window?"
	// and we say yes, give it our HWND, and tell it how big the available area is.
	STDMETHODIMP GetWindow(HWND* p) override { if (!p) return E_POINTER; *p = m_hwnd; return S_OK; }
	STDMETHODIMP ContextSensitiveHelp(BOOL) override { return E_NOTIMPL; }
	STDMETHODIMP CanInPlaceActivate() override { return S_OK; }
	STDMETHODIMP OnInPlaceActivate() override {
		BOOL dummy = FALSE;
		return OnInPlaceActivateEx(&dummy, 0);
	}
	STDMETHODIMP OnUIActivate() override { return S_OK; }

	// GetWindowContext - IE asks "give me a frame, a UI window, and the rectangles
	// I can draw in". We provide a dummy InlineFrame (which is basically a no-op
	// implementation of IOleInPlaceFrame), and tell IE it can use our entire client area.
	STDMETHODIMP GetWindowContext(IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc, LPRECT pPos, LPRECT pClip, LPOLEINPLACEFRAMEINFO pFI) override {
		if (!ppFrame || !ppDoc || !pPos || !pClip || !pFI) return E_POINTER;
		*ppFrame = new InlineFrame(m_hwnd);
		*ppDoc = NULL; // No separate UI window (we're not an MDI app)
		GetClientRect(m_hwnd, pPos);
		*pClip = *pPos; // Clipping rect = position rect (no clipping)
		pFI->cb = sizeof(OLEINPLACEFRAMEINFO);
		pFI->fMDIApp = FALSE;
		pFI->hwndFrame = m_hwnd;
		pFI->haccel = NULL; // No accelerator table
		pFI->cAccelEntries = 0;
		return S_OK;
	}
	STDMETHODIMP Scroll(SIZE) override { return E_NOTIMPL; }
	STDMETHODIMP OnUIDeactivate(BOOL) override { return S_OK; }
	STDMETHODIMP OnInPlaceDeactivate() override { return OnInPlaceDeactivateEx(TRUE); }
	STDMETHODIMP DiscardUndoState() override { return E_NOTIMPL; }
	STDMETHODIMP DeactivateAndUndo() override { return E_NOTIMPL; }
	STDMETHODIMP OnPosRectChange(LPCRECT) override { return E_NOTIMPL; }

	// OnInPlaceActivateEx - IE is going live inside our window.
	// OleLockRunning prevents the OLE object from being garbage-collected while active.
	STDMETHODIMP OnInPlaceActivateEx(BOOL*, DWORD) override {
		if (!m_hwnd || !m_pOle) return E_UNEXPECTED;
		HRESULT hr = m_pOle->QueryInterface(IID_IOleInPlaceObject, (void**)&m_pInPlace);
		if (FAILED(hr)) return hr;
		OleLockRunning(m_pOle, TRUE, FALSE);
		RECT rc; GetClientRect(m_hwnd, &rc);
		m_pInPlace->SetObjectRects(&rc, &rc);
		return S_OK;
	}
	STDMETHODIMP OnInPlaceDeactivateEx(BOOL) override {
		if (m_pOle) OleLockRunning(m_pOle, FALSE, FALSE);
		if (m_pInPlace) {
			m_pInPlace->Release();
			m_pInPlace = NULL;
		}
		return S_OK;
	}
	STDMETHODIMP RequestUIActivate() override { return S_OK; }
	STDMETHODIMP OnControlInfoChanged() override { return S_OK; }
	STDMETHODIMP LockInPlaceActive(BOOL) override { return S_OK; }
	STDMETHODIMP GetExtendedControl(IDispatch** p) override {
		if (!p || !m_pOle) return E_POINTER;
		return m_pOle->QueryInterface(IID_IDispatch, (void**)p);
	}
	STDMETHODIMP TransformCoords(POINTL*, POINTF*, DWORD) override { return E_NOTIMPL; }
	STDMETHODIMP TranslateAccelerator(MSG*, DWORD) override { return S_FALSE; }
	STDMETHODIMP OnFocus(BOOL) override { return S_OK; }
	STDMETHODIMP ShowPropertyFrame() override { return E_NOTIMPL; }
	STDMETHODIMP EnumObjects(DWORD, IEnumUnknown** pp) override {
		if (!pp || !m_pOle) return E_POINTER;
		*pp = new InlineEnum(m_pOle);
		return S_OK;
	}
	STDMETHODIMP LockContainer(BOOL) override { return S_OK; }
	STDMETHODIMP ParseDisplayName(IBindCtx*, LPOLESTR, ULONG*, IMoniker**) override { return E_NOTIMPL; }

	// ---- IDocHostUIHandler ----
	// This is our main UI customization point. IE calls these to ask how we want
	// the browser to look and behave.

	// ShowContextMenu - return S_FALSE to show the default IE menu.
	// Return S_OK if you want to BLOCK the default right-click menu.
	STDMETHODIMP ShowContextMenu(DWORD, POINT*, IUnknown*, IDispatch*) override { return S_FALSE; }

	// GetHostInfo - we set DOCHOSTUIFLAG_DPI_AWARE to make DPI scaling correct.
	STDMETHODIMP GetHostInfo(DOCHOSTUIINFO* p) override {
		if (p) {
			p->cbSize = sizeof(DOCHOSTUIINFO);
			p->dwFlags = DOCHOSTUIFLAG_DPI_AWARE;
			p->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT;
		}
		return S_OK;
	}
	STDMETHODIMP ShowUI(DWORD, IOleInPlaceActiveObject*, IOleCommandTarget*, IOleInPlaceFrame*, IOleInPlaceUIWindow*) override { return S_FALSE; }
	STDMETHODIMP HideUI() override { return S_OK; }
	STDMETHODIMP UpdateUI() override { return S_OK; }
	STDMETHODIMP EnableModeless(BOOL) override { return S_OK; }
	STDMETHODIMP OnDocWindowActivate(BOOL) override { return S_OK; }
	STDMETHODIMP OnFrameWindowActivate(BOOL) override { return S_OK; }
	STDMETHODIMP ResizeBorder(LPCRECT, IOleInPlaceUIWindow*, BOOL) override { return S_FALSE; }
	STDMETHODIMP TranslateAccelerator(LPMSG, const GUID*, DWORD) override { return S_FALSE; }
	STDMETHODIMP GetOptionKeyPath(LPOLESTR*, DWORD) override { return S_FALSE; }
	STDMETHODIMP GetDropTarget(IDropTarget*, IDropTarget**) override { return E_NOTIMPL; }

	// GetExternal - THE KEY METHOD for the JS bridge!
	// When JavaScript calls "window.external", IE calls this to get the IDispatch
	// that represents the external object. We return our ExternalDispatch, which
	// maps JS function calls to C++ callbacks registered via BindFunction().
	STDMETHODIMP GetExternal(IDispatch** pp) override {
		if (!pp) return E_POINTER;
		*pp = &ext;
		ext.AddRef();
		return S_OK;
	}
	STDMETHODIMP TranslateUrl(DWORD, LPWSTR, LPWSTR*) override { return E_NOTIMPL; }
	STDMETHODIMP FilterDataObject(IDataObject*, IDataObject**) override { return E_NOTIMPL; }

	// ---- IDispatch (for browser events) ----
	// We subscribe to DWebBrowserEvents2 to know when navigation completes.
	// When IE fires DISPID_DOCUMENTCOMPLETE, we call SetupUIHandler() to hook
	// our IDocHostUIHandler into the new document. This is necessary because
	// the document object changes on every navigation.
	STDMETHODIMP GetTypeInfoCount(UINT* p) override { if (!p) return E_POINTER; *p = 0; return S_OK; }
	STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
	STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override { return E_NOTIMPL; }
	STDMETHODIMP Invoke(DISPID id, REFIID riid, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) override {
		if (riid != IID_NULL) return E_INVALIDARG;
		if (id == DISPID_NAVIGATECOMPLETE2 || id == DISPID_DOCUMENTCOMPLETE) SetupUIHandler();
		return S_OK;
	}

private:
	// ConnectEvents - subscribe/unsubscribe to IE's navigation events.
	// IE exposes events through "connection points" - a COM pattern where you
	// give IE an IDispatch, and it calls Invoke() on it whenever something happens.
	// We connect to DWebBrowserEvents2 to catch DISPID_DOCUMENTCOMPLETE.
	void ConnectEvents(BOOL advise) {
		if (!m_pBrowser) return;
		IConnectionPointContainer* cpc = NULL;
		if (FAILED(m_pBrowser->QueryInterface(IID_IConnectionPointContainer, (void**)&cpc))) return;
		IConnectionPoint* cp = NULL;
		if (SUCCEEDED(cpc->FindConnectionPoint(DIID_DWebBrowserEvents2, &cp))) {
			if (advise) {
				cp->Advise(static_cast<IDispatch*>(this), &m_cookie);
			}
			else if (m_cookie) {
				cp->Unadvise(m_cookie);
				m_cookie = NULL;
			}
			cp->Release();
		}
		cpc->Release();
	}
	class InlineFrame : public IOleInPlaceFrame {
		ULONG m_ref = 0;
		HWND m_hwnd = NULL;
	public:
		InlineFrame(HWND h) : m_hwnd(h) { InterlockedIncrement(&m_ref); }
		STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
			if (riid == IID_IUnknown || riid == IID_IOleInPlaceFrame ||
				riid == IID_IOleWindow || riid == IID_IOleInPlaceUIWindow)
			{
				*ppv = static_cast<IOleInPlaceFrame*>(this);
				AddRef();
				return S_OK;
			}
			*ppv = NULL;
			return E_NOINTERFACE;
		}
		STDMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&m_ref); }
		STDMETHODIMP_(ULONG) Release() override { LONG r = InterlockedDecrement(&m_ref); if (!r) delete this; return r; }
		STDMETHODIMP GetWindow(HWND* p) override { if (!p) return E_POINTER; *p = m_hwnd; return S_OK; }
		STDMETHODIMP ContextSensitiveHelp(BOOL) override { return S_OK; }
		STDMETHODIMP GetBorder(LPRECT) override { return S_OK; }
		STDMETHODIMP RequestBorderSpace(LPCBORDERWIDTHS) override { return INPLACE_E_NOTOOLSPACE; }
		STDMETHODIMP SetBorderSpace(LPCBORDERWIDTHS) override { return S_OK; }
		STDMETHODIMP SetActiveObject(IOleInPlaceActiveObject*, LPCOLESTR) override { return S_OK; }
		STDMETHODIMP InsertMenus(HMENU, LPOLEMENUGROUPWIDTHS) override { return S_OK; }
		STDMETHODIMP SetMenu(HMENU, HOLEMENU, HWND) override { return S_OK; }
		STDMETHODIMP RemoveMenus(HMENU) override { return S_OK; }
		STDMETHODIMP SetStatusText(LPCOLESTR) override { return S_OK; }
		STDMETHODIMP EnableModeless(BOOL) override { return S_OK; }
		STDMETHODIMP TranslateAccelerator(LPMSG, WORD) override { return S_FALSE; }
	};
	class InlineEnum : public IEnumUnknown {
		ULONG m_ref = 0;
		LONG m_pos = 0;
		IUnknown* m_ptr = NULL;
	public:
		InlineEnum(IUnknown* p) : m_pos(0), m_ptr(p) {
			InterlockedIncrement(&m_ref);
			m_ptr->AddRef();
		}
		~InlineEnum() { m_ptr->Release(); }
		STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
			if (riid == IID_IUnknown || riid == IID_IEnumUnknown)
			{
				*ppv = static_cast<IEnumUnknown*>(this);
				AddRef();
				return S_OK;
			}
			*ppv = NULL;
			return E_NOINTERFACE;
		}
		STDMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&m_ref); }
		STDMETHODIMP_(ULONG) Release() override { LONG r = InterlockedDecrement(&m_ref); if (!r) delete this; return r; }
		STDMETHODIMP Next(ULONG celt, IUnknown** p, ULONG* f) override {
			if (!p) return E_POINTER;
			ULONG fetched = 0;
			while (fetched < celt && m_pos < 1) {
				p[fetched] = m_ptr; p[fetched]->AddRef();
				fetched++; m_pos++;
			}
			if (f) *f = fetched;
			return (fetched == celt) ? S_OK : S_FALSE;
		}
		STDMETHODIMP Skip(ULONG c) override { m_pos += c; return S_OK; }
		STDMETHODIMP Reset() override { m_pos = 0; return S_OK; }
		STDMETHODIMP Clone(IEnumUnknown**) override { return E_NOTIMPL; }
	};
	// ---- Private members ----
	LONG m_ref = 0; // COM reference count
	HWND m_hwnd = NULL; // The host window we draw into
	IUnknown* m_pUnkSite = NULL; // IObjectWithSite back-pointer
	IViewObject* m_pView = NULL; // IE's drawing interface
	IOleInPlaceObject* m_pInPlace = NULL; // IE's in-place positioning interface
	IOleObject* m_pOle = NULL; // THE main IE COM object
	IWebBrowser2* m_pBrowser = NULL; // Our handle to the browser (Navigate, get_Document, etc.)
	DWORD m_cookie = NULL; // Connection point cookie (for unsubscribing events)
};

// TridentWindowData_ - the internal data behind an hTrident handle.
// This is a node in a doubly-linked list (g_trident_head). Each node owns
// an OleSite (the COM hosting layer) and optionally a user WndProc and user data.
struct TridentWindowData_ {
	HWND hwnd = NULL; // The Win32 window
	OleSite* site = NULL; // The COM hosting layer (Release'd, not deleted)
	TridentWndProc userProc = NULL; // User's custom WndProc (optional)
	void* userData = NULL; // User's custom data (optional)
	bool alive = true; // False after WM_NCDESTROY
	TridentWindowData_* prev = nullptr; // Previous node in the linked list
	TridentWindowData_* next = nullptr; // Next node in the linked list
};

// TridentHostWndProc - the Win32 window procedure for the host window.
// 
// Message flow:
//  	1. User's WndProc gets first crack (if set via SetWndProc). If it sets handled=true, we stop.
//  	2. WM_SIZE -> resize the IE control to fill our client area
//  	3. WM_CLOSE -> DestroyWindow, which triggers WM_NCDESTROY
//  	4. WM_NCDESTROY -> THE cleanup handler. This is the LAST message a window ever receives.
// 
//  	We clear GWLP_USERDATA (so NavigateToHTML can detect "window is gone"),
//  	destroy the OLE site, remove from the linked list, and delete h.
//  	After this point, the hTrident handle is INVALID - using it is undefined behavior.
inline LRESULT CALLBACK TridentHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
	hTrident trdt = (hTrident)GetWindowLongPtr(hwnd, GWLP_USERDATA);
	if (msg == WM_NCCREATE) {
		trdt = (hTrident)((CREATESTRUCT*)lp)->lpCreateParams;
		SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)trdt);
		trdt->hwnd = hwnd;
	}
	if (trdt) {
		if (msg == WM_SIZE && trdt->site) {
			trdt->site->Resize(LOWORD(lp), HIWORD(lp));
		}
		if (msg == WM_CLOSE) {
			DestroyWindow(hwnd);
			return 0;
		}
		if (msg == WM_NCDESTROY) {
			SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
			trdt->hwnd = NULL;
			trdt->alive = false;
			if (trdt->site) { trdt->site->Destroy(); trdt->site->Release(); trdt->site = NULL; }
			// Remove from doubly-linked list (O(1))
			if (trdt->prev) trdt->prev->next = trdt->next;
			else g_trident_head = trdt->next;
			if (trdt->next) trdt->next->prev = trdt->prev;
			delete trdt;
			return 0;
		}
		if (trdt->userProc) {
			bool implemented = false;
			LRESULT r = trdt->userProc(trdt, hwnd, msg, wp, lp, &implemented);
			if (implemented) return r;
		}
	}

	return DefWindowProc(hwnd, msg, wp, lp);
}

// Public API implementations
inline void TridentInit() {
	OleInitialize(NULL);
	WNDCLASSEXW wc = { sizeof(wc) };
	wc.lpfnWndProc = TridentHostWndProc;
	wc.hInstance = GetModuleHandle(NULL);
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	wc.lpszClassName = g_trident_wndclass;
	RegisterClassExW(&wc);
}

inline void TridentShutdown() {
	while (g_trident_head) CloseTridentWindow(g_trident_head);
	UnregisterClassW(g_trident_wndclass, GetModuleHandle(NULL));
	OleUninitialize();
}

// Returns true if hwnd is one of our libTridentUI host windows.
// We verify both the window class name and that GWLP_USERDATA points to a live hTrident.
inline bool IsTridentHostWindow(HWND hwnd) {
	if (!hwnd) return false;
	wchar_t cls[LONG_MAX_PATH] = {};
	if (!GetClassNameW(hwnd, cls, LONG_MAX_PATH)) return false;
	if (lstrcmpW(cls, g_trident_wndclass) != 0) return false;
	hTrident h = (hTrident)GetWindowLongPtr(hwnd, GWLP_USERDATA);
	return h && h->alive && h->hwnd == hwnd;
}

// Walks upward from an arbitrary HWND and returns the nearest libTridentUI host.
// This is the correct rule for nested Trident windows: route to the closest host,
// not to the root window and not to every ancestor host that happens to contain it.
inline hTrident FindNearestTridentHostFromHWND(HWND hwnd) {
	for (HWND cur = hwnd; cur; cur = GetParent(cur)) {
		if (IsTridentHostWindow(cur)) {
			return (hTrident)GetWindowLongPtr(cur, GWLP_USERDATA);
		}
	}
	return NULL;
}

// TridentRunMessageLoop - standard Win32 message loop with IE keyboard handling.
// IE needs to process Tab, Enter, Ctrl+C, etc. through its own TranslateAccelerator.
// Without this, keyboard input in the browser would be broken (you couldn't tab
// between form fields, copy text, etc.). We route each message to the NEAREST
// libTridentUI host that owns the target HWND. This keeps nested Trident windows
// from stealing accelerators from one another.
inline void TridentRunMessageLoop() {
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0) > 0) {
		bool MessageEaten = false;
		TridentProcessMessage(msg, &MessageEaten);
		if (!MessageEaten) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		// Quit the program if there are no more Trident windows!
		if (!g_trident_head) break;
	}
	PostQuitMessage(0);
}

// TridentProcessMessage - processes a single message for libTridentUI windows.
// 
// Call this from YOUR OWN message loop when your application has a mix of
// libTridentUI windows and other windows.
// 
// Usage:
//  	MSG msg;
//  	while (GetMessage(&msg, NULL, 0, 0) > 0) {
//  		bool eaten = false;
//  		TridentProcessMessage(msg, &eaten);
//  		if (!eaten) {
//  			TranslateMessage(&msg);
//  			DispatchMessage(&msg);
//  		}
//  	}
// 
// If the message belongs to an IE control inside one of our windows (Tab key,
// Ctrl+C, Enter, etc.), it gets handled internally and *is_eaten is set to true.
// In that case, do NOT call TranslateMessage/DispatchMessage -- IE already ate it.
// If *is_eaten is false, the message is not ours -- pass it along as normal.
// 
// If you don't have your own message loop and all your windows are libTridentUI,
// just call TridentRunMessageLoop() instead -- it calls this function internally.
inline void TridentProcessMessage(MSG& msg, bool* is_eaten) {
	if (!is_eaten) return;
	*is_eaten = false; // Initialize *eaten to always be false at the beginning.
	if (!msg.hwnd) return;
	hTrident host = FindNearestTridentHostFromHWND(msg.hwnd);
	if (!host || !host->alive || !host->site || !host->hwnd) return;
	IWebBrowser2* BrowserPtr = host->site->GetBrowser();
	if (!BrowserPtr) return;
	BrowserPtr->AddRef();
	IOleInPlaceActiveObject* ObjectPtr = NULL;
	if (SUCCEEDED(BrowserPtr->QueryInterface(IID_IOleInPlaceActiveObject, (void**)&ObjectPtr))) {
		*is_eaten = (ObjectPtr->TranslateAccelerator(&msg) == S_OK);
		ObjectPtr->Release();
	}
	BrowserPtr->Release();
}

inline hTrident NewTridentWindow(const wchar_t* title, int x, int y, int w, int h, HWND owner, DWORD style, DWORD styleEX) {
	TridentWindowData_* d = new TridentWindowData_();
	CreateWindowExW(styleEX, g_trident_wndclass, title, style, x, y, w, h, owner, NULL, GetModuleHandle(NULL), d);
	if (!d->hwnd) { delete d; return NULL; }
	d->site = new OleSite();
	if (!d->site->Create(d->hwnd)) {
		d->site->Release(); d->site = NULL;
		SetWindowLongPtr(d->hwnd, GWLP_USERDATA, 0);
		DestroyWindow(d->hwnd);
		delete d;
		return NULL;
	}
	// Insert at head of doubly-linked list
	d->prev = nullptr;
	d->next = g_trident_head;
	if (g_trident_head) g_trident_head->prev = d;
	g_trident_head = d;
	return d;
}

inline void CloseTridentWindow(hTrident h) {
	if (!h || !h->alive) return;
	if (h->hwnd) {
		DestroyWindow(h->hwnd);
	}
	else {
		// This shouldn't happen normally but we still handle it here anyways.
		h->alive = false;
		if (h->site) { h->site->Destroy(); h->site->Release(); h->site = NULL; }
		if (h->prev) h->prev->next = h->next;
		else g_trident_head = h->next;
		if (h->next) h->next->prev = h->prev;
		delete h;
	}
}

inline HWND GetTridentHWND(hTrident h) { return h ? h->hwnd : NULL; }

inline IWebBrowser2* GetTridentBrowser(hTrident h) {
	return (h && h->site) ? h->site->GetBrowser() : NULL;
}

inline void BindFunction(hTrident h, const wchar_t* name, MethodCallback cb) {
	if (h && h->site) h->site->ext.Bind(name, cb);
}

inline void SetWndProc(hTrident h, TridentWndProc proc) { if (h) h->userProc = proc; }
inline void SetUserData(hTrident h, void* data) { if (h) h->userData = data; }
inline void* GetUserData(hTrident h) { return h ? h->userData : NULL; }

inline void NavigateTo(hTrident h, const wchar_t* url) {
	IWebBrowser2* b = GetTridentBrowser(h);
	if (!b) return;
	VARIANT v;
	v.vt = VT_BSTR;
	v.bstrVal = SysAllocString(url);
	b->Navigate2(&v, NULL, NULL, NULL, NULL);
	SysFreeString(v.bstrVal);
}

inline void NavigateToRes(hTrident h, const wchar_t* resName) {
	wchar_t exe[LONG_MAX_PATH] = {};
	GetModuleFileNameW(NULL, exe, LONG_MAX_PATH);
	std::wstring url = L"res://";
	url += exe;
	url += L"/";
	url += resName;
	NavigateTo(h, url.c_str());
}

// NavigateToHTML - loads an HTML string directly into the browser.
// 
// This is trickier than you'd think. IE doesn't have a simple "load this string" API.
// Instead, we have to:
//  	1. Navigate to about:blank (creates an empty document)
//  	2. Wait for about:blank to fully load (pump messages until READYSTATE_COMPLETE)
//  	3. Get the IHTMLDocument2 from the loaded document
//  	4. Call document.write() to inject our HTML
//  	5. Call document.close() to finalize
// 
// The message pump in step 2 is dangerous: while we're pumping, ANY Windows message
// can fire - including WM_CLOSE, which would destroy our window and delete h.
// That's why we:
//  	- AddRef the browser to keep it alive during the pump
//  	- Save hwnd to a local variable (safe even after h is deleted)
//  	- Check IsWindow(hwnd) in the loop (if window died, bail out)
//  	- Re-fetch h from GWLP_USERDATA after the loop (if it's 0, h was deleted)
inline void NavigateToHTML(hTrident h, const wchar_t* html) {
	IWebBrowser2* b = GetTridentBrowser(h);
	if (!b) return;
	b->AddRef(); // prevent browser from being freed during pump
	HWND hwnd = GetTridentHWND(h); // stack copy - safe even if h is deleted

	NavigateTo(h, L"about:blank");

	// Pump messages until about:blank is fully loaded.
	// We can't touch h during this loop - it might get deleted by WM_NCDESTROY!
	READYSTATE st;
	while (SUCCEEDED(b->get_ReadyState(&st)) && st != READYSTATE_COMPLETE) {
		if (!IsWindow(hwnd)) { b->Release(); return; } // window destroyed during pump
		MSG msg;
		while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		_mm_pause();
	}

	// Re-acquire h from the window. If WM_NCDESTROY fired, GWLP_USERDATA is 0.
	h = (hTrident)GetWindowLongPtr(hwnd, GWLP_USERDATA);
	if (!h || !IsWindow(hwnd)) { b->Release(); return; }
	if (h->site) h->site->SetupUIHandler();

	// Now write the HTML into the blank document using document.write().
	hDocument d = GetDocument(h);
	if (!d) { b->Release(); return; }
	// document.write() expects a SAFEARRAY of VARIANTs containing BSTRs.
	SAFEARRAY* sa = SafeArrayCreateVector(VT_VARIANT, 0, 1);
	VARIANT* pv;
	SafeArrayAccessData(sa, (void**)&pv);
	pv->vt = VT_BSTR;
	pv->bstrVal = SysAllocString(html);
	SafeArrayUnaccessData(sa);
	d->write(sa);
	d->close(); // Finalize the document - like closing a file after writing
	SafeArrayDestroy(sa);
	d->Release();
	b->Release();
}

// GetDocument - fetches the current DOM document as IHTMLDocument2.
//
// This gives callers direct access to the Trident DOM API for advanced operations
// like createElement(), get_body(), querying additional interfaces (IHTMLDocument3,
// IHTMLDOMNode, etc.), and custom DOM traversal/manipulation.
//
// Ownership: the returned interface has an outstanding COM reference from QueryInterface,
// so the caller must release it with doc->Release().
inline hDocument GetDocument(hTrident h) {
	IWebBrowser2* b = GetTridentBrowser(h);
	if (!b) return NULL;
	IDispatch* p = NULL;
	if (FAILED(b->get_Document(&p))) return NULL;
	IHTMLDocument2* d = NULL;
	if (FAILED(p->QueryInterface(IID_IHTMLDocument2, (void**)&d))) { p->Release(); return NULL; }
	p->Release();
	return d;
}

inline hDocument2 GetDocument2(hDocument d) {
	IHTMLDocument3* d2 = NULL;
	if (FAILED(d->QueryInterface(IID_IHTMLDocument3, (void**)&d2))) return NULL;
	return d2;
}

// Methods to call JavaScript functions
inline VARIANT EvaluateExpression(hTrident h, const wchar_t* expr) {
	VARIANT result;
	VariantInit(&result);
	if (!expr) return result;
	hDocument doc = GetDocument(h);
	if (!doc) return result;
	IDispatch* scriptDisp = NULL;
	doc->get_Script(&scriptDisp);
	doc->Release();
	if (!scriptDisp) return result;
	DISPID evalId;
	LPOLESTR evalName = SysAllocString(L"eval");
	if (SUCCEEDED(scriptDisp->GetIDsOfNames(IID_NULL, &evalName, 1, LOCALE_USER_DEFAULT, &evalId))) {
		VARIANT arg;
		arg.vt = VT_BSTR;
		arg.bstrVal = SysAllocString(expr);
		DISPPARAMS dp = { &arg, NULL, 1, 0 };
		if (FAILED(scriptDisp->Invoke(evalId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL))) {
			SysFreeString(arg.bstrVal);
			SysFreeString(evalName);
			scriptDisp->Release();
			VariantClear(&result);
			return result;
		}
		SysFreeString(arg.bstrVal);
	}
	SysFreeString(evalName);
	scriptDisp->Release();
	return result;
}

inline VARIANT InvokeGlobalFunction(hTrident h, const wchar_t* jsfunc, VARIANT* args, unsigned int argc) {
	VARIANT result;
	VariantInit(&result);
	if (argc && !args) return result;
	wchar_t** path = new wchar_t*[1];
	path[0] = const_cast<wchar_t*>(jsfunc);
	VARIANT ret = InvokeFunctionByPath(h, path, 1, args, argc);
	delete[] path;
	return ret;
}

// Invoke a function at an arbitrary namespace depth.
// Path is an array of property names to traverse from window.
// 
// Examples:
//  	{ L"alert" }					-> window.alert
//  	{ L"console", L"log" }			-> window.console.log
//  	{ L"external", L"MyMethod" }	-> window.external.MyMethod
//  	{ L"myNS", L"util", L"calc" }	-> window.myNS.util.calc
inline VARIANT InvokeFunctionByPath(hTrident h, wchar_t** path, unsigned int pathlen, VARIANT* args, unsigned int argc) {
	VARIANT result;
	VariantInit(&result);
	if (argc && !args) return result;
	if (!path || !pathlen) return result;
	hDocument doc = GetDocument(h);
	if (!doc) return result;
	IDispatch* scriptDisp = NULL;
	doc->get_Script(&scriptDisp);
	doc->Release();
	if (!scriptDisp) return result;
	for (unsigned int i = 0; i < pathlen; i++) {
		const wchar_t* segment = path[i];
		if (!segment) {
			scriptDisp->Release();
			return result;
		}
		DISPID funcid;
		LPOLESTR funcname = SysAllocString(segment);
		if (FAILED(scriptDisp->GetIDsOfNames(IID_NULL, &funcname, 1, LOCALE_USER_DEFAULT, &funcid))) {
			SysFreeString(funcname);
			scriptDisp->Release();
			return result;
		}
		SysFreeString(funcname);
		if (i < pathlen-1){
			DISPPARAMS dp = { NULL, NULL, 0, 0 };
			if (FAILED(scriptDisp->Invoke(funcid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL)) || result.vt != VT_DISPATCH || !result.pdispVal) {
				scriptDisp->Release();
				VariantClear(&result);
				return result;
			}
			scriptDisp->Release();
			scriptDisp = result.pdispVal;
		}else{
			DISPPARAMS dp = { args, NULL, argc, 0 };
			if (FAILED(scriptDisp->Invoke(funcid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL))){
				scriptDisp->Release();
				VariantClear(&result);
				return result;
			}
			scriptDisp->Release();
			return result;
		}
	}
	return result;
}

inline void ExecuteScript(hTrident h, const wchar_t* jscript) {
	if (!jscript) return;
	hDocument d = GetDocument(h);
	if (!d) return;
	IHTMLWindow2* w = NULL;
	HRESULT hr = d->get_parentWindow(&w);
	if (FAILED(hr)) {d->Release();return;}
	BSTR code = SysAllocString(jscript);
	BSTR lang = SysAllocString(L"JavaScript");
	hr = w->execScript(code, lang, nullptr);
	SysFreeString(code);
	SysFreeString(lang);
	w->Release();
	d->Release();
}

inline HRESULT GetIDsOfNamesEx(IDispatch* disp, BSTR name, DISPID* dispid, bool create) {
	if (!disp || !dispid || !name) return E_POINTER;
	IDispatchEx* dispEx = NULL;
	HRESULT hr = disp->QueryInterface(IID_IDispatchEx, (void**)&dispEx);
	if (FAILED(hr)) return hr;
	hr = dispEx->GetDispID(name, create ? fdexNameEnsure : NULL, dispid);
	dispEx->Release();
	return hr;
}

inline VARIANT DispatchGetNamedProperty(IDispatch* disp, const wchar_t* name) {
	VARIANT result;
	VariantInit(&result);
	if (!disp || !name) return result;
	BSTR bstrName = SysAllocString(name);
	if (!bstrName) return result;
	DISPID dispid;
	HRESULT hr = GetIDsOfNamesEx(disp, bstrName, &dispid);
	SysFreeString(bstrName);
	if (FAILED(hr)) return result;
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	if (FAILED(disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL))) VariantClear(&result);
	return result;
}

inline void DispatchSetNamedProperty(IDispatch* disp, const wchar_t* name, const VARIANT* value) {
	if (!disp || !name || !value) return;
	BSTR bstrName = SysAllocString(name);
	if (!bstrName) return;
	DISPID dispid;
	HRESULT hr = GetIDsOfNamesEx(disp, bstrName, &dispid, true);
	SysFreeString(bstrName);
	if (FAILED(hr)) return;
	DISPID named = DISPID_PROPERTYPUT;
	DISPPARAMS dp = { const_cast<VARIANT*>(value), &named, 1, 1 };
	disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, NULL, NULL, NULL);
}

// Create new JavaScript objects
inline VARIANT CreateObject(hTrident h, const wchar_t* classname, VARIANT* args, unsigned int argc) {
	VARIANT result;
	VariantInit(&result);
	if (argc && !args) return result;
	if (!classname) return result;
	hDocument doc = GetDocument(h);
	if (!doc) return result;
	IDispatch* scriptDisp = NULL;
	doc->get_Script(&scriptDisp);
	doc->Release();
	if (!scriptDisp) return result;
	DISPID dispid;
	LPOLESTR name = SysAllocString(classname);
	if (SUCCEEDED(scriptDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid))) {
		DISPPARAMS dp = { args, NULL, argc, 0 };
		if (FAILED(scriptDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_CONSTRUCT, &dp, &result, NULL, NULL))) {
			SysFreeString(name);
			scriptDisp->Release();
			VariantClear(&result);
			return result;
		}
	}
	SysFreeString(name);
	scriptDisp->Release();
	return result;
}

inline VARIANT CreateEmptyArray(hTrident h) {
	return CreateObject(h, L"Array");
}

inline VARIANT CreateEmptyObject(hTrident h) {
	return CreateObject(h, L"Object");
}

inline VARIANT GetDispatchIndex(VARIANT* arr, unsigned int index) {
	VARIANT result;
	VariantInit(&result);
	if (!arr || arr->vt != VT_DISPATCH || !arr->pdispVal) return result;
	std::wstring indexStr = std::to_wstring(index);
	return DispatchGetNamedProperty(arr->pdispVal, indexStr.c_str());
}

inline void SetDispatchIndex(VARIANT* arr, unsigned int index, const VARIANT* value) {
	if (!arr || arr->vt != VT_DISPATCH || !arr->pdispVal || !value) return;
	std::wstring indexStr = std::to_wstring(index);
	DispatchSetNamedProperty(arr->pdispVal, indexStr.c_str(), value);
}

inline VARIANT GetDispatchSize(VARIANT* arr) {
	VARIANT result;
	VariantInit(&result);
	if (!arr || arr->vt != VT_DISPATCH || !arr->pdispVal) return result;
	return DispatchGetNamedProperty(arr->pdispVal, L"length");
}

inline void DispatchResize(VARIANT* arr, unsigned int newSize) {
	if (!arr || arr->vt != VT_DISPATCH || !arr->pdispVal) return;
	VARIANT sizeVar;
	sizeVar.vt = VT_UI4;
	sizeVar.ulVal = newSize;
	DispatchSetNamedProperty(arr->pdispVal, L"length", &sizeVar);
}

inline VARIANT GetDispatchProperty(VARIANT* obj, const wchar_t* prop) {
	VARIANT result;
	VariantInit(&result);
	if (!obj || obj->vt != VT_DISPATCH || !obj->pdispVal || !prop) return result;
	return DispatchGetNamedProperty(obj->pdispVal, prop);
}

inline void SetDispatchProperty(VARIANT* obj, const wchar_t* prop, const VARIANT* value) {
	if (!obj || obj->vt != VT_DISPATCH || !obj->pdispVal || !prop || !value) return;
	DispatchSetNamedProperty(obj->pdispVal, prop, value);
}

inline VARIANT CallDispatchFunction(VARIANT* func, VARIANT* args, unsigned int argc) {
	VARIANT result;
	VariantInit(&result);
	if (argc && !args) return result;
	if (!func || func->vt != VT_DISPATCH || !func->pdispVal) return result;
	DISPPARAMS dp = { args, NULL, argc, 0 };
	func->pdispVal->Invoke(DISPID_VALUE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL);
	return result;
}

inline VARIANT GetGlobalFunction(hTrident h, const wchar_t* jsfunc) {
	wchar_t** path = new wchar_t*[1];
	path[0] = const_cast<wchar_t*>(jsfunc);
	VARIANT ret = GetFunctionByPath(h, path, 1);
	delete[] path;
	return ret;
}

inline VARIANT GetExternalFunction(hTrident h, const wchar_t* funcName) {
	wchar_t** path = new wchar_t*[2];
	path[0] = const_cast<wchar_t*>(L"external");
	path[1] = const_cast<wchar_t*>(funcName);
	VARIANT ret = GetFunctionByPath(h, path, 2);
	delete[] path;
	return ret;
}

// Obtain a pointer to a function at an arbitrary namespace depth.
// Path is an array of property names to traverse from window.
// 
// Example: window.console.log()
//  	wchar_t** path = new wchar_t*[2];
//  	path[0] = const_cast<wchar_t*>(L"console");
//  	path[1] = const_cast<wchar_t*>(L"log");
//  	GetFunctionByPath(MyTridentWindow, path, 2);
//  	delete[] path;
inline VARIANT GetFunctionByPath(hTrident h, wchar_t** path, unsigned int pathlen) {
	VARIANT result;
	VariantInit(&result);
	if (!path || !pathlen) return result;
	hDocument doc = GetDocument(h);
	if (!doc) return result;
	IDispatch* scriptDisp = NULL;
	doc->get_Script(&scriptDisp);
	doc->Release();
	if (!scriptDisp) return result;
	for (unsigned int i = 0; i < pathlen; i++) {
		const wchar_t* segment = path[i];
		if (!segment) {
			scriptDisp->Release();
			return result;
		}
		DISPID funcid;
		LPOLESTR funcname = SysAllocString(segment);
		if (FAILED(scriptDisp->GetIDsOfNames(IID_NULL, &funcname, 1, LOCALE_USER_DEFAULT, &funcid))) {
			SysFreeString(funcname);
			scriptDisp->Release();
			return result;
		}
		SysFreeString(funcname);
		DISPPARAMS dp = { NULL, NULL, 0, 0 };
		if (FAILED(scriptDisp->Invoke(funcid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL)) || result.vt != VT_DISPATCH || !result.pdispVal) {
			scriptDisp->Release();
			VariantClear(&result);
			return result;
		}
		scriptDisp->Release();
		scriptDisp = result.pdispVal;
	}
	return result;
}

// Obtain a pointer to an anonymous function.
// Parameters are defined as wide strings in reverse order.
// 
// Example:
//  	wchar_t** params = new wchar_t*[3];
//  	params[2] = const_cast<wchar_t*>(L"a");
//  	params[1] = const_cast<wchar_t*>(L"b");
//  	params[0] = const_cast<wchar_t*>(L"c");
//  	CreateFunction(MyTridentWindow, params, 3, L"return a + b + c");
//  	delete[] params;
//  	
// This is equivalent to:
//  	(a, b, c) => {return a + b + c}
inline VARIANT CreateFunction(hTrident h, wchar_t** args, unsigned int argc, const wchar_t* impl) {
	VARIANT result;
	VariantInit(&result);
	if (argc && !args) return result;
	if (!impl) return result;
	argc++; // Increase by one because we need to include the function body.
	VARIANT* newparams = new VARIANT[argc];
	// Function body (last parameter)
	newparams[0].vt = VT_BSTR;
	newparams[0].bstrVal = SysAllocString(impl);
	// Fill the parameters...
	for (unsigned int i = 1; i < argc; i++) {
		newparams[i].vt = VT_BSTR;
		newparams[i].bstrVal = SysAllocString(args[i - 1]);
	}
	result = CreateObject(h, L"Function", newparams, argc);
	for (unsigned int i = 0; i < argc; i++) VariantClear(&newparams[i]);
	delete[] newparams;
	return result;
}

inline hElement GetElement(hTrident h, const wchar_t* id) {
	hDocument d = GetDocument(h);
	if (!d) return NULL;
	hDocument2 d2 = GetDocument2(d);
	if (!d2) { d->Release(); return NULL; }
	BSTR bstrId = SysAllocString(id);
	IHTMLElement* e = NULL;
	d2->getElementById(bstrId, &e);
	SysFreeString(bstrId);
	d2->Release();
	d->Release();
	return e;
}

inline void SetHTML(hElement e, const wchar_t* html) {
	if (!e) return;
	BSTR bstr = SysAllocString(html);
	e->put_innerHTML(bstr);
	SysFreeString(bstr);
}

inline std::wstring GetHTML(hElement e) {
	if (!e) return L"";
	BSTR bstr = NULL;
	e->get_innerHTML(&bstr);
	std::wstring result = bstr ? bstr : L"";
	SysFreeString(bstr);
	return result;
}

inline void SetText(hElement e, const wchar_t* text) {
	if (!e) return;
	BSTR bstr = SysAllocString(text);
	e->put_innerText(bstr);
	SysFreeString(bstr);
}

inline std::wstring GetText(hElement e) {
	if (!e) return L"";
	BSTR bstr = NULL;
	e->get_innerText(&bstr);
	std::wstring result = bstr ? bstr : L"";
	SysFreeString(bstr);
	return result;
}

inline void SetOuterHTML(hElement e, const wchar_t* html) {
	if (!e) return;
	BSTR bstr = SysAllocString(html);
	e->put_outerHTML(bstr);
	SysFreeString(bstr);
}

inline void SetClass(hElement e, const wchar_t* cls) {
	if (!e) return;
	BSTR bstr = SysAllocString(cls);
	e->put_className(bstr);
	SysFreeString(bstr);
}

inline void SetStyle(hElement e, const wchar_t* css) {
	if (!e) return;
	IHTMLStyle* style = NULL;
	if (SUCCEEDED(e->get_style(&style)) && style) {
		BSTR bstr = SysAllocString(css);
		style->put_cssText(bstr);
		SysFreeString(bstr);
		style->Release();
	}
}

inline void SetAttr(hElement e, const wchar_t* name, const wchar_t* val) {
	if (!e) return;
	BSTR bstrName = SysAllocString(name);
	VARIANT v;
	v.vt = VT_BSTR;
	v.bstrVal = SysAllocString(val);
	e->setAttribute(bstrName, v, 0);
	SysFreeString(v.bstrVal);
	SysFreeString(bstrName);
}

inline std::wstring GetAttr(hElement e, const wchar_t* name) {
	if (!e) return L"";
	BSTR bstrName = SysAllocString(name);
	VARIANT v;
	VariantInit(&v);
	e->getAttribute(bstrName, 0, &v);
	SysFreeString(bstrName);
	std::wstring result;
	if (v.vt == VT_BSTR && v.bstrVal)
		result = v.bstrVal;
	VariantClear(&v);
	return result;
}

inline void SetElementHTML(hTrident h, const wchar_t* id, const wchar_t* html) {
	hElement e = GetElement(h, id);
	SetHTML(e, html);
	e->Release();
}

inline std::wstring GetElementHTML(hTrident h, const wchar_t* id) {
	hElement e = GetElement(h, id);
	std::wstring result = GetHTML(e);
	e->Release();
	return result;
}

inline void SetElementText(hTrident h, const wchar_t* id, const wchar_t* text) {
	hElement e = GetElement(h, id);
	SetText(e, text);
	e->Release();
}

inline std::wstring GetElementText(hTrident h, const wchar_t* id) {
	hElement e = GetElement(h, id);
	std::wstring result = GetText(e);
	e->Release();
	return result;
}

// InsertHTML - insert HTML content relative to the <body> element.
// 'where' uses IE's insertAdjacentHTML positions:
//  	"afterbegin"	- first child of <body>
//  	"beforeend"		- last child of <body> (most common - appends to page)
//  	"beforebegin"	- before <body> (rarely useful)
//  	"afterend"		- after <body> (rarely useful)
inline void InsertHTML(hTrident h, const wchar_t* id, InsertLocation where, const wchar_t* html) {
	hElement e = GetElement(h, id);
	if (!e) return;
	static const wchar_t* positions[] = {L"beforeBegin", L"afterBegin", L"beforeEnd", L"afterEnd"};
	BSTR bstrWhere = SysAllocString(positions[where]);
	BSTR bstrHTML = SysAllocString(html);
	e->insertAdjacentHTML(bstrWhere, bstrHTML);
	SysFreeString(bstrWhere);
	SysFreeString(bstrHTML);
	e->Release();
}

inline void InsertHTMLAtBody(hTrident h, InsertLocation where, const wchar_t* html) {
	hDocument doc = GetDocument(h);
	if (!doc) return;
	IHTMLElement* body = NULL;
	doc->get_body(&body);
	doc->Release();
	if (!body) return;
	static const wchar_t* positions[] = {L"beforeBegin", L"afterBegin", L"beforeEnd", L"afterEnd"};
	BSTR bstrWhere = SysAllocString(positions[where]);
	BSTR bstrHTML = SysAllocString(html);
	body->insertAdjacentHTML(bstrWhere, bstrHTML);
	SysFreeString(bstrWhere);
	SysFreeString(bstrHTML);
	body->Release();
}

// RemoveElement - remove an element from the DOM by its ID.
inline void RemoveElement(hTrident h, const wchar_t* id) {
	hElement e = GetElement(h, id);
	if (!e) return;
	BSTR empty = SysAllocString(L"");
	e->put_outerHTML(empty);
	SysFreeString(empty);
	e->Release();
}

// ============================================================================
// ActiveX Control Creation
// ============================================================================
// 
// This section lets you create CUSTOM ActiveX controls that can be embedded
// in HTML pages via <object> tags. It's the reverse of the OleSite section above:
// 
//  	OleSite = "we host IE's ActiveX control inside our window"
//  	ActiveXControl = "IE hosts OUR ActiveX control inside its page"
// 
// How it works:
//  	1. You call RegisterControl("MyWidget", myWndProc) - this registers a COM
//  	   class factory in memory (no registry!) using CoRegisterClassObject.
//  	2. You get back a deterministic CLSID generated from the name hash.
//  	3. In your HTML: <object classid="clsid:{that-guid}" width="300" height="200">
//  	4. When IE encounters that <object> tag, it calls CoCreateInstance with the CLSID,
//  	   which hits our class factory, which creates an ActiveXControl object.
//  	5. IE calls DoVerb(INPLACEACTIVATE) on it, and our control creates a child window
//  	   inside IE's document area. Your WndProc handles WM_PAINT, mouse clicks, etc.
//  	6. JavaScript can call methods/properties on the control via IDispatch.
// 
// The control implements these interfaces (IE requires all of them):
//  	IOleObject					- Basic embedding protocol
//  	IOleInPlaceObject			- "I have a window inside your window"
//  	IOleInPlaceActiveObject		- Keyboard accelerator forwarding
//  	IOleControl					- Control-specific protocol
//  	IViewObject2				- "Here's how to draw me when I'm not active"
//  	IPersistStreamInit			- "I have no persistent state" (stub)
//  	IObjectSafety				- "Trust me, I'm safe" (self-declaration)
//  	IDispatch					- Methods and properties callable from JavaScript
//  	IConnectionPointContainer	- Outbound events (C++ -> JS)
//  	IDropTarget/IDropSource		- OLE drag & drop
//  	IProvideClassInfo2			- Tells IE which outbound event interface to use

// CLSIDFromName - generates a deterministic GUID from a control name using FNV-1a hash.
// Same name always produces the same CLSID, so you can hardcode it in HTML.
inline CLSID CLSIDFromName(const wchar_t* name) {
	CLSID result = {};
	if (!name) return result;

	HCRYPTPROV hProv = NULL;
	HCRYPTHASH hHash = NULL;

	CryptAcquireContext(&hProv, NULL, NULL, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT);
	CryptCreateHash(hProv, CALG_MD5, 0, 0, &hHash);
	DWORD dwDataLen = (DWORD)(wcslen(name) * sizeof(wchar_t));
	CryptHashData(hHash, (const BYTE*)name, dwDataLen, 0);
	DWORD dwHashLen = sizeof(CLSID);
	CryptGetHashParam(hHash, HP_HASHVAL, reinterpret_cast<unsigned char*>(&result), &dwHashLen, 0);

	CryptDestroyHash(hHash);
	CryptReleaseContext(hProv, 0);
	return result;
}

struct CLSIDCompare {
	bool operator()(const CLSID& a, const CLSID& b) const {
		return memcmp(&a, &b, sizeof(CLSID)) < 0;
	}
};

struct DispEntry {
	enum Type { METHOD, PROPERTY } type;
	ControlMethodCallback method;
	PropertyGetter getter;
	PropertySetter setter;
};

// EventTypeInfo - ITypeInfo for the outgoing event dispatch interface.
// 
// This is half of the type info system that makes IE's native attachEvent() work.
// It describes our event interface as TKIND_DISPATCH and provides GetIDsOfNames()
// to map event names like "OnClick" to DISPIDs.
// 
// It also implements GetFuncDesc (enumerate events by index) and GetNames
// (reverse-lookup DISPID -> name) for richer type info support.
// 
// Event name -> DISPID registry lives here (moved from ControlReg_).
// RegisterEvent() adds events, GetEventId() looks them up for FireEventByName.
// 
// IE's attachEvent("OnClick", handler) flow:
//  	1. IE gets CoclassTypeInfo via GetClassInfo()
//  	2. IE finds the [source,default] impl type -> navigates to THIS EventTypeInfo
//  	3. IE calls GetIDsOfNames("OnClick") -> gets the DISPID
//  	4. IE calls IConnectionPoint::Advise with a sink that maps that DISPID to the JS handler
//  	5. C++ calls FireEvent(ctrl, dispid) -> iterates sinks -> IE routes to JS handler
class EventTypeInfo : public ITypeInfo {
	LONG m_ref = 0;
	CRITICAL_SECTION m_cs;
	CLSID m_clsid = {};
	DISPID m_nextId = INT32_MAX;
	std::map<std::wstring, DISPID> m_name2id = {};
	std::map<DISPID, std::wstring> m_id2name = {}; // Reverse lookup for GetNames
	std::vector<DISPID> m_orderedIds = {}; // Ordered list for GetFuncDesc by index
	std::map<DISPID, unsigned int> m_id2cParams = {}; // Parameter count per event (for GetFuncDesc)
public:
	EventTypeInfo(const CLSID& clsid) : m_clsid(clsid) {
		InterlockedIncrement(&m_ref);
		InitializeCriticalSection(&m_cs);
	}
	~EventTypeInfo() { DeleteCriticalSection(&m_cs); }

	// Register a new event name with its parameter count. Returns DISPID. Idempotent for same name.
	DISPID RegisterEvent(const wchar_t* name, unsigned int paramCount = 0) {
		EnterCriticalSection(&m_cs);
		auto it = m_name2id.find(name);
		if (it != m_name2id.end()) { DISPID d = it->second; LeaveCriticalSection(&m_cs); return d; }
		if (m_nextId <= 0) { LeaveCriticalSection(&m_cs); return DISPID_UNKNOWN; }
		DISPID id = m_nextId--;
		m_name2id[name] = id;
		m_id2name[id] = name;
		m_orderedIds.push_back(id);
		m_id2cParams[id] = paramCount;
		LeaveCriticalSection(&m_cs);
		return id;
	}

	// Look up event DISPID by name. Returns DISPID_UNKNOWN if not found.
	DISPID GetEventId(const wchar_t* name) {
		EnterCriticalSection(&m_cs);
		auto it = m_name2id.find(name);
		DISPID id = (it != m_name2id.end()) ? it->second : DISPID_UNKNOWN;
		LeaveCriticalSection(&m_cs);
		return id;
	}

	// Get the registered parameter count for an event. Returns 0 if not found.
	unsigned int GetParamCount(DISPID id) {
		EnterCriticalSection(&m_cs);
		auto it = m_id2cParams.find(id);
		unsigned int n = (it != m_id2cParams.end()) ? it->second : 0;
		LeaveCriticalSection(&m_cs);
		return n;
	}

	// ---- IUnknown ----
	STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
		if (riid == IID_IUnknown || riid == IID_ITypeInfo) { *ppv = this; AddRef(); return S_OK; }
		*ppv = NULL; return E_NOINTERFACE;
	}
	STDMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&m_ref); }
	STDMETHODIMP_(ULONG) Release() override { LONG r = InterlockedDecrement(&m_ref); if (!r) delete this; return r; }

	// ---- ITypeInfo - the methods IE actually calls for attachEvent ----

	// GetTypeAttr - describe ourselves as a dispatch interface for outgoing events.
	STDMETHODIMP GetTypeAttr(TYPEATTR** pp) override {
		if (!pp) return E_POINTER;
		TYPEATTR* ta = (TYPEATTR*)CoTaskMemAlloc(sizeof(TYPEATTR));
		if (!ta) return E_OUTOFMEMORY;
		memset(ta, 0, sizeof(TYPEATTR));
		ta->guid = IID_IDispatch;
		ta->typekind = TKIND_DISPATCH;
		ta->cbSizeVft = 7 * sizeof(void*);
		ta->wTypeFlags = TYPEFLAG_FDISPATCHABLE;
		EnterCriticalSection(&m_cs);
		ta->cFuncs = (WORD)m_name2id.size();
		LeaveCriticalSection(&m_cs);
		*pp = ta;
		return S_OK;
	}
	STDMETHODIMP_(void) ReleaseTypeAttr(TYPEATTR* p) override { if (p) CoTaskMemFree(p); }

	// GetFuncDesc - describe event N as a dispinterface method with its parameters.
	// IE uses cParams and lprgelemdescParam to know how many args to pass to JS handlers.
	STDMETHODIMP GetFuncDesc(UINT index, FUNCDESC** ppDesc) override {
		if (!ppDesc) return E_POINTER;
		EnterCriticalSection(&m_cs);
		if (index >= m_orderedIds.size()) { LeaveCriticalSection(&m_cs); return TYPE_E_ELEMENTNOTFOUND; }
		DISPID id = m_orderedIds[index];
		unsigned int cParams = m_id2cParams.count(id) ? m_id2cParams[id] : 0;
		LeaveCriticalSection(&m_cs);
		FUNCDESC* fd = (FUNCDESC*)CoTaskMemAlloc(sizeof(FUNCDESC));
		if (!fd) return E_OUTOFMEMORY;
		memset(fd, 0, sizeof(FUNCDESC));
		fd->memid = id;
		fd->funckind = FUNC_DISPATCH;
		fd->invkind = INVOKE_FUNC;
		fd->callconv = CC_STDCALL;
		fd->elemdescFunc.tdesc.vt = VT_VOID;
		fd->cParams = cParams;
		if (cParams > 0) {
			fd->lprgelemdescParam = (ELEMDESC*)CoTaskMemAlloc(cParams * sizeof(ELEMDESC));
			if (!fd->lprgelemdescParam) { CoTaskMemFree(fd); return E_OUTOFMEMORY; }
			memset(fd->lprgelemdescParam, 0, cParams * sizeof(ELEMDESC));
			for (unsigned int i = 0; i < cParams; i++)
				fd->lprgelemdescParam[i].tdesc.vt = VT_VARIANT; // Accept any type
		}
		*ppDesc = fd;
		return S_OK;
	}
	STDMETHODIMP_(void) ReleaseFuncDesc(FUNCDESC* p) override {
		if (p) {
			if (p->lprgelemdescParam) CoTaskMemFree(p->lprgelemdescParam);
			CoTaskMemFree(p);
		}
	}

	// GetNames - reverse-lookup DISPID -> event name.
	STDMETHODIMP GetNames(MEMBERID id, BSTR* names, UINT maxNames, UINT* pcNames) override {
		if (!names || !pcNames) return E_POINTER;
		*pcNames = 0;
		if (maxNames == 0) return S_OK;
		EnterCriticalSection(&m_cs);
		auto it = m_id2name.find(id);
		if (it == m_id2name.end()) { LeaveCriticalSection(&m_cs); return TYPE_E_ELEMENTNOTFOUND; }
		names[0] = SysAllocString(it->second.c_str());
		LeaveCriticalSection(&m_cs);
		*pcNames = 1;
		return S_OK;
	}

	// GetIDsOfNames - THE key method. Maps "OnClick" -> DISPID for attachEvent.
	STDMETHODIMP GetIDsOfNames(LPOLESTR* names, UINT cnt, MEMBERID* ids) override {
		HRESULT hr = S_OK;
		EnterCriticalSection(&m_cs);
		for (UINT i = 0; i < cnt; i++) {
			auto it = m_name2id.find(names[i]);
			if (it != m_name2id.end()) {
				ids[i] = it->second;
			}
			else {
				ids[i] = MEMBERID_NIL;
				hr = DISP_E_UNKNOWNNAME;
			}
		}
		LeaveCriticalSection(&m_cs);
		return hr;
	}

	// ---- ITypeInfo stubs - not needed for attachEvent ----
	STDMETHODIMP GetTypeComp(ITypeComp**) override { return E_NOTIMPL; }
	STDMETHODIMP GetVarDesc(UINT, VARDESC**) override { return E_NOTIMPL; }
	STDMETHODIMP GetRefTypeOfImplType(UINT, HREFTYPE*) override { return E_NOTIMPL; }
	STDMETHODIMP GetRefTypeInfo(HREFTYPE, ITypeInfo**) override { return E_NOTIMPL; }
	STDMETHODIMP GetImplTypeFlags(UINT, INT*) override { return E_NOTIMPL; }
	STDMETHODIMP GetDocumentation(MEMBERID, BSTR*, BSTR*, DWORD*, BSTR*) override { return E_NOTIMPL; }
	STDMETHODIMP GetDllEntry(MEMBERID, INVOKEKIND, BSTR*, BSTR*, WORD*) override { return E_NOTIMPL; }
	STDMETHODIMP Invoke(void*, MEMBERID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) override { return E_NOTIMPL; }
	STDMETHODIMP AddressOfMember(MEMBERID, INVOKEKIND, void**) override { return E_NOTIMPL; }
	STDMETHODIMP CreateInstance(IUnknown*, REFIID, void**) override { return E_NOTIMPL; }
	STDMETHODIMP GetMops(MEMBERID, BSTR*) override { return E_NOTIMPL; }
	STDMETHODIMP GetContainingTypeLib(ITypeLib**, UINT*) override { return E_NOTIMPL; }
	STDMETHODIMP_(void) ReleaseVarDesc(VARDESC*) override {}
};

// CoclassTypeInfo - coclass-level ITypeInfo that navigates IE to EventTypeInfo.
// 
// IE's attachEvent navigates: coclass -> [source,default] impl type -> event ITypeInfo.
// This class represents the coclass level. It has one impl type (our EventTypeInfo)
// flagged as IMPLTYPEFLAG_FSOURCE | IMPLTYPEFLAG_FDEFAULT.
// When IE calls GetRefTypeInfo(), we return the EventTypeInfo.
class CoclassTypeInfo : public ITypeInfo {
	LONG m_ref = 0;
	CLSID m_clsid = {};
	EventTypeInfo* m_eventTI = NULL; // Strong reference (AddRef'd)
public:
	CoclassTypeInfo(const CLSID& clsid, EventTypeInfo* eti) : m_clsid(clsid), m_eventTI(eti) {
		InterlockedIncrement(&m_ref);
		if (m_eventTI) m_eventTI->AddRef();
	}
	~CoclassTypeInfo() { if (m_eventTI) m_eventTI->Release(); }

	STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
		if (riid == IID_IUnknown || riid == IID_ITypeInfo) { *ppv = this; AddRef(); return S_OK; }
		*ppv = NULL; return E_NOINTERFACE;
	}
	STDMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&m_ref); }
	STDMETHODIMP_(ULONG) Release() override { LONG r = InterlockedDecrement(&m_ref); if (!r) delete this; return r; }

	// GetTypeAttr - we're a coclass with 1 impl type (the source event interface)
	STDMETHODIMP GetTypeAttr(TYPEATTR** pp) override {
		if (!pp) return E_POINTER;
		TYPEATTR* ta = (TYPEATTR*)CoTaskMemAlloc(sizeof(TYPEATTR));
		if (!ta) return E_OUTOFMEMORY;
		memset(ta, 0, sizeof(TYPEATTR));
		ta->guid = m_clsid;
		ta->typekind = TKIND_COCLASS;
		ta->cImplTypes = 1;
		*pp = ta;
		return S_OK;
	}
	STDMETHODIMP_(void) ReleaseTypeAttr(TYPEATTR* p) override { if (p) CoTaskMemFree(p); }

	// GetImplTypeFlags - our one impl type is [source, default]
	STDMETHODIMP GetImplTypeFlags(UINT index, INT* flags) override {
		if (index != 0 || !flags) return TYPE_E_ELEMENTNOTFOUND;
		*flags = IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE;
		return S_OK;
	}

	// GetRefTypeOfImplType - return a handle for our source interface
	STDMETHODIMP GetRefTypeOfImplType(UINT index, HREFTYPE* pRef) override {
		if (index != 0 || !pRef) return TYPE_E_ELEMENTNOTFOUND;
		*pRef = 1;
		return S_OK;
	}

	// GetRefTypeInfo - given the handle, return the EventTypeInfo
	STDMETHODIMP GetRefTypeInfo(HREFTYPE ref, ITypeInfo** ppTI) override {
		if (!ppTI) return E_POINTER;
		if (ref != 1 || !m_eventTI) { *ppTI = NULL; return E_INVALIDARG; }
		*ppTI = m_eventTI;
		m_eventTI->AddRef();
		return S_OK;
	}

	// Stubs - not needed at the coclass level
	STDMETHODIMP GetTypeComp(ITypeComp**) override { return E_NOTIMPL; }
	STDMETHODIMP GetFuncDesc(UINT, FUNCDESC**) override { return E_NOTIMPL; }
	STDMETHODIMP GetVarDesc(UINT, VARDESC**) override { return E_NOTIMPL; }
	STDMETHODIMP GetNames(MEMBERID, BSTR*, UINT, UINT*) override { return E_NOTIMPL; }
	STDMETHODIMP GetIDsOfNames(LPOLESTR*, UINT, MEMBERID*) override { return E_NOTIMPL; }
	STDMETHODIMP GetDocumentation(MEMBERID, BSTR*, BSTR*, DWORD*, BSTR*) override { return E_NOTIMPL; }
	STDMETHODIMP GetDllEntry(MEMBERID, INVOKEKIND, BSTR*, BSTR*, WORD*) override { return E_NOTIMPL; }
	STDMETHODIMP Invoke(void*, MEMBERID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) override { return E_NOTIMPL; }
	STDMETHODIMP AddressOfMember(MEMBERID, INVOKEKIND, void**) override { return E_NOTIMPL; }
	STDMETHODIMP CreateInstance(IUnknown*, REFIID, void**) override { return E_NOTIMPL; }
	STDMETHODIMP GetMops(MEMBERID, BSTR*) override { return E_NOTIMPL; }
	STDMETHODIMP GetContainingTypeLib(ITypeLib**, UINT*) override { return E_NOTIMPL; }
	STDMETHODIMP_(void) ReleaseFuncDesc(FUNCDESC*) override {}
	STDMETHODIMP_(void) ReleaseVarDesc(VARDESC*) override {}
};

// ControlReg_ - stores everything about a registered control CLASS (shared by all instances).
// Think of it like a "template": the WndProc, the CLSID, the bound methods/properties.
// Reference-counted because both the registry and active instances hold pointers to it.
// UnregisterControl releases the registry's ref; instances release theirs in ~ActiveXControl.
struct ControlReg_ {
	LONG refCount = 0;
	std::wstring name;
	ControlWndProc wndProc = NULL;
	void* defaultUserData = NULL;
	CLSID clsid = {};
	DWORD comCookie = NULL; // CoRegisterClassObject cookie - needed to unregister
	std::wstring wndClassName; // "TridentControl_" + name
	ATOM wndAtom = 0; // RegisterClassEx return value
	DWORD extraStyle = NULL; // Extra WS_ flags for CreateWindowEx
	CRITICAL_SECTION cs; // Protects the maps below
	DISPID nextDispId = INT32_MAX; // Method/property IDs count down from here
	std::map<std::wstring, DISPID> name2id = {};
	std::map<DISPID, DispEntry> entries = {};
	EventTypeInfo* eventInfo = NULL; // Owns event name->DISPID registry + dispatch ITypeInfo
	CoclassTypeInfo* coclassInfo = NULL; // Coclass-level ITypeInfo for GetClassInfo()
	ControlReg_() {
		InterlockedIncrement(&refCount);
		InitializeCriticalSection(&cs);
	}
	~ControlReg_() {
		if (coclassInfo) coclassInfo->Release(); // Release dependent first
		if (eventInfo) eventInfo->Release();
		if (wndAtom) UnregisterClassW(wndClassName.c_str(), GetModuleHandle(NULL));
		DeleteCriticalSection(&cs);
	}
	void AddRef() { InterlockedIncrement(&refCount); }
	void Release() { if (InterlockedDecrement(&refCount) == 0) delete this; }
};

// Global registry: CLSID -> ControlReg_*. This is how CoCreateInstance finds our factory.
inline std::map<CLSID, ControlReg_*, CLSIDCompare>& CtrlRegistry() {
	static std::map<CLSID, ControlReg_*, CLSIDCompare> s;
	return s;
}

// ControlInstance_ - per-instance data for each <object> tag on a page.
// If you have 3 <object> tags with the same classid, you get 3 ControlInstance_ objects
// but they all share the same ControlReg_ (same WndProc, same methods).
struct ControlInstance_ {
	ControlReg_* reg = NULL; // Back-pointer to the class registration
	HWND hwnd = NULL; // The child window we create inside IE
	void* userData = NULL; // Per-instance user data
	IOleClientSite* clientSite = NULL; // IE's hosting site for this control
	IOleAdviseHolder* adviseHolder = NULL;
	IOleInPlaceSite* inPlaceSite = NULL;
	bool active = false; // True after DoVerb(INPLACEACTIVATE)
	bool uiActive = false; // True after DoVerb(UIACTIVATE)
	RECT rcPos = { 0,0,0,0 }; // Our position inside IE's document
	SIZEL extent = { 0, 0 }; // Size in HIMETRIC units (IE sends this via SetExtent)
	std::map<DWORD, IDispatch*> sinks = {}; // Event subscribers (from IConnectionPoint::Advise)
	DWORD nextCookie = NULL; // Cookie counter for Advise/Unadvise
	DropCallback onDrop = NULL; // Drag-drop callbacks
	DropEnterCallback onDragEnter = NULL;
	DropOverCallback onDragOver = NULL;
	DropLeaveCallback onDragLeave = NULL;
	bool dropRegistered = false;
};

class CtrlConnectionPoint;
inline LRESULT CALLBACK CtrlWndProc(HWND, UINT, WPARAM, LPARAM);

class ActiveXControl :
	public IOleObject, public IOleInPlaceObject, public IOleInPlaceActiveObject,
	public IOleControl, public IViewObject2, public IPersistStreamInit,
	public IObjectSafety, public IDispatch, public IConnectionPointContainer,
	public IDropTarget, public IDropSource, public IProvideClassInfo2
{
public:
	LONG m_ref = 0;
	ControlInstance_* m_inst = NULL;

	ActiveXControl(ControlReg_* reg) {
		InterlockedIncrement(&m_ref);
		m_inst = new ControlInstance_();
		m_inst->reg = reg;
		reg->AddRef();
		m_inst->userData = reg->defaultUserData;
		memset(&m_inst->rcPos, 0, sizeof(RECT));
	}
	~ActiveXControl() {
		for (auto& kv : m_inst->sinks) if (kv.second) kv.second->Release();
		if (m_inst->clientSite) m_inst->clientSite->Release();
		if (m_inst->adviseHolder) m_inst->adviseHolder->Release();
		if (m_inst->inPlaceSite) m_inst->inPlaceSite->Release();
		ControlReg_* reg = m_inst->reg;
		delete m_inst;
		if (reg) reg->Release();
	}

	STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
		*ppv = NULL;
		if (riid == IID_IUnknown) *ppv = static_cast<IOleObject*>(this);
		else if (riid == IID_IOleObject) *ppv = static_cast<IOleObject*>(this);
		else if (riid == IID_IOleWindow) *ppv = static_cast<IOleInPlaceObject*>(this);
		else if (riid == IID_IOleInPlaceObject) *ppv = static_cast<IOleInPlaceObject*>(this);
		else if (riid == IID_IOleInPlaceActiveObject) *ppv = static_cast<IOleInPlaceActiveObject*>(this);
		else if (riid == IID_IOleControl) *ppv = static_cast<IOleControl*>(this);
		else if (riid == IID_IViewObject || riid == IID_IViewObject2) *ppv = static_cast<IViewObject2*>(this);
		else if (riid == IID_IPersist || riid == IID_IPersistStreamInit) *ppv = static_cast<IPersistStreamInit*>(this);
		else if (riid == IID_IObjectSafety) *ppv = static_cast<IObjectSafety*>(this);
		else if (riid == IID_IDispatch)  *ppv = static_cast<IDispatch*>(this);
		else if (riid == IID_IConnectionPointContainer) *ppv = static_cast<IConnectionPointContainer*>(this);
		else if (riid == IID_IDropTarget) *ppv = static_cast<IDropTarget*>(this);
		else if (riid == IID_IDropSource) *ppv = static_cast<IDropSource*>(this);
		else if (riid == IID_IProvideClassInfo || riid == IID_IProvideClassInfo2) *ppv = static_cast<IProvideClassInfo2*>(this);
		if (*ppv) { AddRef(); return S_OK; }
		return E_NOINTERFACE;
	}
	STDMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&m_ref); }
	STDMETHODIMP_(ULONG) Release() override { LONG r = InterlockedDecrement(&m_ref); if (!r) delete this; return r; }

	// IOleObject
	STDMETHODIMP SetClientSite(IOleClientSite* p) override { if (m_inst->clientSite) m_inst->clientSite->Release(); m_inst->clientSite = p; if (p) p->AddRef(); return S_OK; }
	STDMETHODIMP GetClientSite(IOleClientSite** pp) override { if (!pp) return E_POINTER; *pp = m_inst->clientSite; if (*pp) (*pp)->AddRef(); return S_OK; }
	STDMETHODIMP SetHostNames(LPCOLESTR, LPCOLESTR) override { return S_OK; }
	STDMETHODIMP Close(DWORD) override { if (m_inst->active) InPlaceDeactivate(); return S_OK; }
	STDMETHODIMP SetMoniker(DWORD, IMoniker*) override { return E_NOTIMPL; }
	STDMETHODIMP GetMoniker(DWORD, DWORD, IMoniker** p) override { if (!p) return E_POINTER; *p = NULL; return E_NOTIMPL; }
	STDMETHODIMP InitFromData(IDataObject*, BOOL, DWORD) override { return E_NOTIMPL; }
	STDMETHODIMP GetClipboardData(DWORD, IDataObject**) override { return E_NOTIMPL; }

	// DoVerb - IE tells us to activate. This is where our control comes alive.
	//
	// When IE encounters <object classid="..."> in HTML, it creates our ActiveXControl
	// via the class factory, then calls DoVerb(OLEIVERB_INPLACEACTIVATE).
	// Our job: create a child window inside IE's document and start drawing.
	//
	// The activation sequence:
	//  	1. Get IOleInPlaceSite from our container (IE's OLE site for this <object>)
	//  	2. Ask permission to activate: CanInPlaceActivate() + OnInPlaceActivate()
	//  	3. Lock ourselves as "running" so COM doesn't garbage-collect us
	//  	4. Get the container window and the rectangle we should occupy
	//  	5. Handle zero-size rect (IE sometimes gives us position but no size yet -
	//  	   fall back to the HIMETRIC extent that SetExtent() stored earlier)
	//  	6. Register a Win32 window class and create a child window
	//  	7. Our CtrlWndProc routes messages to the user's ControlWndProc callback
	STDMETHODIMP DoVerb(LONG iVerb, LPMSG, IOleClientSite*, LONG, HWND hwndParent, LPCRECT prc) override {
		if (iVerb == OLEIVERB_INPLACEACTIVATE || iVerb == OLEIVERB_UIACTIVATE || iVerb == OLEIVERB_SHOW) {
			if (!m_inst->active) {
				if (!m_inst->clientSite) return E_FAIL;
				if (m_inst->inPlaceSite) {
					m_inst->inPlaceSite->Release();
					m_inst->inPlaceSite = NULL;
				}
				m_inst->clientSite->QueryInterface(IID_IOleInPlaceSite, (void**)&m_inst->inPlaceSite);
				if (!m_inst->inPlaceSite) return E_FAIL;
				m_inst->inPlaceSite->CanInPlaceActivate();
				m_inst->inPlaceSite->OnInPlaceActivate();
				IOleObject* pOle = NULL;
				QueryInterface(IID_IOleObject, (void**)&pOle);
				if (pOle) {
					OleLockRunning(pOle, TRUE, FALSE);
					pOle->Release();
				}
				IOleInPlaceFrame* pFrame = NULL; IOleInPlaceUIWindow* pUIWin = NULL;
				RECT rcPos, rcClip; OLEINPLACEFRAMEINFO fi = { sizeof(fi) };
				m_inst->inPlaceSite->GetWindowContext(&pFrame, &pUIWin, &rcPos, &rcClip, &fi);
				if (prc) rcPos = *prc;
				int w = rcPos.right - rcPos.left;
				int h = rcPos.bottom - rcPos.top;
				if (w <= 0 && m_inst->extent.cx > 0) w = MulDiv(m_inst->extent.cx, 96, 2540);
				if (h <= 0 && m_inst->extent.cy > 0) h = MulDiv(m_inst->extent.cy, 96, 2540);
				rcPos.right = rcPos.left + w;
				rcPos.bottom = rcPos.top + h;
				m_inst->rcPos = rcPos;
				HWND hwndC = NULL; m_inst->inPlaceSite->GetWindow(&hwndC);
				if (!m_inst->reg->wndAtom) {
					WNDCLASSEXW wc = { sizeof(wc) };
					wc.style = CS_HREDRAW | CS_VREDRAW;
					wc.lpfnWndProc = CtrlWndProc;
					wc.hInstance = GetModuleHandle(NULL);
					wc.hCursor = LoadCursor(NULL, IDC_ARROW);
					wc.lpszClassName = m_inst->reg->wndClassName.c_str();
					m_inst->reg->wndAtom = RegisterClassExW(&wc);
				}
				m_inst->hwnd = CreateWindowExW(0, m_inst->reg->wndClassName.c_str(), NULL,
					WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | m_inst->reg->extraStyle,
					rcPos.left, rcPos.top, rcPos.right - rcPos.left, rcPos.bottom - rcPos.top,
					hwndC, NULL, GetModuleHandle(NULL), this);
				if (pFrame) pFrame->Release();
				if (pUIWin) pUIWin->Release();
				if (!m_inst->hwnd) {
					// Window creation failed - roll back activation state
					IOleObject* pOle2 = NULL;
					QueryInterface(IID_IOleObject, (void**)&pOle2);
					if (pOle2) { OleLockRunning(pOle2, FALSE, FALSE); pOle2->Release(); }
					m_inst->inPlaceSite->OnInPlaceDeactivate();
					return E_FAIL;
				}
				m_inst->active = true;
			}
			if (iVerb == OLEIVERB_UIACTIVATE && !m_inst->uiActive) {
				m_inst->uiActive = true;
				if (m_inst->inPlaceSite) m_inst->inPlaceSite->OnUIActivate();
			}
			return S_OK;
		}
		if (iVerb == OLEIVERB_HIDE) {
			if (m_inst->hwnd) ShowWindow(m_inst->hwnd, SW_HIDE);
			return S_OK;
		}
		if (iVerb > 0) return OLEOBJ_S_INVALIDVERB;
		return E_NOTIMPL;
	}

	STDMETHODIMP EnumVerbs(IEnumOLEVERB**) override { return E_NOTIMPL; }
	STDMETHODIMP Update() override { return S_OK; }
	STDMETHODIMP IsUpToDate() override { return S_OK; }
	STDMETHODIMP GetUserClassID(CLSID* p) override { if (!p) return E_POINTER; *p = m_inst->reg->clsid; return S_OK; }
	STDMETHODIMP GetUserType(DWORD, LPOLESTR*) override { return E_NOTIMPL; }
	// SetExtent - IE tells us how big we should be, in HIMETRIC units (1 HIMETRIC = 0.01mm).
	// This gets called BEFORE DoVerb, from the HTML width/height attributes.
	// We store it so DoVerb can use it as fallback when IE gives us a zero-size rect.
	// Conversion: pixels = HIMETRIC * 96 / 2540 (at standard 96 DPI)
	STDMETHODIMP SetExtent(DWORD dwAspect, SIZEL* pSizel) override {
		if (dwAspect != DVASPECT_CONTENT || !pSizel) return E_INVALIDARG;
		m_inst->extent = *pSizel;
		if (m_inst->hwnd) {
			int w = MulDiv(pSizel->cx, 96, 2540);
			int h = MulDiv(pSizel->cy, 96, 2540);
			SetWindowPos(m_inst->hwnd, NULL, 0, 0, w, h, SWP_NOMOVE | SWP_NOZORDER);
		}
		return S_OK;
	}
	STDMETHODIMP GetExtent(DWORD, SIZEL* p) override {
		if (!p) return E_POINTER;
		if (m_inst->extent.cx > 0 && m_inst->extent.cy > 0) {
			*p = m_inst->extent;
		}
		else {
			p->cx = MulDiv(m_inst->rcPos.right - m_inst->rcPos.left, 2540, 96);
			p->cy = MulDiv(m_inst->rcPos.bottom - m_inst->rcPos.top, 2540, 96);
		}
		return S_OK;
	}
	STDMETHODIMP Advise(IAdviseSink* s, DWORD* c) override {
		if (!m_inst->adviseHolder) CreateOleAdviseHolder(&m_inst->adviseHolder);
		return m_inst->adviseHolder->Advise(s, c);
	}
	STDMETHODIMP Unadvise(DWORD c) override { return m_inst->adviseHolder ? m_inst->adviseHolder->Unadvise(c) : E_FAIL; }
	STDMETHODIMP EnumAdvise(IEnumSTATDATA** p) override { return m_inst->adviseHolder ? m_inst->adviseHolder->EnumAdvise(p) : E_FAIL; }
	STDMETHODIMP GetMiscStatus(DWORD, DWORD* p) override {
		if (!p) return E_POINTER;
		*p = OLEMISC_RECOMPOSEONRESIZE | OLEMISC_INSIDEOUT | OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_SETCLIENTSITEFIRST;
		return S_OK;
	}
	STDMETHODIMP SetColorScheme(LOGPALETTE*) override { return E_NOTIMPL; }

	STDMETHODIMP GetWindow(HWND* p) override { if (!p) return E_POINTER; *p = m_inst->hwnd; return m_inst->hwnd ? S_OK : E_FAIL; }
	STDMETHODIMP ContextSensitiveHelp(BOOL) override { return E_NOTIMPL; }
	STDMETHODIMP InPlaceDeactivate() override {
		if (!m_inst->active) return S_OK;
		UIDeactivate();
		m_inst->active = false;
		IOleObject* pOle = NULL;
		QueryInterface(IID_IOleObject, (void**)&pOle);
		if (pOle) { OleLockRunning(pOle, FALSE, FALSE); pOle->Release(); }
		if (m_inst->dropRegistered) {
			RevokeDragDrop(m_inst->hwnd);
			m_inst->dropRegistered = false;
		}
		if (m_inst->hwnd) {
			DestroyWindow(m_inst->hwnd);
			m_inst->hwnd = NULL;
		}
		if (m_inst->inPlaceSite) m_inst->inPlaceSite->OnInPlaceDeactivate();
		return S_OK;
	}
	STDMETHODIMP UIDeactivate() override {
		if (!m_inst->uiActive) return S_OK;
		m_inst->uiActive = false;
		if (m_inst->inPlaceSite) m_inst->inPlaceSite->OnUIDeactivate(FALSE);
		return S_OK;
	}
	STDMETHODIMP SetObjectRects(LPCRECT pPos, LPCRECT) override {
		if (pPos) {
			m_inst->rcPos = *pPos;
			if (m_inst->hwnd)
				SetWindowPos(m_inst->hwnd, HWND_TOP,
					pPos->left, pPos->top,
					pPos->right - pPos->left, pPos->bottom - pPos->top,
					SWP_SHOWWINDOW);
		}
		return S_OK;
	}
	STDMETHODIMP ReactivateAndUndo() override { return E_NOTIMPL; }

	STDMETHODIMP TranslateAccelerator(LPMSG) override { return S_FALSE; }
	STDMETHODIMP OnFrameWindowActivate(BOOL) override { return S_OK; }
	STDMETHODIMP OnDocWindowActivate(BOOL) override { return S_OK; }
	STDMETHODIMP ResizeBorder(LPCRECT, IOleInPlaceUIWindow*, BOOL) override { return S_OK; }
	STDMETHODIMP EnableModeless(BOOL) override { return S_OK; }

	STDMETHODIMP GetControlInfo(CONTROLINFO* p) override {
		if (!p) return E_POINTER;
		p->cb = sizeof(CONTROLINFO);
		p->hAccel = NULL;
		p->cAccel = 0;
		p->dwFlags = NULL;
		return S_OK;
	}
	STDMETHODIMP OnMnemonic(MSG*) override { return E_NOTIMPL; }
	STDMETHODIMP OnAmbientPropertyChange(DISPID) override { return S_OK; }
	STDMETHODIMP FreezeEvents(BOOL) override { return S_OK; }

	STDMETHODIMP Draw(DWORD asp, LONG, void*, DVTARGETDEVICE*, HDC, HDC hdc, LPCRECTL pBounds, LPCRECTL, BOOL(STDMETHODCALLTYPE*)(ULONG_PTR), ULONG_PTR) override {
		if (asp != DVASPECT_CONTENT || !pBounds || !hdc) return E_INVALIDARG;
		RECT rc = { pBounds->left, pBounds->top, pBounds->right, pBounds->bottom };
		if (m_inst->hwnd && IsWindow(m_inst->hwnd))
			SendMessage(m_inst->hwnd, WM_PRINT, (WPARAM)hdc, PRF_CLIENT | PRF_CHILDREN);
		else
			FillRect(hdc, &rc, (HBRUSH)GetStockObject(WHITE_BRUSH));
		return S_OK;
	}
	STDMETHODIMP GetColorSet(DWORD, LONG, void*, DVTARGETDEVICE*, HDC, LOGPALETTE**) override { return E_NOTIMPL; }
	STDMETHODIMP Freeze(DWORD, LONG, void*, DWORD*) override { return E_NOTIMPL; }
	STDMETHODIMP Unfreeze(DWORD) override { return E_NOTIMPL; }
	STDMETHODIMP SetAdvise(DWORD, DWORD, IAdviseSink*) override { return S_OK; }
	STDMETHODIMP GetAdvise(DWORD*, DWORD*, IAdviseSink**) override { return E_NOTIMPL; }
	STDMETHODIMP GetExtent(DWORD, LONG, DVTARGETDEVICE*, LPSIZEL p) override { return GetExtent(DVASPECT_CONTENT, p); }

	STDMETHODIMP GetClassID(CLSID* p) override { if (!p) return E_POINTER; *p = m_inst->reg->clsid; return S_OK; }
	STDMETHODIMP IsDirty() override { return S_FALSE; }
	STDMETHODIMP Load(IStream*) override { return S_OK; }
	STDMETHODIMP Save(IStream*, BOOL) override { return S_OK; }
	STDMETHODIMP GetSizeMax(ULARGE_INTEGER* p) override { if (!p) return E_POINTER; p->QuadPart = NULL; return S_OK; }
	STDMETHODIMP InitNew() override { return S_OK; }

	// ---- IObjectSafety ----
	// IE asks "are you safe to run without security warnings?"
	// We say YES to both UNTRUSTED_CALLER (safe for scripting) and UNTRUSTED_DATA
	// (safe for initialization). This is pure self-declaration - there's no
	// verification, no signature, no certificate. This is literally why ActiveX
	// got such a bad security reputation: any control can say "trust me" and IE believes it.
	// For us it's fine because our controls run in our own process, not from the internet.
	STDMETHODIMP GetInterfaceSafetyOptions(REFIID, DWORD* sup, DWORD* en) override {
		if (!sup || !en) return E_POINTER;
		*sup = *en = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
		return S_OK;
	}
	STDMETHODIMP SetInterfaceSafetyOptions(REFIID, DWORD, DWORD) override { return S_OK; }

	// ---- IProvideClassInfo2 ----
	// GetClassInfo returns our CoclassTypeInfo. IE navigates:
	//   coclass (TKIND_COCLASS) -> impl type [source,default] -> EventTypeInfo (TKIND_DISPATCH)
	//   -> GetIDsOfNames("OnClick") -> DISPID
	// This enables IE's native attachEvent() to discover our custom event DISPIDs.
	STDMETHODIMP GetClassInfo(ITypeInfo** ppTI) override {
		if (!ppTI) return E_POINTER;
		if (!m_inst->reg->coclassInfo) { *ppTI = NULL; return E_NOTIMPL; }
		*ppTI = m_inst->reg->coclassInfo;
		m_inst->reg->coclassInfo->AddRef();
		return S_OK;
	}
	STDMETHODIMP GetGUID(DWORD dwGuidKind, GUID* pGUID) override {
		if (!pGUID) return E_POINTER;
		if (dwGuidKind == GUIDKIND_DEFAULT_SOURCE_DISP_IID) { *pGUID = IID_IDispatch; return S_OK; }
		return E_INVALIDARG;
	}

	// ---- IDispatch (control methods/properties) ----
	// Same pattern as ExternalDispatch, but for the control itself.
	// JS calls canvas1.Add(3,7) -> GetIDsOfNames("Add") -> Invoke(dispid, [3,7])
	// Methods use ControlMethodCallback which includes the hControl parameter,
	// so your callback knows WHICH control instance is being called.
	STDMETHODIMP GetTypeInfoCount(UINT* p) override { if (!p) return E_POINTER; *p = 0; return S_OK; }
	STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
	STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR* names, UINT cnt, LCID, DISPID* ids) override {
		HRESULT hr = S_OK;
		EnterCriticalSection(&m_inst->reg->cs);
		for (UINT i = 0; i < cnt; i++) {
			auto it = m_inst->reg->name2id.find(names[i]);
			if (it != m_inst->reg->name2id.end()) {
				ids[i] = it->second;
			}
			else {
				ids[i] = DISPID_UNKNOWN;
				hr = DISP_E_UNKNOWNNAME;
			}
		}
		LeaveCriticalSection(&m_inst->reg->cs);
		return hr;
	}
	STDMETHODIMP Invoke(DISPID id, REFIID riid, LCID, WORD wf, DISPPARAMS* dp, VARIANT* res, EXCEPINFO*, UINT*) override {
		if (riid != IID_NULL) return E_INVALIDARG;
		EnterCriticalSection(&m_inst->reg->cs);
		auto it = m_inst->reg->entries.find(id);
		if (it == m_inst->reg->entries.end()) {
			LeaveCriticalSection(&m_inst->reg->cs);
			return DISP_E_MEMBERNOTFOUND;
		}
		DispEntry entry = it->second;
		LeaveCriticalSection(&m_inst->reg->cs);
		if ((wf & DISPATCH_METHOD) && entry.type == DispEntry::METHOD) {
			if (res) VariantInit(res);
			return entry.method(m_inst, dp, res);
		}
		if ((wf & DISPATCH_PROPERTYGET) && entry.type == DispEntry::PROPERTY && entry.getter) {
			if (res) VariantInit(res);
			return entry.getter(m_inst, res);
		}
		if ((wf & DISPATCH_PROPERTYPUT) && entry.type == DispEntry::PROPERTY && entry.setter) {
			if (dp->cArgs > 0) return entry.setter(m_inst, dp->rgvarg[0]);
			return E_INVALIDARG;
		}
		return DISP_E_MEMBERNOTFOUND;
	}

	STDMETHODIMP EnumConnectionPoints(IEnumConnectionPoints**) override { return E_NOTIMPL; }
	STDMETHODIMP FindConnectionPoint(REFIID riid, IConnectionPoint** ppCP) override;

	STDMETHODIMP DragEnter(IDataObject* pObj, DWORD keys, POINTL pt, DWORD* eff) override {
		if (!eff) return E_POINTER;
		if (m_inst->onDragEnter) return m_inst->onDragEnter(m_inst, pObj, keys, pt, eff);
		*eff = DROPEFFECT_NONE; return S_OK;
	}
	STDMETHODIMP DragOver(DWORD keys, POINTL pt, DWORD* eff) override {
		if (!eff) return E_POINTER;
		if (m_inst->onDragOver) return m_inst->onDragOver(m_inst, keys, pt, eff);
		*eff = DROPEFFECT_NONE; return S_OK;
	}
	STDMETHODIMP DragLeave(void) override {
		if (m_inst->onDragLeave) return m_inst->onDragLeave(m_inst);
		return S_OK;
	}
	STDMETHODIMP Drop(IDataObject* pObj, DWORD keys, POINTL pt, DWORD* eff) override {
		if (!eff) return E_POINTER;
		if (m_inst->onDrop) return m_inst->onDrop(m_inst, pObj, keys, pt, eff);
		*eff = DROPEFFECT_NONE; return S_OK;
	}

	STDMETHODIMP QueryContinueDrag(BOOL fEsc, DWORD keys) override {
		if (fEsc) return DRAGDROP_S_CANCEL;
		if (!(keys & MK_LBUTTON)) return DRAGDROP_S_DROP;
		return S_OK;
	}
	STDMETHODIMP GiveFeedback(DWORD) override { return DRAGDROP_S_USEDEFAULTCURSORS; }
};

// CtrlWndProc - the internal window procedure for ActiveX control child windows.
// This is the bridge between Win32 messages and the user's ControlWndProc callback.
// The ActiveXControl* is stored in GWLP_USERDATA at WM_NCCREATE time.
// We clear it in WM_NCDESTROY to prevent dangling pointer access.
inline LRESULT CALLBACK CtrlWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
	ActiveXControl* ctrl = NULL;
	if (msg == WM_NCCREATE) {
		ctrl = (ActiveXControl*)((CREATESTRUCT*)lp)->lpCreateParams;
		SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)ctrl);
		ctrl->m_inst->hwnd = hwnd; // Set early so WM_CREATE handlers can use GetControlHWND/EnableDragDrop
	}
	else {
		ctrl = (ActiveXControl*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
	}
	if (msg == WM_NCDESTROY) {
		SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
		return 0;
	}
	if (ctrl && ctrl->m_inst && ctrl->m_inst->reg && ctrl->m_inst->reg->wndProc) {
		bool implemented = false;
		LRESULT r = ctrl->m_inst->reg->wndProc(ctrl->m_inst, hwnd, msg, wp, lp, &implemented);
		if (implemented) return r;
	}
	return DefWindowProc(hwnd, msg, wp, lp);
}

// CtrlConnectionPoint - handles event subscriptions.
// 
// COM events work through "connection points": the control says "I fire events through
// IDispatch", and IE says "here's my IDispatch sink - call Invoke() on it when something
// happens". When JS does obj.attachEvent("OnClick", handler), IE calls:
//  	1. FindConnectionPoint(IID_IDispatch) -> gets a CtrlConnectionPoint
//  	2. Advise(handler_dispatch) -> we store the handler in our sink list
// 
// When C++ calls FireEvent(), it iterates all sinks and calls Invoke() on each.
// This object holds a strong reference (AddRef) to the parent ActiveXControl via m_pCPC,
// ensuring the control stays alive as long as anyone holds this connection point.
class CtrlConnectionPoint : public IConnectionPoint {
	LONG m_ref = 0;
	ControlInstance_* m_inst = NULL;
	IConnectionPointContainer* m_pCPC = NULL;
public:
	CtrlConnectionPoint(ControlInstance_* inst, IConnectionPointContainer* pCPC) : m_inst(inst), m_pCPC(pCPC) {
		InterlockedIncrement(&m_ref);
		if (m_pCPC) m_pCPC->AddRef();
	}
	~CtrlConnectionPoint() { if (m_pCPC) m_pCPC->Release(); }
	STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
		if (riid == IID_IUnknown || riid == IID_IConnectionPoint) {
			*ppv = this; AddRef();
			return S_OK;
		}
		*ppv = NULL;
		return E_NOINTERFACE;
	}
	STDMETHODIMP_(ULONG) AddRef() override {
		return InterlockedIncrement(&m_ref);
	}
	STDMETHODIMP_(ULONG) Release() override {
		LONG r = InterlockedDecrement(&m_ref);
		if (!r) delete this;
		return r;
	}
	STDMETHODIMP GetConnectionInterface(IID* p) override {
		if (!p) return E_POINTER;
		*p = IID_IDispatch;
		return S_OK;
	}
	STDMETHODIMP GetConnectionPointContainer(IConnectionPointContainer** ppCPC) override {
		if (!ppCPC) return E_POINTER;
		*ppCPC = m_pCPC;
		if (*ppCPC) (*ppCPC)->AddRef();
		return *ppCPC ? S_OK : E_UNEXPECTED;
	}
	STDMETHODIMP Advise(IUnknown* pSink, DWORD* pCookie) override {
		if (!pSink || !pCookie) return E_POINTER;
		IDispatch* disp = NULL;
		if (FAILED(pSink->QueryInterface(IID_IDispatch, (void**)&disp))) return CONNECT_E_CANNOTCONNECT;
		*pCookie = ++m_inst->nextCookie;
		m_inst->sinks[*pCookie] = disp;
		return S_OK;
	}
	STDMETHODIMP Unadvise(DWORD cookie) override {
		auto it = m_inst->sinks.find(cookie);
		if (it == m_inst->sinks.end()) return CONNECT_E_NOCONNECTION;
		it->second->Release();
		m_inst->sinks.erase(it);
		return S_OK;
	}
	STDMETHODIMP EnumConnections(IEnumConnections**) override { return E_NOTIMPL; }
};

inline STDMETHODIMP ActiveXControl::FindConnectionPoint(REFIID riid, IConnectionPoint** ppCP) {
	if (!ppCP) return E_POINTER;
	if (riid == IID_IDispatch) {
		*ppCP = new CtrlConnectionPoint(m_inst, static_cast<IConnectionPointContainer*>(this));
		return S_OK;
	}
	*ppCP = NULL;
	return CONNECT_E_NOCONNECTION;
}

// CtrlClassFactory - creates ActiveXControl instances when IE encounters <object> tags.
// 
// When IE sees <object classid="clsid:{...}">, it calls CoCreateInstance with that CLSID.
// COM looks up the class factory we registered (via CoRegisterClassObject in RegisterControl),
// and calls CreateInstance() on it. We create a new ActiveXControl and return it.
// This is a pure in-memory registration - no registry entries are written.
class CtrlClassFactory : public IClassFactory {
	LONG m_ref = 0;
	ControlReg_* m_reg = NULL;
public:
	CtrlClassFactory(ControlReg_* r) : m_reg(r) { InterlockedIncrement(&m_ref); }
	STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
		if (riid == IID_IUnknown || riid == IID_IClassFactory) {
			*ppv = this; AddRef();
			return S_OK;
		}
		*ppv = NULL;
		return E_NOINTERFACE;
	}
	STDMETHODIMP_(ULONG) AddRef() override { return InterlockedIncrement(&m_ref); }
	STDMETHODIMP_(ULONG) Release() override {
		LONG r = InterlockedDecrement(&m_ref);
		if (!r) delete this;
		return r;
	}
	STDMETHODIMP CreateInstance(IUnknown* pOuter, REFIID riid, void** ppv) override {
		if (pOuter) return CLASS_E_NOAGGREGATION;
		ActiveXControl* c = new ActiveXControl(m_reg);
		HRESULT hr = c->QueryInterface(riid, ppv);
		c->Release();
		return hr;
	}
	STDMETHODIMP LockServer(BOOL) override { return S_OK; }
};

// SimpleDataObject - a minimal IDataObject that holds Unicode text.
// Used for OLE drag & drop. Usage:
//  	IDataObject* pDO = CreateTextDataObject(L"hello");
//  	DWORD effect;
//  	DoDragDrop(pDO, pDropSource, DROPEFFECT_COPY, &effect);
//  	pDO->Release();
class SimpleDataObject : public IDataObject {
	LONG m_ref = 0;
	std::wstring m_text;
public:
	SimpleDataObject(const wchar_t* text) : m_text(text ? text : L"") {
		InterlockedIncrement(&m_ref);
	}
	STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
		if (riid == IID_IUnknown || riid == IID_IDataObject) {
			*ppv = this; AddRef();
			return S_OK;
		}
		*ppv = NULL;
		return E_NOINTERFACE;
	}
	STDMETHODIMP_(ULONG) AddRef() override {
		return InterlockedIncrement(&m_ref);
	}
	STDMETHODIMP_(ULONG) Release() override {
		LONG r = InterlockedDecrement(&m_ref);
		if (!r) delete this;
		return r;
	}
	STDMETHODIMP GetData(FORMATETC* pFmt, STGMEDIUM* pMed) override {
		if (!pFmt || !pMed) return E_POINTER;
		if (pFmt->cfFormat != CF_UNICODETEXT || !(pFmt->tymed & TYMED_HGLOBAL)) return DV_E_FORMATETC;
		size_t bytes = (m_text.size() + 1) * sizeof(wchar_t);
		HGLOBAL hg = GlobalAlloc(GMEM_MOVEABLE, bytes);
		if (!hg) return E_OUTOFMEMORY;
		void* dst = GlobalLock(hg);
		if (!dst) { GlobalFree(hg); return E_OUTOFMEMORY; }
		memcpy(dst, m_text.c_str(), bytes);
		GlobalUnlock(hg);
		pMed->tymed = TYMED_HGLOBAL; pMed->hGlobal = hg; pMed->pUnkForRelease = NULL;
		return S_OK;
	}
	STDMETHODIMP GetDataHere(FORMATETC*, STGMEDIUM*) override { return E_NOTIMPL; }
	STDMETHODIMP QueryGetData(FORMATETC* p) override { if (!p) return E_POINTER; return (p->cfFormat == CF_UNICODETEXT && (p->tymed & TYMED_HGLOBAL)) ? S_OK : DV_E_FORMATETC; }
	STDMETHODIMP GetCanonicalFormatEtc(FORMATETC*, FORMATETC* p) override { if (!p) return E_POINTER; p->ptd = NULL; return DATA_S_SAMEFORMATETC; }
	STDMETHODIMP SetData(FORMATETC*, STGMEDIUM*, BOOL) override { return E_NOTIMPL; }
	STDMETHODIMP EnumFormatEtc(DWORD, IEnumFORMATETC**) override { return E_NOTIMPL; }
	STDMETHODIMP DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) override { return OLE_E_ADVISENOTSUPPORTED; }
	STDMETHODIMP DUnadvise(DWORD) override { return OLE_E_ADVISENOTSUPPORTED; }
	STDMETHODIMP EnumDAdvise(IEnumSTATDATA**) override { return OLE_E_ADVISENOTSUPPORTED; }
};

// ActiveX API implementations
// These are the user-facing functions that wrap all the COM complexity above.

inline std::wstring CLSIDToString(const CLSID& clsid) { wchar_t buf[64]; StringFromGUID2(clsid, buf, 64); return buf; }
inline std::wstring GetControlClassId(hControlClass reg) { return reg ? L"clsid:" + CLSIDToString(reg->clsid) : L""; }
inline CLSID GetControlCLSID(hControlClass reg) { return reg ? reg->clsid : CLSID_NULL; }

// RegisterControl - the main entry point for creating a new control type.
// This does NOT create a control instance - it registers a "class" that IE can
// instantiate later when it encounters an <object> tag with the matching CLSID.
// The CLSID is deterministic: same name always produces the same GUID.
inline hControlClass RegisterControl(const wchar_t* name, ControlWndProc proc, void* userData, DWORD style) {
	if (!name || !proc) return NULL;
	CLSID clsid = CLSIDFromName(name);
	if (CtrlRegistry().find(clsid) != CtrlRegistry().end()) return NULL;
	auto* reg = new ControlReg_();
	reg->name = name;
	reg->wndProc = proc;
	reg->defaultUserData = userData;
	reg->extraStyle = style;
	reg->wndClassName = std::wstring(L"TridentControl_") + name;
	reg->clsid = clsid;
	reg->eventInfo = new EventTypeInfo(clsid);
	reg->coclassInfo = new CoclassTypeInfo(clsid, reg->eventInfo);
	auto* factory = new CtrlClassFactory(reg);
	HRESULT hr = CoRegisterClassObject(reg->clsid, factory, CLSCTX_INPROC_SERVER, REGCLS_MULTIPLEUSE, &reg->comCookie);
	factory->Release();
	if (FAILED(hr)) { delete reg; return NULL; }
	CtrlRegistry()[reg->clsid] = reg;
	return reg;
}

inline void UnregisterControl(hControlClass reg) {
	if (!reg) return;
	CoRevokeClassObject(reg->comCookie);
	CtrlRegistry().erase(reg->clsid);
	reg->Release();
}

inline HWND GetControlHWND(hControl c) { return c ? c->hwnd : NULL; }
inline void* GetControlUserData(hControl c) { return c ? c->userData : NULL; }
inline void SetControlUserData(hControl c, void* d) { if (c) c->userData = d; }
inline void InvalidateControl(hControl c) { if (c && c->hwnd) ::InvalidateRect(c->hwnd, NULL, TRUE); }
inline hControlClass GetControlClass(hControl c) { return c ? c->reg : NULL; }

inline void BindClassMethod(hControlClass reg, const wchar_t* name, ControlMethodCallback cb) {
	if (!reg) return;
	EnterCriticalSection(&reg->cs);
	if (reg->name2id.count(name)) {
		LeaveCriticalSection(&reg->cs);
		throw std::runtime_error("Method already bound");
	}
	if (reg->nextDispId <= 0) {
		LeaveCriticalSection(&reg->cs);
		throw std::runtime_error("DISPID exhausted");
	}
	DISPID id = reg->nextDispId--;
	reg->name2id[name] = id;
	reg->entries[id] = { DispEntry::METHOD, cb, NULL, NULL };
	LeaveCriticalSection(&reg->cs);
}

inline void BindClassProperty(hControlClass reg, const wchar_t* name, PropertyGetter getter, PropertySetter setter) {
	if (!reg) return;
	EnterCriticalSection(&reg->cs);
	auto it = reg->name2id.find(name);
	DISPID id;
	if (it != reg->name2id.end()) {
		id = it->second;
		auto eit = reg->entries.find(id);
		if (eit != reg->entries.end() && eit->second.type == DispEntry::PROPERTY) {
			if (getter) eit->second.getter = getter;
			if (setter) eit->second.setter = setter;
			LeaveCriticalSection(&reg->cs);
			return;
		}
		LeaveCriticalSection(&reg->cs);
		throw std::runtime_error("Name already bound as method");
	}
	if (reg->nextDispId <= 0) {
		LeaveCriticalSection(&reg->cs);
		throw std::runtime_error("DISPID exhausted");
	}
	id = reg->nextDispId--;
	reg->name2id[name] = id;
	reg->entries[id] = { DispEntry::PROPERTY, ControlMethodCallback(), getter, setter };
	LeaveCriticalSection(&reg->cs);
}

inline DISPID RegisterEvent(hControlClass reg, const wchar_t* name, unsigned int paramCount) {
	if (!reg || !reg->eventInfo) return DISPID_UNKNOWN;
	return reg->eventInfo->RegisterEvent(name, paramCount);
}

inline unsigned int GetEventExpectedArgc(hControlClass reg, DISPID eventId) {
	if (!reg || !reg->eventInfo) return 0;
	return reg->eventInfo->GetParamCount(eventId);
}

inline unsigned int GetEventExpectedArgcByName(hControlClass reg, const wchar_t* name) {
	if (!reg || !reg->eventInfo || !name) return 0;
	DISPID id = reg->eventInfo->GetEventId(name);
	if (id == DISPID_UNKNOWN) return 0;
	return reg->eventInfo->GetParamCount(id);
}

// FireEvent - sends an event from C++ to all connected JavaScript sinks.
// 
// How COM events work:
//  	1. IE subscribes to events via IConnectionPoint::Advise(), giving us an IDispatch* "sink"
//  	2. When we want to fire an event, we call sink->Invoke(eventId, args)
//  	3. IE's script engine routes this to the JavaScript event handler
// 
// IMPORTANT: args must be in COM reverse order (same as DISPPARAMS convention).
//  	args[0] = last parameter, args[argc-1] = first parameter.
//  	The parameter count is automaticaly looked up from the EventTypeInfo registry (set by RegisterEvent).
// 
// We take a snapshot of the sink list before iterating, with AddRef on each sink.
// This prevents crashes if a sink calls Unadvise() on itself during the event callback
// (which would modify the map while we're iterating it).
inline void FireEvent(hControl c, DISPID eventId, VARIANT* args) {
	if (!c || !c->reg || !c->reg->eventInfo) return;
	unsigned int argc = c->reg->eventInfo->GetParamCount(eventId);
	if (argc > 0 && !args) return;
	DISPPARAMS dp = {};
	dp.rgvarg = args;
	dp.cArgs = argc;
	// Snapshot sinks with AddRef to survive re-entrant Unadvise
	std::vector<IDispatch*> snapshot;
	snapshot.reserve(c->sinks.size());
	for (auto& kv : c->sinks) {
		if (kv.second) { kv.second->AddRef(); snapshot.push_back(kv.second); }
	}
	for (IDispatch* sink : snapshot) {
		sink->Invoke(eventId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, NULL, NULL, NULL);
		sink->Release();
	}
}

inline void FireEventByName(hControl c, const wchar_t* name, VARIANT* args) {
	if (!c || !c->reg || !c->reg->eventInfo) return;
	DISPID id = c->reg->eventInfo->GetEventId(name);
	if (id == DISPID_UNKNOWN) return;
	FireEvent(c, id, args);
}

inline hTrident FindHostFromControl(hControl c) {
	if (!c) return NULL;
	return FindNearestTridentHostFromHWND(GetControlHWND(c));
}

inline void EnableDragDrop(hControl c, DropCallback onDrop, DropEnterCallback onDragEnter, DropOverCallback onDragOver, DropLeaveCallback onDragLeave) {
	if (!c || !c->hwnd) return;
	c->onDrop = onDrop; c->onDragEnter = onDragEnter; c->onDragOver = onDragOver; c->onDragLeave = onDragLeave;
	if (!c->dropRegistered) {
		ActiveXControl* ctrl = (ActiveXControl*)GetWindowLongPtr(c->hwnd, GWLP_USERDATA);
		if (ctrl) {
			HRESULT hr = RegisterDragDrop(c->hwnd, static_cast<IDropTarget*>(ctrl));
			if (SUCCEEDED(hr) || hr == DRAGDROP_E_ALREADYREGISTERED)
				c->dropRegistered = true;
		}
	}
}

inline void DisableDragDrop(hControl c) {
	if (!c || !c->hwnd || !c->dropRegistered) return;
	RevokeDragDrop(c->hwnd);
	c->dropRegistered = false;
	c->onDrop = NULL;
	c->onDragEnter = NULL;
	c->onDragOver = NULL;
	c->onDragLeave = NULL;
}

inline HRESULT StartDragDrop(hControl c, IDataObject* data, DWORD allowedEffects, DWORD* outEffect) {
	if (!c || !c->hwnd || !data) return E_INVALIDARG;
	ActiveXControl* ctrl = (ActiveXControl*)GetWindowLongPtr(c->hwnd, GWLP_USERDATA);
	if (!ctrl) return E_FAIL;
	HRESULT hr = DoDragDrop(data, static_cast<IDropSource*>(ctrl), allowedEffects, outEffect);
	return hr;
}