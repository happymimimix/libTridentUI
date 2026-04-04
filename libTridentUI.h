/*
libTridentUI - v0.2.1 Beta
An extremely lightweight Chrome Embedded Framework alternative for legacy operating systems.
Works perfectly on Microsoft Windows XP SP3!
Made possible by: Claude Opus 4.6 and Happy_mimimix

Thread safety notice:
All UI calls must be made from the same STA thread that called TridentInit().
This is a COM/OLE requirement, not a library limitation.
*/
#pragma once

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
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
#define LONG_MAX_PATH 0x0FFF

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "uuid.lib")
#pragma comment(lib, "user32.lib")

// Public APIs
struct TridentWindowData_;
typedef TridentWindowData_* hTrident;
typedef IHTMLElement* hElement;

struct DispArgs {
    DISPPARAMS* params;
    
    size_t Count() const { return params ? params->cArgs : 0; }
    
    VARIANT& operator[](size_t i) {
        return params->rgvarg[params->cArgs - 1 - i];
    }
    
    std::wstring GetString(size_t i) {
        if (i >= Count()) return L"";
        VARIANT& v = (*this)[i];
        return (v.vt == VT_BSTR && v.bstrVal) ? v.bstrVal : L"";
    }
    
    int GetInt(size_t i, int def = 0) {
        if (i >= Count()) return def;
        VARIANT& v = (*this)[i];
        if (v.vt == VT_I4) return v.lVal;
        if (v.vt == VT_I2) return v.iVal;
        if (v.vt == VT_R8) return (int)v.dblVal;
        if (v.vt == VT_R4) return (int)v.fltVal;
        return def;
    }
    
    double GetDouble(size_t i, double def = 0.0) {
        if (i >= Count()) return def;
        VARIANT& v = (*this)[i];
        if (v.vt == VT_R8) return v.dblVal;
        if (v.vt == VT_R4) return (double)v.fltVal;
        if (v.vt == VT_I4) return (double)v.lVal;
        return def;
    }
    
    bool GetBool(size_t i, bool def = false) {
        if (i >= Count()) return def;
        VARIANT& v = (*this)[i];
        return (v.vt == VT_BOOL) ? (v.boolVal != VARIANT_FALSE) : def;
    }
};

struct ControlInstance_;
typedef ControlInstance_* hControl;
struct ControlReg_;
typedef ControlReg_* hControlClass;

using MethodCallback = std::function<HRESULT(DispArgs&, VARIANT*)>;
using ControlMethodCallback = std::function<HRESULT(hControl, DispArgs&, VARIANT*)>;

typedef LRESULT (*TridentWndProc)(hTrident h, HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool* handled);
typedef LRESULT (*ControlWndProc)(hControl c, HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
typedef HRESULT (*PropertyGetter)(hControl c, VARIANT* result);
typedef HRESULT (*PropertySetter)(hControl c, VARIANT value);
typedef HRESULT (*DropEnterCallback)(hControl c, IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);
typedef HRESULT (*DropOverCallback)(hControl c, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);
typedef void (*DropLeaveCallback)(hControl c);
typedef HRESULT (*DropCallback)(hControl c, IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect);

inline void TridentInit();
inline void TridentShutdown();
inline void RunMessageLoop();

inline hTrident NewTridentWindow(const wchar_t* title, int x, int y, int w, int h, HWND owner = NULL, DWORD style = WS_OVERLAPPEDWINDOW);
inline void CloseTridentWindow(hTrident h);
inline HWND GetTridentHWND(hTrident h);
inline IWebBrowser2* GetTridentBrowser(hTrident h);

inline void BindFunction(hTrident h, const wchar_t* name, MethodCallback cb);

inline void NavigateTo(hTrident h, const wchar_t* url);
inline void NavigateToRes(hTrident h, const wchar_t* resName);
inline void NavigateToHTML(hTrident h, const wchar_t* html);

inline void ExecScript(hTrident h, const wchar_t* js);

inline hElement GetElement(hTrident h, const wchar_t* id);
inline void ReleaseElement(hElement e);

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

inline void SetWndProc(hTrident h, TridentWndProc proc);
inline void SetUserData(hTrident h, void* data);
inline void* GetUserData(hTrident h);

inline hControlClass RegisterControl(const wchar_t* name, ControlWndProc proc, void* userData = NULL, DWORD style = 0);
inline void UnregisterControl(hControlClass reg);
inline CLSID GetControlCLSID(hControlClass reg);
inline std::wstring CLSIDToString(const CLSID& clsid);
inline std::wstring GetControlClassId(hControlClass reg);
inline void BindClassMethod(hControlClass reg, const wchar_t* name, ControlMethodCallback cb);
inline void BindClassProperty(hControlClass reg, const wchar_t* name, PropertyGetter getter, PropertySetter setter);
inline DISPID RegisterEvent(hControlClass reg, const wchar_t* name);
inline void FireEvent(hControl c, DISPID eventId, VARIANT* args = NULL, int argc = 0);
inline void FireEventByName(hControl c, const wchar_t* name, VARIANT* args = NULL, int argc = 0);
inline HWND GetControlHWND(hControl c);
inline void* GetControlUserData(hControl c);
inline void SetControlUserData(hControl c, void* data);
inline void InvalidateControl(hControl c);
inline hControlClass GetControlClass(hControl c);
inline void EnableDragDrop(hControl c, DropCallback onDrop, DropEnterCallback onDragEnter = NULL, DropOverCallback onDragOver = NULL, DropLeaveCallback onDragLeave = NULL);
inline void DisableDragDrop(hControl c);
inline IDataObject* CreateTextDataObject(const wchar_t* text);

// COM implementations
inline const wchar_t* g_trident_wndclass = L"libTridentUI";
inline TridentWindowData_* g_trident_head = NULL;

class ExternalDispatch : public IDispatch {
    LONG m_ref = 0;
    CRITICAL_SECTION m_cs;
    DISPID m_nextId = INT32_MAX;
    std::map<std::wstring, DISPID> m_name2id;
    std::map<DISPID, MethodCallback> m_id2func;
    
public:
    ExternalDispatch()  { InitializeCriticalSection(&m_cs); }
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
    STDMETHODIMP_(ULONG) AddRef()  override { return InterlockedIncrement(&m_ref); }
    STDMETHODIMP_(ULONG) Release() override { return InterlockedDecrement(&m_ref); }

    STDMETHODIMP GetTypeInfoCount(UINT* p) override { if (!p) return E_POINTER; *p = 0; return S_OK; }
    STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
    
    STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR* names, UINT cnt, LCID, DISPID* ids) override {
        HRESULT hr = S_OK;
        EnterCriticalSection(&m_cs);
        for (UINT i = 0; i < cnt; i++) {
            auto it = m_name2id.find(names[i]);
            if (it != m_name2id.end()) {
                ids[i] = it->second;
            } else {
                ids[i] = DISPID_UNKNOWN;
                hr = DISP_E_UNKNOWNNAME;
            }
        }
        LeaveCriticalSection(&m_cs);
        return hr;
    }
    
    STDMETHODIMP Invoke(DISPID id, REFIID, LCID, WORD wf, DISPPARAMS* dp,
                        VARIANT* res, EXCEPINFO*, UINT*) override {
        if (!(wf & DISPATCH_METHOD)) return DISP_E_MEMBERNOTFOUND;
        EnterCriticalSection(&m_cs);
        auto it = m_id2func.find(id);
        if (it == m_id2func.end()) { LeaveCriticalSection(&m_cs); return DISP_E_MEMBERNOTFOUND; }
        auto fn = it->second;
        LeaveCriticalSection(&m_cs);
        DispArgs args = { dp };
        if (res) VariantInit(res);
        return fn(args, res);
    }
};

class OleSite :
    public IOleClientSite,
    public IOleInPlaceSiteWindowless,
    public IOleControlSite,
    public IObjectWithSite,
    public IOleContainer,
    public IDocHostUIHandler,
    public IDispatch
{
public:
    ExternalDispatch ext;
    
    OleSite() : m_ref(1), m_hwnd(NULL), m_pUnkSite(NULL),
        m_pView(NULL), m_pInPlace(NULL), m_pOle(NULL), m_pBrowser(NULL),
        m_cookie(0), m_bWindowless(true), m_bFocused(false), m_bCaptured(false),
        m_bUIActivated(false), m_bInPlaceActive(false), m_bLocked(false) {}
    
    ~OleSite() {
        Destroy();
        if (m_pUnkSite) m_pUnkSite->Release();
        if (m_pView) m_pView->Release();
        if (m_pInPlace) m_pInPlace->Release();
    }
    
    IWebBrowser2* GetBrowser() { return m_pBrowser; }
    
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

    bool Create(HWND hwnd) {
        m_hwnd = hwnd;

        HRESULT hr = CoCreateInstance(CLSID_WebBrowser, NULL, CLSCTX_ALL, IID_IOleObject, (void**)&m_pOle);
        if (FAILED(hr)) return false;

        m_pOle->SetClientSite(static_cast<IOleClientSite*>(this));

        IPersistStreamInit* psi = NULL;
        m_pOle->QueryInterface(IID_IPersistStreamInit, (void**)&psi);
        if (psi) {
            psi->InitNew();
            psi->Release();
        }

        hr = m_pOle->QueryInterface(IID_IViewObjectEx, (void**)&m_pView);
        if (FAILED(hr)) hr = m_pOle->QueryInterface(IID_IViewObject2, (void**)&m_pView);
        if (FAILED(hr)) m_pOle->QueryInterface(IID_IViewObject, (void**)&m_pView);
        m_pOle->SetHostNames(OLESTR("TridentUI"), NULL);

        RECT rc;
        GetClientRect(hwnd, &rc);
        hr = m_pOle->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, static_cast<IOleClientSite*>(this), 0, hwnd, &rc);
        if (FAILED(hr)) { Destroy(); return false; }

        IObjectWithSite* pSite = NULL;
        m_pOle->QueryInterface(IID_IObjectWithSite, (void**)&pSite);
        if (pSite) {
            pSite->SetSite(static_cast<IOleClientSite*>(this));
            pSite->Release();
        }
        m_pOle->QueryInterface(IID_IWebBrowser2, (void**)&m_pBrowser);
        if (!m_pBrowser) { Destroy(); return false; }

        ConnectEvents(TRUE);
        return true;
    }
    
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

    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        *ppv = NULL;
        if (riid == IID_IUnknown) *ppv = static_cast<IOleWindow*>(this);
        else if (riid == IID_IOleClientSite) *ppv = static_cast<IOleClientSite*>(this);
        else if (riid == IID_IOleInPlaceSiteWindowless) *ppv = static_cast<IOleInPlaceSiteWindowless*>(this);
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
    STDMETHODIMP_(ULONG) AddRef()  override { return InterlockedIncrement(&m_ref); }
    STDMETHODIMP_(ULONG) Release() override { LONG r = InterlockedDecrement(&m_ref); if (!r) delete this; return r; }

    STDMETHODIMP SaveObject() override { return E_NOTIMPL; }
    STDMETHODIMP GetMoniker(DWORD, DWORD, IMoniker** p) override { if (!p) return E_POINTER; *p = NULL; return E_NOTIMPL; }
    STDMETHODIMP GetContainer(IOleContainer** p) override {
        return QueryInterface(IID_IOleContainer, (void**)p);
    }
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

    STDMETHODIMP GetWindow(HWND* p) override { if (!p) return E_POINTER; *p = m_hwnd; return S_OK; }
    STDMETHODIMP ContextSensitiveHelp(BOOL) override { return S_OK; }
    STDMETHODIMP CanInPlaceActivate() override { return S_OK; }
    STDMETHODIMP OnInPlaceActivate() override {
        BOOL dummy = FALSE;
        return OnInPlaceActivateEx(&dummy, 0);
    }
    STDMETHODIMP OnUIActivate() override { m_bUIActivated = true; return S_OK; }
    STDMETHODIMP GetWindowContext(IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc, LPRECT pPos, LPRECT pClip, LPOLEINPLACEFRAMEINFO pFI) override {
        if (!ppFrame || !ppDoc || !pPos || !pClip || !pFI) return E_POINTER;
        *ppFrame = new InlineFrame(m_hwnd);
        *ppDoc = NULL;
        GetClientRect(m_hwnd, pPos);
        *pClip = *pPos;
        pFI->cb = sizeof(OLEINPLACEFRAMEINFO);
        pFI->fMDIApp = FALSE;
        pFI->hwndFrame = m_hwnd;
        pFI->haccel = NULL;
        pFI->cAccelEntries = 0;
        return S_OK;
    }
    STDMETHODIMP Scroll(SIZE) override { return E_NOTIMPL; }
    STDMETHODIMP OnUIDeactivate(BOOL) override { m_bUIActivated = false; return S_OK; }
    STDMETHODIMP OnInPlaceDeactivate() override { return OnInPlaceDeactivateEx(TRUE); }
    STDMETHODIMP DiscardUndoState() override { return E_NOTIMPL; }
    STDMETHODIMP DeactivateAndUndo() override { return E_NOTIMPL; }
    STDMETHODIMP OnPosRectChange(LPCRECT) override { return E_NOTIMPL; }
    STDMETHODIMP OnInPlaceActivateEx(BOOL*, DWORD dwFlags) override {
        if (!m_hwnd || !m_pOle) return E_UNEXPECTED;
        OleLockRunning(m_pOle, TRUE, FALSE);
        HRESULT hr = E_FAIL;
        if (dwFlags & ACTIVATE_WINDOWLESS) {
            m_bWindowless = true;
            hr = m_pOle->QueryInterface(IID_IOleInPlaceObjectWindowless, (void**)&m_pInPlace);
        }
        if (FAILED(hr)) {
            m_bWindowless = false;
            hr = m_pOle->QueryInterface(IID_IOleInPlaceObject, (void**)&m_pInPlace);
        }
        if (m_pInPlace) {
            RECT rc; GetClientRect(m_hwnd, &rc);
            m_pInPlace->SetObjectRects(&rc, &rc);
        }
        m_bInPlaceActive = SUCCEEDED(hr);
        return hr;
    }
    STDMETHODIMP OnInPlaceDeactivateEx(BOOL) override {
        m_bInPlaceActive = false;
        if (m_pOle) OleLockRunning(m_pOle, FALSE, FALSE);
        if (m_pInPlace) {
            m_pInPlace->Release();
            m_pInPlace = NULL;
        }
        return S_OK;
    }
    STDMETHODIMP RequestUIActivate() override { return S_OK; }
    STDMETHODIMP CanWindowlessActivate() override { return S_OK; }
    STDMETHODIMP GetCapture() override { return m_bCaptured ? S_OK : S_FALSE; }
    STDMETHODIMP SetCapture(BOOL f) override {
        m_bCaptured = (f == TRUE);
        if (f) ::SetCapture(m_hwnd); else ::ReleaseCapture();
        return S_OK;
    }
    STDMETHODIMP GetFocus() override { return m_bFocused ? S_OK : S_FALSE; }
    STDMETHODIMP SetFocus(BOOL f) override {
        if (f) ::SetFocus(m_hwnd);
        m_bFocused = (f == TRUE);
        return S_OK;
    }
    STDMETHODIMP GetDC(LPCRECT, DWORD grfFlags, HDC* phDC) override {
        if (!phDC) return E_POINTER;
        *phDC = ::GetDC(m_hwnd);
        return S_OK;
    }
    STDMETHODIMP ReleaseDC(HDC h) override {
        ::ReleaseDC(m_hwnd, h);
        return S_OK;
    }
    STDMETHODIMP InvalidateRect(LPCRECT p, BOOL e) override {
        ::InvalidateRect(m_hwnd, p, e);
        return S_OK;
    }
    STDMETHODIMP InvalidateRgn(HRGN h, BOOL e) override {
        ::InvalidateRgn(m_hwnd, h, e);
        return S_OK;
    }
    STDMETHODIMP ScrollRect(INT, INT, LPCRECT, LPCRECT) override { return S_OK; }
    STDMETHODIMP AdjustRect(LPRECT) override { return S_OK; }
    STDMETHODIMP OnDefWindowMessage(UINT msg, WPARAM wp, LPARAM lp, LRESULT* pr) override {
        *pr = DefWindowProc(m_hwnd, msg, wp, lp);
        return S_OK;
    }
    STDMETHODIMP OnControlInfoChanged() override { return S_OK; }
    STDMETHODIMP LockInPlaceActive(BOOL) override { return S_OK; }
    STDMETHODIMP GetExtendedControl(IDispatch** p) override {
        if (!p || !m_pOle) return E_POINTER;
        return m_pOle->QueryInterface(IID_IDispatch, (void**)p);
    }
    STDMETHODIMP TransformCoords(POINTL*, POINTF*, DWORD) override { return S_OK; }
    STDMETHODIMP TranslateAccelerator(MSG*, DWORD) override { return S_FALSE; }
    STDMETHODIMP OnFocus(BOOL f) override { m_bFocused = (f == TRUE); return S_OK; }
    STDMETHODIMP ShowPropertyFrame() override { return E_NOTIMPL; }
    STDMETHODIMP EnumObjects(DWORD, IEnumUnknown** pp) override {
        if (!pp || !m_pOle) return E_POINTER;
        *pp = new InlineEnum(m_pOle);
        return S_OK;
    }
    STDMETHODIMP LockContainer(BOOL f) override { m_bLocked = (f != FALSE); return S_OK; }
    STDMETHODIMP ParseDisplayName(IBindCtx*, LPOLESTR, ULONG*, IMoniker**) override { return E_NOTIMPL; }
    STDMETHODIMP ShowContextMenu(DWORD, POINT*, IUnknown*, IDispatch*) override { return S_OK; }
    STDMETHODIMP GetHostInfo(DOCHOSTUIINFO* p) override {
        if (p) {
            p->cbSize = sizeof(DOCHOSTUIINFO);
            p->dwFlags = DOCHOSTUIFLAG_NO3DBORDER;
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
    STDMETHODIMP ResizeBorder(LPCRECT, IOleInPlaceUIWindow*, BOOL) override { return S_OK; }
    STDMETHODIMP TranslateAccelerator(LPMSG, const GUID*, DWORD) override { return S_FALSE; }
    STDMETHODIMP GetOptionKeyPath(LPOLESTR*, DWORD) override { return S_OK; }
    STDMETHODIMP GetDropTarget(IDropTarget*, IDropTarget**) override { return E_NOTIMPL; }
    STDMETHODIMP GetExternal(IDispatch** pp) override {
        if (!pp) return E_POINTER;
        *pp = &ext;
        ext.AddRef();
        return S_OK;
    }
    STDMETHODIMP TranslateUrl(DWORD, LPWSTR, LPWSTR* pp) override { *pp = NULL; return S_OK; }
    STDMETHODIMP FilterDataObject(IDataObject*, IDataObject** pp) override { *pp = NULL; return S_OK; }
    STDMETHODIMP GetTypeInfoCount(UINT* p) override { if (!p) return E_POINTER; *p = 0; return S_OK; }
    STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
    STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override { return E_NOTIMPL; }
    STDMETHODIMP Invoke(DISPID id, REFIID riid, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) override {
        if (riid != IID_NULL) return E_INVALIDARG;
        if (id == DISPID_NAVIGATECOMPLETE2 || id == DISPID_DOCUMENTCOMPLETE) SetupUIHandler();
        return S_OK;
    }

private:
    void ConnectEvents(BOOL advise) {
        if (!m_pBrowser) return;
        IConnectionPointContainer* cpc = NULL;
        if (FAILED(m_pBrowser->QueryInterface(IID_IConnectionPointContainer, (void**)&cpc))) return;
        IConnectionPoint* cp = NULL;
        if (SUCCEEDED(cpc->FindConnectionPoint(DIID_DWebBrowserEvents2, &cp))) {
            if (advise) {
                cp->Advise(static_cast<IDispatch*>(this), &m_cookie);
            } else if (m_cookie) {
                cp->Unadvise(m_cookie);
                m_cookie = 0;
            }
            cp->Release();
        }
        cpc->Release();
    }
    class InlineFrame : public IOleInPlaceFrame {
        ULONG m_r; HWND m_h;
    public:
        InlineFrame(HWND h) : m_r(1), m_h(h) {}
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
        STDMETHODIMP_(ULONG) AddRef()  override { return InterlockedIncrement(&m_r); }
        STDMETHODIMP_(ULONG) Release() override { LONG r = InterlockedDecrement(&m_r); if (!r) delete this; return r; }
        STDMETHODIMP GetWindow(HWND* p) override { if (!p) return E_POINTER; *p = m_h; return S_OK; }
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
        ULONG m_r; LONG m_pos; IUnknown* m_p;
    public:
        InlineEnum(IUnknown* p) : m_r(1), m_pos(0), m_p(p) { m_p->AddRef(); }
        ~InlineEnum() { m_p->Release(); }
        STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
            if (riid == IID_IUnknown || riid == IID_IEnumUnknown)
                { *ppv = static_cast<IEnumUnknown*>(this); AddRef(); return S_OK; }
            *ppv = NULL; return E_NOINTERFACE;
        }
        STDMETHODIMP_(ULONG) AddRef()  override { return InterlockedIncrement(&m_r); }
        STDMETHODIMP_(ULONG) Release() override { LONG r = InterlockedDecrement(&m_r); if (!r) delete this; return r; }
        STDMETHODIMP Next(ULONG celt, IUnknown** p, ULONG* f) override {
            if (!p) return E_POINTER;
            ULONG fetched = 0;
            while (fetched < celt && m_pos < 1) {
                p[fetched] = m_p; p[fetched]->AddRef();
                fetched++; m_pos++;
            }
            if (f) *f = fetched;
            return (fetched == celt) ? S_OK : S_FALSE;
        }
        STDMETHODIMP Skip(ULONG c) override { m_pos += c; return S_OK; }
        STDMETHODIMP Reset() override { m_pos = 0; return S_OK; }
        STDMETHODIMP Clone(IEnumUnknown**) override { return E_NOTIMPL; }
    };
    LONG m_ref;
    HWND m_hwnd;
    IUnknown* m_pUnkSite;
    IViewObject* m_pView;
    IOleInPlaceObjectWindowless* m_pInPlace;
    IOleObject* m_pOle;
    IWebBrowser2* m_pBrowser;
    DWORD m_cookie;
    bool m_bWindowless;
    bool m_bFocused;
    bool m_bCaptured;
    bool m_bUIActivated;
    bool m_bInPlaceActive;
    bool m_bLocked;
};
struct TridentWindowData_ {
    HWND hwnd = NULL;
    OleSite* site = NULL;
    TridentWndProc userProc = NULL;
    void* userData = NULL;
    bool alive = true;
    TridentWindowData_* next = NULL;
};
inline LRESULT CALLBACK TridentHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    hTrident h = (hTrident)GetWindowLongPtr(hwnd, GWLP_USERDATA);
    
    if (msg == WM_NCCREATE) {
        h = (hTrident)((CREATESTRUCT*)lp)->lpCreateParams;
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)h);
        h->hwnd = hwnd;
    }
    
    if (h) {
        if (h->userProc) {
            bool handled = false;
            LRESULT r = h->userProc(h, hwnd, msg, wp, lp, &handled);
            if (handled) return r;
        }
        if (msg == WM_SIZE && h->site) {
            h->site->Resize(LOWORD(lp), HIWORD(lp));
            return 0;
        }
        if (msg == WM_CLOSE) {
            DestroyWindow(hwnd);
            return 0;
        }
        if (msg == WM_NCDESTROY) {
            SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
            h->hwnd = NULL;
            h->alive = false;
            if (h->site) { h->site->Destroy(); h->site->Release(); h->site = NULL; }
            // Remove from chain
            TridentWindowData_** pp = &g_trident_head;
            while (*pp) {
                if (*pp == h) { *pp = h->next; break; }
                pp = &(*pp)->next;
            }
            delete h;
            return 0;
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

inline void RunMessageLoop() {
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        bool ate = false;
        for (TridentWindowData_* p = g_trident_head; p; ) {
            TridentWindowData_* next = p->next;
            if (p->alive && p->site && p->hwnd &&
                (msg.hwnd == p->hwnd || IsChild(p->hwnd, msg.hwnd))) {
                IWebBrowser2* b = p->site->GetBrowser();
                if (b) {
                    b->AddRef();
                    IOleInPlaceActiveObject* ao = NULL;
                    if (SUCCEEDED(b->QueryInterface(IID_IOleInPlaceActiveObject, (void**)&ao))) {
                        ate = (ao->TranslateAccelerator(&msg) == S_OK);
                        ao->Release();
                    }
                    b->Release();
                }
            }
            if (ate) break;
            p = next;
        }
        if (!ate) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
}

inline hTrident NewTridentWindow(const wchar_t* title, int x, int y, int w, int h, HWND owner, DWORD style) {
    TridentWindowData_* d = new TridentWindowData_();
    CreateWindowExW(0, g_trident_wndclass, title, style, x, y, w, h, owner, NULL, GetModuleHandle(NULL), d);
    if (!d->hwnd) { delete d; return NULL; }
    d->site = new OleSite();
    if (!d->site->Create(d->hwnd)) {
        d->site->Release(); d->site = NULL;
        SetWindowLongPtr(d->hwnd, GWLP_USERDATA, 0);
        DestroyWindow(d->hwnd);
        delete d;
        return NULL;
    }
    // Add to chain
    d->next = g_trident_head;
    g_trident_head = d;
    return d;
}

inline void CloseTridentWindow(hTrident h) {
    if (!h || !h->alive) return;
    if (h->hwnd) {
        DestroyWindow(h->hwnd);
    } else {
        // This shouldn't happen normally but we still handle it here anyways.
        h->alive = false;
        if (h->site) { h->site->Destroy(); h->site->Release(); h->site = NULL; }
        TridentWindowData_** pp = &g_trident_head;
        while (*pp) {
            if (*pp == h) { *pp = h->next; break; }
            pp = &(*pp)->next;
        }
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

inline void NavigateToHTML(hTrident h, const wchar_t* html) {
    IWebBrowser2* b = GetTridentBrowser(h);
    if (!b) return;
    b->AddRef();
    HWND hwnd = GetTridentHWND(h);
    
    NavigateTo(h, L"about:blank");
    
    READYSTATE st;
    while (SUCCEEDED(b->get_ReadyState(&st)) && st != READYSTATE_COMPLETE) {
        if (!IsWindow(hwnd)) { b->Release(); return; }
        MSG msg;
        while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        _mm_pause();
    }
    
    // Get h again after the message loop
    h = (hTrident)GetWindowLongPtr(hwnd, GWLP_USERDATA);
    if (!h || !IsWindow(hwnd)) { b->Release(); return; }
    if (h->site) h->site->SetupUIHandler();
    
    IDispatch* d = NULL;
    b->get_Document(&d);
    if (!d) { b->Release(); return; }
    
    IHTMLDocument2* doc = NULL;
    if (SUCCEEDED(d->QueryInterface(IID_IHTMLDocument2, (void**)&doc))) {
        SAFEARRAY* sa = SafeArrayCreateVector(VT_VARIANT, 0, 1);
        VARIANT* pv;
        SafeArrayAccessData(sa, (void**)&pv);
        pv->vt = VT_BSTR;
        pv->bstrVal = SysAllocString(html);
        SafeArrayUnaccessData(sa);
        doc->write(sa);
        doc->close();
        SafeArrayDestroy(sa);
        doc->Release();
    }
    d->Release();
    b->Release();
}

inline void ExecScript(hTrident h, const wchar_t* js) {
    IWebBrowser2* b = GetTridentBrowser(h);
    if (!b) return;
    
    IDispatch* d = NULL;
    b->get_Document(&d);
    if (!d) return;
    
    IHTMLDocument2* doc = NULL;
    if (SUCCEEDED(d->QueryInterface(IID_IHTMLDocument2, (void**)&doc))) {
        IHTMLWindow2* w = NULL;
        if (SUCCEEDED(doc->get_parentWindow(&w))) {
            VARIANT r;
            VariantInit(&r);
            BSTR code = SysAllocString(js);
            BSTR lang = SysAllocString(L"JavaScript");
            w->execScript(code, lang, &r);
            SysFreeString(code);
            SysFreeString(lang);
            VariantClear(&r);
            w->Release();
        }
        doc->Release();
    }
    d->Release();
}

inline hElement GetElement(hTrident h, const wchar_t* id) {
    IWebBrowser2* b = GetTridentBrowser(h);
    if (!b) return NULL;
    
    IDispatch* d = NULL;
    b->get_Document(&d);
    if (!d) return NULL;
    
    IHTMLElement* elem = NULL;
    IHTMLDocument3* doc3 = NULL;
    if (SUCCEEDED(d->QueryInterface(IID_IHTMLDocument3, (void**)&doc3))) {
        BSTR bstrId = SysAllocString(id);
        doc3->getElementById(bstrId, &elem);
        SysFreeString(bstrId);
        doc3->Release();
    }
    d->Release();
    return elem;
}

inline void ReleaseElement(hElement e) {
    if (e) e->Release();
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
    ReleaseElement(e);
}

inline std::wstring GetElementHTML(hTrident h, const wchar_t* id) {
    hElement e = GetElement(h, id);
    std::wstring result = GetHTML(e);
    ReleaseElement(e);
    return result;
}

inline void SetElementText(hTrident h, const wchar_t* id, const wchar_t* text) {
    hElement e = GetElement(h, id);
    SetText(e, text);
    ReleaseElement(e);
}

inline std::wstring GetElementText(hTrident h, const wchar_t* id) {
    hElement e = GetElement(h, id);
    std::wstring result = GetText(e);
    ReleaseElement(e);
    return result;
}

// ActiveX stuff
inline CLSID CLSIDFromName(const wchar_t* name) {
    CLSID c = {};
    uint32_t h1 = 0x811c9dc5, h2 = 0x01000193, h3 = 0xfedcba98, h4 = 0x76543210;
    for (const wchar_t* p = name; *p; p++) {
        h1 ^= (uint8_t)(*p & 0xFF); h1 *= 0x01000193;
        h2 ^= (uint8_t)(*p >> 8);   h2 *= 0x01000193;
        h3 ^= (uint8_t)(*p & 0xFF); h3 *= 0x100001b3;
        h4 ^= (uint8_t)(*p >> 8);   h4 *= 0x100001b3;
    }
    c.Data1 = h1;
    c.Data2 = (USHORT)(h2 & 0xFFFF);
    c.Data3 = (USHORT)(h3 & 0x0FFF) | 0x4000;
    memcpy(c.Data4, &h3, 4);
    memcpy(c.Data4 + 4, &h4, 4);
    c.Data4[0] = (c.Data4[0] & 0x3F) | 0x80;
    return c;
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

struct ControlReg_ {
    LONG refCount = 1;
    std::wstring name;
    ControlWndProc wndProc;
    void* defaultUserData;
    CLSID clsid;
    DWORD comCookie;
    std::wstring wndClassName;
    ATOM wndAtom = 0;
    DWORD extraStyle;
    CRITICAL_SECTION cs;
    DISPID nextDispId = INT32_MAX;
    std::map<std::wstring, DISPID> name2id;
    std::map<DISPID, DispEntry> entries;
    DISPID nextEventId = INT32_MAX;
    std::map<std::wstring, DISPID> eventName2Id;
    ControlReg_() { InitializeCriticalSection(&cs); }
    ~ControlReg_() {
        if (wndAtom) UnregisterClassW(wndClassName.c_str(), GetModuleHandle(NULL));
        DeleteCriticalSection(&cs);
    }
    void AddRef()  { InterlockedIncrement(&refCount); }
    void Release() { if (InterlockedDecrement(&refCount) == 0) delete this; }
};

inline std::map<CLSID, ControlReg_*, CLSIDCompare>& CtrlRegistry() {
    static std::map<CLSID, ControlReg_*, CLSIDCompare> s;
    return s;
}

struct ControlInstance_ {
    ControlReg_* reg = NULL;
    HWND hwnd = NULL;
    void* userData = NULL;
    IOleClientSite* clientSite = NULL;
    IOleAdviseHolder* adviseHolder = NULL;
    IOleInPlaceSite* inPlaceSite = NULL;
    bool active = false;
    bool uiActive = false;
    RECT rcPos;
    SIZEL extent = { 0, 0 };
    std::map<DWORD, IDispatch*> sinks;
    DWORD nextCookie = 1;
    DropCallback onDrop = NULL;
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
    LONG m_ref;
    ControlInstance_* m_inst;

    ActiveXControl(ControlReg_* reg) : m_ref(1) {
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
    STDMETHODIMP_(ULONG) AddRef()  override { return InterlockedIncrement(&m_ref); }
    STDMETHODIMP_(ULONG) Release() override { LONG r = InterlockedDecrement(&m_ref); if (!r) delete this; return r; }

    // IOleObject
    STDMETHODIMP SetClientSite(IOleClientSite* p) override { if (m_inst->clientSite) m_inst->clientSite->Release(); m_inst->clientSite = p; if (p) p->AddRef(); return S_OK; }
    STDMETHODIMP GetClientSite(IOleClientSite** pp) override { *pp = m_inst->clientSite; if (*pp) (*pp)->AddRef(); return S_OK; }
    STDMETHODIMP SetHostNames(LPCOLESTR, LPCOLESTR) override { return S_OK; }
    STDMETHODIMP Close(DWORD) override { if (m_inst->active) InPlaceDeactivate(); return S_OK; }
    STDMETHODIMP SetMoniker(DWORD, IMoniker*) override { return E_NOTIMPL; }
    STDMETHODIMP GetMoniker(DWORD, DWORD, IMoniker** p) override { if (!p) return E_POINTER; *p = NULL; return E_NOTIMPL; }
    STDMETHODIMP InitFromData(IDataObject*, BOOL, DWORD) override { return E_NOTIMPL; }
    STDMETHODIMP GetClipboardData(DWORD, IDataObject**) override { return E_NOTIMPL; }

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
                m_inst->active = true;
                if (pFrame) pFrame->Release();
                if (pUIWin) pUIWin->Release();
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
    STDMETHODIMP GetUserClassID(CLSID* p) override { *p = m_inst->reg->clsid; return S_OK; }
    STDMETHODIMP GetUserType(DWORD, LPOLESTR*) override { return E_NOTIMPL; }
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
        if (m_inst->extent.cx > 0 && m_inst->extent.cy > 0) {
            *p = m_inst->extent;
        } else {
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
        p->dwFlags = 0;
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
    STDMETHODIMP GetSizeMax(ULARGE_INTEGER* p) override { if (!p) return E_POINTER; p->QuadPart = 0; return S_OK; }
    STDMETHODIMP InitNew() override { return S_OK; }

    STDMETHODIMP GetInterfaceSafetyOptions(REFIID, DWORD* sup, DWORD* en) override {
        *sup = *en = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
        return S_OK;
    }
    STDMETHODIMP SetInterfaceSafetyOptions(REFIID, DWORD, DWORD) override { return S_OK; }

    STDMETHODIMP GetClassInfo(ITypeInfo**) override { return E_NOTIMPL; }
    STDMETHODIMP GetGUID(DWORD dwGuidKind, GUID* pGUID) override {
        if (!pGUID) return E_POINTER;
        if (dwGuidKind == GUIDKIND_DEFAULT_SOURCE_DISP_IID) { *pGUID = IID_IDispatch; return S_OK; }
        return E_INVALIDARG;
    }

    STDMETHODIMP GetTypeInfoCount(UINT* p) override { if (!p) return E_POINTER; *p = 0; return S_OK; }
    STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
    STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR* names, UINT cnt, LCID, DISPID* ids) override {
        HRESULT hr = S_OK;
        EnterCriticalSection(&m_inst->reg->cs);
        for (UINT i = 0; i < cnt; i++) {
            auto it = m_inst->reg->name2id.find(names[i]);
            if (it != m_inst->reg->name2id.end()) {
                ids[i] = it->second;
            } else {
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
            DispArgs args = { dp };
            if (res) VariantInit(res);
            return entry.method(m_inst, args, res);
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
        if (m_inst->onDragOver) return m_inst->onDragOver(m_inst, keys, pt, eff);
        *eff = DROPEFFECT_NONE; return S_OK;
    }
    STDMETHODIMP DragOver(DWORD keys, POINTL pt, DWORD* eff) override {
        if (!eff) return E_POINTER;
        if (m_inst->onDragOver) return m_inst->onDragOver(m_inst, keys, pt, eff);
        *eff = DROPEFFECT_NONE; return S_OK;
    }
    STDMETHODIMP DragLeave() override {
        if (m_inst->onDragLeave) m_inst->onDragLeave(m_inst);
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

inline LRESULT CALLBACK CtrlWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    ActiveXControl* ctrl = NULL;
    if (msg == WM_NCCREATE) {
        ctrl = (ActiveXControl*)((CREATESTRUCT*)lp)->lpCreateParams;
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)ctrl);
    }
    else {
        ctrl = (ActiveXControl*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
    }
    if (ctrl && ctrl->m_inst && ctrl->m_inst->reg && ctrl->m_inst->reg->wndProc) {
        LRESULT r = ctrl->m_inst->reg->wndProc(ctrl->m_inst, hwnd, msg, wp, lp);
        if (msg == WM_NCDESTROY) SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
        return r;
    }
    if (msg == WM_NCDESTROY) SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
    return DefWindowProc(hwnd, msg, wp, lp);
}

class CtrlConnectionPoint : public IConnectionPoint {
    LONG m_ref; ControlInstance_* m_inst;
    IConnectionPointContainer* m_pCPC;
public:
    CtrlConnectionPoint(ControlInstance_* inst, IConnectionPointContainer* pCPC)
        : m_ref(1), m_inst(inst), m_pCPC(pCPC) { if (m_pCPC) m_pCPC->AddRef(); }
    ~CtrlConnectionPoint() { if (m_pCPC) m_pCPC->Release(); }
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (riid == IID_IUnknown || riid == IID_IConnectionPoint) {
            *ppv = this; AddRef();
            return S_OK;
        }
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    STDMETHODIMP_(ULONG) AddRef()  override {
        return InterlockedIncrement(&m_ref);
    }
    STDMETHODIMP_(ULONG) Release() override {
        LONG r = InterlockedDecrement(&m_ref);
        if (!r) delete this;
        return r;
    }
    STDMETHODIMP GetConnectionInterface(IID* p) override {
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
        *pCookie = m_inst->nextCookie++;
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

class CtrlClassFactory : public IClassFactory {
    LONG m_ref; ControlReg_* m_reg;
public:
    CtrlClassFactory(ControlReg_* r) : m_ref(1), m_reg(r) {}
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (riid == IID_IUnknown || riid == IID_IClassFactory) {
            *ppv = this; AddRef();
            return S_OK;
        }
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    STDMETHODIMP_(ULONG) AddRef()  override { return InterlockedIncrement(&m_ref); }
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

class SimpleDataObject : public IDataObject {
    LONG m_ref; std::wstring m_text;
public:
    SimpleDataObject(const wchar_t* text) : m_ref(1), m_text(text ? text : L"") {}
    STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (riid == IID_IUnknown || riid == IID_IDataObject) {
            *ppv = this; AddRef();
            return S_OK;
        }
        *ppv = NULL;
        return E_NOINTERFACE;
    }
    STDMETHODIMP_(ULONG) AddRef()  override {
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
inline std::wstring CLSIDToString(const CLSID& clsid) { wchar_t buf[64]; StringFromGUID2(clsid, buf, 64); return buf; }
inline std::wstring GetControlClassId(hControlClass reg) { return reg ? L"clsid:" + CLSIDToString(reg->clsid) : L""; }
inline CLSID GetControlCLSID(hControlClass reg) { return reg ? reg->clsid : CLSID_NULL; }

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

inline DISPID RegisterEvent(hControlClass reg, const wchar_t* name) {
    if (!reg) return DISPID_UNKNOWN;
    EnterCriticalSection(&reg->cs);
    if (reg->eventName2Id.count(name)) {
        DISPID existing = reg->eventName2Id[name];
        LeaveCriticalSection(&reg->cs); return existing;
    }
    if (reg->nextEventId <= 0) {
        LeaveCriticalSection(&reg->cs);
        return DISPID_UNKNOWN;
    }
    DISPID id = reg->nextEventId--;
    reg->eventName2Id[name] = id;
    LeaveCriticalSection(&reg->cs);
    return id;
}

inline void FireEvent(hControl c, DISPID eventId, VARIANT* args, int argc) {
    if (!c || argc < 0) return;
    if (argc > 0 && !args) return;
    // Invert parameter order (COM needs this!)
    VARIANT* reversed = nullptr;
    reversed = new VARIANT[argc];
    for (int i = 0; i < argc; i++) reversed[i] = args[argc - 1 - i];
    DISPPARAMS dp = {};
    dp.rgvarg = reversed;
    dp.cArgs = argc;
    // Make a snapshot to ensure thread safty
    std::vector<IDispatch*> snapshot;
    snapshot.reserve(c->sinks.size());
    for (auto& kv : c->sinks) {
        if (kv.second) { kv.second->AddRef(); snapshot.push_back(kv.second); }
    }
    for (IDispatch* sink : snapshot) {
        sink->Invoke(eventId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, NULL, NULL, NULL);
        sink->Release();
    }
    delete[] reversed;
}

inline void FireEventByName(hControl c, const wchar_t* name, VARIANT* args, int argc) {
    if (!c || !c->reg) return;
    EnterCriticalSection(&c->reg->cs);
    auto it = c->reg->eventName2Id.find(name);
    if (it == c->reg->eventName2Id.end()) { LeaveCriticalSection(&c->reg->cs); return; }
    DISPID id = it->second;
    LeaveCriticalSection(&c->reg->cs);
    FireEvent(c, id, args, argc);
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

inline IDataObject* CreateTextDataObject(const wchar_t* text) { return new SimpleDataObject(text ? text : L""); }
