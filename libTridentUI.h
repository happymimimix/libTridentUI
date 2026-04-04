/*
libTridentUI - v0.1.0 Alpha
An extremely lightweight Chrome Embedded Framework alternative for legacy operating systems.
Works perfectly on Microsoft Windows XP SP3!
Made possible by: Claude Opus 4.6 and Happy_mimimix
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

using MethodCallback = std::function<HRESULT(DispArgs&, VARIANT*)>;

typedef LRESULT (*TridentWndProc)(hTrident h, HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool* handled);

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

// Implementation
static const wchar_t* g_trident_wndclass = L"libTridentUI";
static TridentWindowData_* g_trident_head = NULL;

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

    STDMETHODIMP GetTypeInfoCount(UINT* p) override { *p = 0; return S_OK; }
    STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
    
    STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR* names, UINT cnt, LCID, DISPID* ids) override {
        EnterCriticalSection(&m_cs);
        for (UINT i = 0; i < cnt; i++) {
            auto it = m_name2id.find(names[i]);
            ids[i] = (it != m_name2id.end()) ? it->second : DISPID_UNKNOWN;
        }
        LeaveCriticalSection(&m_cs);
        return S_OK;
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

        HRESULT hr = CoCreateInstance(CLSID_WebBrowser, NULL, CLSCTX_ALL,
                                      IID_IOleObject, (void**)&m_pOle);
        if (FAILED(hr)) return false;

        m_pOle->SetClientSite(static_cast<IOleClientSite*>(this));

        IPersistStreamInit* psi = NULL;
        m_pOle->QueryInterface(IID_IPersistStreamInit, (void**)&psi);
        if (psi) { psi->InitNew(); psi->Release(); }
        hr = m_pOle->QueryInterface(IID_IViewObjectEx, (void**)&m_pView);
        if (FAILED(hr)) hr = m_pOle->QueryInterface(IID_IViewObject2, (void**)&m_pView);
        if (FAILED(hr)) m_pOle->QueryInterface(IID_IViewObject, (void**)&m_pView);
        m_pOle->SetHostNames(OLESTR("TridentUI"), NULL);
        RECT rc;
        GetClientRect(hwnd, &rc);
        m_pOle->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL,
                       static_cast<IOleClientSite*>(this), 0, hwnd, &rc);
        IObjectWithSite* pSite = NULL;
        m_pOle->QueryInterface(IID_IObjectWithSite, (void**)&pSite);
        if (pSite) { pSite->SetSite(static_cast<IOleClientSite*>(this)); pSite->Release(); }
        m_pOle->QueryInterface(IID_IWebBrowser2, (void**)&m_pBrowser);
        ConnectEvents(TRUE);
        
        return m_pBrowser != NULL;
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
    STDMETHODIMP_(ULONG) AddRef()  override { return ++m_ref; }
    STDMETHODIMP_(ULONG) Release() override { LONG r = --m_ref; if (!r) delete this; return r; }

    STDMETHODIMP SaveObject() override { return E_NOTIMPL; }
    STDMETHODIMP GetMoniker(DWORD, DWORD, IMoniker** p) override { *p = NULL; return E_NOTIMPL; }
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

    STDMETHODIMP GetWindow(HWND* p) override { *p = m_hwnd; return S_OK; }
    STDMETHODIMP ContextSensitiveHelp(BOOL) override { return S_OK; }
    STDMETHODIMP CanInPlaceActivate() override { return S_OK; }
    STDMETHODIMP OnInPlaceActivate() override {
        BOOL dummy = FALSE;
        return OnInPlaceActivateEx(&dummy, 0);
    }
    STDMETHODIMP OnUIActivate() override { m_bUIActivated = true; return S_OK; }
    STDMETHODIMP GetWindowContext(IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc, LPRECT pPos, LPRECT pClip, LPOLEINPLACEFRAMEINFO pFI) override {
        if (!ppFrame || !ppDoc || !pPos || !pClip) return E_POINTER;
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
        if (m_pInPlace) { m_pInPlace->Release(); m_pInPlace = NULL; }
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
        if (grfFlags & OLEDC_PAINTBKGND) {
            RECT rc; GetClientRect(m_hwnd, &rc);
            FillRect(*phDC, &rc, (HBRUSH)(COLOR_WINDOW + 1));
        }
        return S_OK;
    }
    STDMETHODIMP ReleaseDC(HDC h) override { ::ReleaseDC(m_hwnd, h); return S_OK; }
    STDMETHODIMP InvalidateRect(LPCRECT p, BOOL e) override { ::InvalidateRect(m_hwnd, p, e); return S_OK; }
    STDMETHODIMP InvalidateRgn(HRGN h, BOOL e) override { ::InvalidateRgn(m_hwnd, h, e); return S_OK; }
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
    STDMETHODIMP ShowContextMenu(DWORD, POINT*, IUnknown*, IDispatch*) override {
        return S_OK;  // S_OK = 屏蔽右键菜单
    }
    STDMETHODIMP GetHostInfo(DOCHOSTUIINFO* p) override {
        if (p) {
            p->cbSize = sizeof(DOCHOSTUIINFO);
            p->dwFlags = DOCHOSTUIFLAG_NO3DBORDER;
            p->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT;
        }
        return S_OK;
    }
    STDMETHODIMP ShowUI(DWORD, IOleInPlaceActiveObject*, IOleCommandTarget*,
                        IOleInPlaceFrame*, IOleInPlaceUIWindow*) override { return S_FALSE; }
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
    STDMETHODIMP GetTypeInfoCount(UINT* p) override { *p = 0; return S_OK; }
    STDMETHODIMP GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
    STDMETHODIMP GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override { return E_NOTIMPL; }
    STDMETHODIMP Invoke(DISPID id, REFIID riid, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) override {
        if (riid != IID_NULL) return E_INVALIDARG;
        if (id == DISPID_NAVIGATECOMPLETE2 || id == DISPID_DOCUMENTCOMPLETE)
            SetupUIHandler();
        return S_OK;
    }

private:
    void ConnectEvents(BOOL advise) {
        if (!m_pBrowser) return;
        IConnectionPointContainer* cpc = NULL;
        if (FAILED(m_pBrowser->QueryInterface(IID_IConnectionPointContainer, (void**)&cpc))) return;
        IConnectionPoint* cp = NULL;
        if (SUCCEEDED(cpc->FindConnectionPoint(DIID_DWebBrowserEvents2, &cp))) {
            if (advise) cp->Advise(static_cast<IDispatch*>(this), &m_cookie);
            else        cp->Unadvise(m_cookie);
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
                { *ppv = static_cast<IOleInPlaceFrame*>(this); AddRef(); return S_OK; }
            *ppv = NULL; return E_NOINTERFACE;
        }
        STDMETHODIMP_(ULONG) AddRef()  override { return ++m_r; }
        STDMETHODIMP_(ULONG) Release() override { if (--m_r == 0) { delete this; return 0; } return m_r; }
        STDMETHODIMP GetWindow(HWND* p) override { *p = m_h; return S_OK; }
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
        STDMETHODIMP_(ULONG) AddRef()  override { return ++m_r; }
        STDMETHODIMP_(ULONG) Release() override { if (--m_r == 0) { delete this; return 0; } return m_r; }
        STDMETHODIMP Next(ULONG, IUnknown** p, ULONG* f) override {                // [DUILIB-AX:88-96]
            if (f) *f = 0;
            if (++m_pos > 1) return S_FALSE;
            *p = m_p; (*p)->AddRef();
            if (f) *f = 1;
            return S_OK;
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
static LRESULT CALLBACK TridentHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
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
        if (msg == WM_SIZE && h->site) { h->site->Resize(LOWORD(lp), HIWORD(lp)); return 0; }
        if (msg == WM_CLOSE) { CloseTridentWindow(h); return 0; }
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
        for (TridentWindowData_* p = g_trident_head; p; p = p->next) {
            if (!p->alive || !p->site) continue;
            IWebBrowser2* b = p->site->GetBrowser();
            if (!b) continue;
            IOleInPlaceActiveObject* ao = NULL;
            if (SUCCEEDED(b->QueryInterface(IID_IOleInPlaceActiveObject, (void**)&ao))) {
                ate = (ao->TranslateAccelerator(&msg) == S_OK);
                ao->Release();
            }
            if (ate) break;
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
        delete d->site;
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
    h->alive = false;
    if (h->site) { h->site->Destroy(); delete h->site; h->site = NULL; }
    if (h->hwnd) { DestroyWindow(h->hwnd); h->hwnd = NULL; }
    // Remove from chain
    TridentWindowData_** pp = &g_trident_head;
    while (*pp) {
        if (*pp == h) { *pp = h->next; break; }
        pp = &(*pp)->next;
    }
    delete h;
    if (!g_trident_head) PostQuitMessage(0);
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
    wchar_t exe[MAX_PATH + 1] = {};
    GetModuleFileNameW(NULL, exe, MAX_PATH);
    std::wstring url = L"res://";
    url += exe;
    url += L"/";
    url += resName;
    NavigateTo(h, url.c_str());
}

inline void NavigateToHTML(hTrident h, const wchar_t* html) {
    IWebBrowser2* b = GetTridentBrowser(h);
    if (!b) return;
    
    NavigateTo(h, L"about:blank");
    
    READYSTATE st;
    while (SUCCEEDED(b->get_ReadyState(&st)) && st != READYSTATE_COMPLETE) {
        MSG msg;
        while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        Sleep(1);
    }
    
    if (h->site) h->site->SetupUIHandler();
    
    IDispatch* d = NULL;
    b->get_Document(&d);
    if (!d) return;
    
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
