#pragma once

#include <atlbase.h>
#include <atlcom.h>
#include <oleidl.h>

#include <string>
#include <vector>

#include "FormulaOleContract.h"
#include "resource.h"

void TraceOleFactoryCall(const wchar_t* method) noexcept;

class CVisualTeXFormulaClassFactory : public ATL::CComClassFactory
{
public:
    STDMETHOD(CreateInstance)(LPUNKNOWN outer, REFIID interfaceId, void** object) override
    {
        TraceOleFactoryCall(L"IClassFactory::CreateInstance enter");
        const HRESULT result = __super::CreateInstance(outer, interfaceId, object);
        TraceOleFactoryCall(
            SUCCEEDED(result)
                ? L"IClassFactory::CreateInstance succeeded"
                : L"IClassFactory::CreateInstance failed");
        return result;
    }

    STDMETHOD(LockServer)(BOOL lock) override
    {
        TraceOleFactoryCall(lock
            ? L"IClassFactory::LockServer true"
            : L"IClassFactory::LockServer false");
        return __super::LockServer(lock);
    }
};

class ATL_NO_VTABLE CFormulaOleObject
    : public CComObjectRootEx<CComSingleThreadModel>,
      public CComCoClass<CFormulaOleObject, &CLSID_VisualTeXFormula>,
      public IOleObject,
      public IDataObject,
      public IPersistStorage,
      public IViewObject2,
      public ATL::IDispatchImpl<
          IVisualTeXFormulaObject,
          &__uuidof(IVisualTeXFormulaObject),
          &LIBID_VisualTeXFormulaOleLib,
          1,
          0>
{
public:
    CFormulaOleObject() noexcept;

    DECLARE_REGISTRY_RESOURCEID(IDR_FORMULAOLEOBJECT)
    DECLARE_CLASSFACTORY_EX(CVisualTeXFormulaClassFactory)
    DECLARE_NOT_AGGREGATABLE(CFormulaOleObject)
    DECLARE_PROTECT_FINAL_CONSTRUCT()

    BEGIN_COM_MAP(CFormulaOleObject)
        COM_INTERFACE_ENTRY(IOleObject)
        COM_INTERFACE_ENTRY(IDataObject)
        COM_INTERFACE_ENTRY(IPersistStorage)
        COM_INTERFACE_ENTRY2(IPersist, IPersistStorage)
        COM_INTERFACE_ENTRY(IViewObject2)
        COM_INTERFACE_ENTRY2(IViewObject, IViewObject2)
        COM_INTERFACE_ENTRY(IVisualTeXFormulaObject)
        COM_INTERFACE_ENTRY(IDispatch)
    END_COM_MAP()

    HRESULT FinalConstruct() noexcept;
    void FinalRelease() noexcept;

    // IOleObject
    HRESULT STDMETHODCALLTYPE SetClientSite(IOleClientSite* clientSite) override;
    HRESULT STDMETHODCALLTYPE GetClientSite(IOleClientSite** clientSite) override;
    HRESULT STDMETHODCALLTYPE SetHostNames(LPCOLESTR containerApp, LPCOLESTR containerObject) override;
    HRESULT STDMETHODCALLTYPE Close(DWORD saveOption) override;
    HRESULT STDMETHODCALLTYPE SetMoniker(DWORD whichMoniker, IMoniker* moniker) override;
    HRESULT STDMETHODCALLTYPE GetMoniker(DWORD assign, DWORD whichMoniker, IMoniker** moniker) override;
    HRESULT STDMETHODCALLTYPE InitFromData(IDataObject* dataObject, BOOL creation, DWORD reserved) override;
    HRESULT STDMETHODCALLTYPE GetClipboardData(DWORD reserved, IDataObject** dataObject) override;
    HRESULT STDMETHODCALLTYPE DoVerb(
        LONG verb,
        LPMSG message,
        IOleClientSite* activeSite,
        LONG index,
        HWND parentWindow,
        LPCRECT positionRectangle) override;
    HRESULT STDMETHODCALLTYPE EnumVerbs(IEnumOLEVERB** enumerator) override;
    HRESULT STDMETHODCALLTYPE Update() override;
    HRESULT STDMETHODCALLTYPE IsUpToDate() override;
    HRESULT STDMETHODCALLTYPE GetUserClassID(CLSID* classId) override;
    HRESULT STDMETHODCALLTYPE GetUserType(DWORD formOfType, LPOLESTR* userType) override;
    HRESULT STDMETHODCALLTYPE SetExtent(DWORD drawAspect, SIZEL* size) override;
    HRESULT STDMETHODCALLTYPE GetExtent(DWORD drawAspect, SIZEL* size) override;
    HRESULT STDMETHODCALLTYPE Advise(IAdviseSink* adviseSink, DWORD* connection) override;
    HRESULT STDMETHODCALLTYPE Unadvise(DWORD connection) override;
    HRESULT STDMETHODCALLTYPE EnumAdvise(IEnumSTATDATA** enumerator) override;
    HRESULT STDMETHODCALLTYPE GetMiscStatus(DWORD drawAspect, DWORD* status) override;
    HRESULT STDMETHODCALLTYPE SetColorScheme(LOGPALETTE* palette) override;

    // IDataObject
    HRESULT STDMETHODCALLTYPE GetData(FORMATETC* format, STGMEDIUM* medium) override;
    HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC* format, STGMEDIUM* medium) override;
    HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* format) override;
    HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC* input, FORMATETC* output) override;
    HRESULT STDMETHODCALLTYPE SetData(FORMATETC* format, STGMEDIUM* medium, BOOL release) override;
    HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD direction, IEnumFORMATETC** enumerator) override;
    HRESULT STDMETHODCALLTYPE DAdvise(
        FORMATETC* format,
        DWORD flags,
        IAdviseSink* adviseSink,
        DWORD* connection) override;
    HRESULT STDMETHODCALLTYPE DUnadvise(DWORD connection) override;
    HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA** enumerator) override;

    // IPersistStorage
    HRESULT STDMETHODCALLTYPE GetClassID(CLSID* classId) override;
    HRESULT STDMETHODCALLTYPE IsDirty() override;
    HRESULT STDMETHODCALLTYPE InitNew(IStorage* storage) override;
    HRESULT STDMETHODCALLTYPE Load(IStorage* storage) override;
    HRESULT STDMETHODCALLTYPE Save(IStorage* storage, BOOL sameAsLoad) override;
    HRESULT STDMETHODCALLTYPE SaveCompleted(IStorage* storage) override;
    HRESULT STDMETHODCALLTYPE HandsOffStorage() override;

    // IViewObject2
    HRESULT STDMETHODCALLTYPE Draw(
        DWORD drawAspect,
        LONG index,
        void* aspectInfo,
        DVTARGETDEVICE* targetDevice,
        HDC targetDeviceContext,
        HDC drawContext,
        LPCRECTL bounds,
        LPCRECTL windowBounds,
        BOOL(CALLBACK* continueDrawing)(ULONG_PTR),
        ULONG_PTR continueCookie) override;
    HRESULT STDMETHODCALLTYPE GetColorSet(
        DWORD drawAspect,
        LONG index,
        void* aspectInfo,
        DVTARGETDEVICE* targetDevice,
        HDC targetDeviceContext,
        LOGPALETTE** colorSet) override;
    HRESULT STDMETHODCALLTYPE Freeze(DWORD drawAspect, LONG index, void* aspectInfo, DWORD* freezeToken) override;
    HRESULT STDMETHODCALLTYPE Unfreeze(DWORD freezeToken) override;
    HRESULT STDMETHODCALLTYPE SetAdvise(DWORD aspects, DWORD flags, IAdviseSink* adviseSink) override;
    HRESULT STDMETHODCALLTYPE GetAdvise(DWORD* aspects, DWORD* flags, IAdviseSink** adviseSink) override;
    HRESULT STDMETHODCALLTYPE GetExtent(
        DWORD drawAspect,
        LONG index,
        DVTARGETDEVICE* targetDevice,
        LPSIZEL size) override;

    // IVisualTeXFormulaObject
    HRESULT STDMETHODCALLTYPE InitializeFromFiles(BSTR metadataJson, BSTR emfPath, BSTR pngPath) override;
    HRESULT STDMETHODCALLTYPE UpdateFromFiles(BSTR metadataJson, BSTR emfPath, BSTR pngPath) override;
    HRESULT STDMETHODCALLTYPE GetFormulaJson(BSTR* metadataJson) override;

private:
    HRESULT InitializeOrUpdate(BSTR metadataJson, BSTR emfPath, BSTR pngPath, bool requireUninitialized);
    HRESULT ReadOfficeTempFile(BSTR path, const wchar_t* expectedExtension, std::vector<BYTE>& bytes) const;
    HRESULT ReadStorageStream(IStorage* storage, const wchar_t* name, std::vector<BYTE>& bytes, bool required) const;
    HRESULT WriteStorageStream(IStorage* storage, const wchar_t* name, const std::vector<BYTE>& bytes) const;
    HRESULT LaunchVisualTeX() const;
    HRESULT CreatePlaceholderPreview() noexcept;
    HRESULT CopyBytesToGlobal(const std::vector<BYTE>& bytes, HGLOBAL* global) const;
    void UpdateExtentFromEmf(bool resetHostExtent = true) noexcept;
    void NotifyChanged() noexcept;
    bool DrawEmf(HDC drawContext, const RECT& bounds) const noexcept;
    bool DrawPng(HDC drawContext, const RECT& bounds) const noexcept;
    void DrawPlaceholder(HDC drawContext, const RECT& bounds) const noexcept;

    static HRESULT Utf8ToWide(const std::vector<BYTE>& bytes, std::wstring& value);
    static HRESULT WideToUtf8(const std::wstring& value, std::vector<BYTE>& bytes);

    CComPtr<IOleClientSite> clientSite_;
    CComPtr<IStorage> storage_;
    CComPtr<IOleAdviseHolder> oleAdviseHolder_;
    CComPtr<IDataAdviseHolder> dataAdviseHolder_;
    CComPtr<IAdviseSink> viewAdviseSink_;
    DWORD viewAspects_ = 0;
    DWORD viewFlags_ = 0;
    std::wstring metadataJson_;
    std::vector<BYTE> emfBytes_;
    std::vector<BYTE> pngBytes_;
    SIZEL extent_{};
    SIZEL naturalExtent_{};
    bool dirty_ = false;
    bool initialized_ = false;
};

OBJECT_ENTRY_AUTO(CLSID_VisualTeXFormula, CFormulaOleObject)
