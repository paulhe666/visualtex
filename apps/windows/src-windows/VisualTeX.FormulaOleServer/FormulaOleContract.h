#pragma once

#include <Unknwn.h>
#include <oaidl.h>

// Published ABI. These identities must remain stable after release.
inline constexpr wchar_t kVisualTeXFormulaProgId[] = L"VisualTeX.Formula.1";
inline constexpr wchar_t kVisualTeXFormulaVersionIndependentProgId[] = L"VisualTeX.Formula";
inline constexpr wchar_t kVisualTeXFormulaClassIdString[] = L"{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}";
inline constexpr wchar_t kVisualTeXFormulaInterfaceIdString[] = L"{6C672AF0-7321-4D21-B325-868CB34592C2}";
inline constexpr wchar_t kVisualTeXFormulaAppIdString[] = L"{3C72FF7F-B04A-4FD0-AA7D-61D110D8B3C1}";
inline constexpr wchar_t kVisualTeXFormulaTypeLibraryIdString[] = L"{DF66EC66-3B3A-4675-A7BE-30456A04EB96}";

inline constexpr wchar_t kVisualTeXMetadataStream[] = L"VisualTeX.Formula.json";
inline constexpr wchar_t kVisualTeXEmfPreviewStream[] = L"VisualTeX.Preview.emf";
inline constexpr wchar_t kVisualTeXPngPreviewStream[] = L"VisualTeX.Preview.png";

// {8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}
inline constexpr GUID LIBID_VisualTeXFormulaOleLib = {
    0xdf66ec66,
    0x3b3a,
    0x4675,
    {0xa7, 0xbe, 0x30, 0x45, 0x6a, 0x04, 0xeb, 0x96},
};

inline constexpr CLSID CLSID_VisualTeXFormula = {
    0x8ff7f5aa,
    0x0d60,
    0x48d5,
    {0xad, 0xbd, 0x65, 0xa6, 0x4b, 0x4c, 0x82, 0x7b},
};

MIDL_INTERFACE("6C672AF0-7321-4D21-B325-868CB34592C2")
IVisualTeXFormulaObject : public IDispatch
{
public:
    virtual HRESULT STDMETHODCALLTYPE InitializeFromFiles(
        BSTR metadataJson,
        BSTR emfPath,
        BSTR pngPath) = 0;

    virtual HRESULT STDMETHODCALLTYPE UpdateFromFiles(
        BSTR metadataJson,
        BSTR emfPath,
        BSTR pngPath) = 0;

    virtual HRESULT STDMETHODCALLTYPE GetFormulaJson(BSTR* metadataJson) = 0;
};
