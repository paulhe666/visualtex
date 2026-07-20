#include "FormulaOleObject.h"

#include <gdiplus.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shlwapi.h>

#include <algorithm>
#include <cwchar>
#include <filesystem>
#include <limits>

#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "shlwapi.lib")

namespace
{
constexpr ULONGLONG kMaximumPreviewBytes = 64ULL * 1024ULL * 1024ULL;
constexpr size_t kMaximumMetadataCharacters = 4ULL * 1024ULL * 1024ULL;
constexpr SIZEL kDefaultExtent = {2540, 635};
constexpr wchar_t kPlaceholderMetadataJson[] =
    L"{\"schemaVersion\":1,\"formulaId\":\"00000000-0000-0000-0000-000000000000\",\"placeholder\":true}";

void TraceOleCall(const wchar_t* method) noexcept
{
    std::wstring tracePath;
    wchar_t environmentPath[32768] = {};
    const DWORD length = GetEnvironmentVariableW(
        L"VISUALTEX_OLE_TRACE_PATH",
        environmentPath,
        static_cast<DWORD>(std::size(environmentPath)));
    if (length > 0 && length < std::size(environmentPath))
    {
        tracePath.assign(environmentPath, length);
    }
    else
    {
        PWSTR localApplicationData = nullptr;
        if (FAILED(SHGetKnownFolderPath(
                FOLDERID_LocalAppData,
                KF_FLAG_DEFAULT,
                nullptr,
                &localApplicationData)))
            return;
        std::wstring traceRoot(localApplicationData);
        CoTaskMemFree(localApplicationData);
        traceRoot.append(L"\\VisualTeX\\office");
        const std::wstring markerPath = traceRoot + L"\\ole-server-trace.enabled";
        if (GetFileAttributesW(markerPath.c_str()) == INVALID_FILE_ATTRIBUTES)
            return;
        tracePath = traceRoot + L"\\ole-server-trace.log";
    }

    SYSTEMTIME now = {};
    GetLocalTime(&now);
    wchar_t line[512] = {};
    const int characters = swprintf_s(
        line,
        L"%04u-%02u-%02u %02u:%02u:%02u.%03u pid=%lu tid=%lu %ls\r\n",
        now.wYear,
        now.wMonth,
        now.wDay,
        now.wHour,
        now.wMinute,
        now.wSecond,
        now.wMilliseconds,
        GetCurrentProcessId(),
        GetCurrentThreadId(),
        method);
    if (characters <= 0)
        return;

    const int byteCount = WideCharToMultiByte(
        CP_UTF8,
        0,
        line,
        characters,
        nullptr,
        0,
        nullptr,
        nullptr);
    if (byteCount <= 0)
        return;
    std::vector<char> bytes(static_cast<size_t>(byteCount));
    if (WideCharToMultiByte(
            CP_UTF8,
            0,
            line,
            characters,
            bytes.data(),
            byteCount,
            nullptr,
            nullptr) != byteCount)
        return;

    HANDLE file = CreateFileW(
        tracePath.c_str(),
        FILE_APPEND_DATA,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr,
        OPEN_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (file == INVALID_HANDLE_VALUE)
        return;
    DWORD written = 0;
    WriteFile(file, bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr);
    CloseHandle(file);
}

UINT PngClipboardFormat() noexcept
{
    static const UINT format = RegisterClipboardFormatW(L"PNG");
    return format;
}

HRESULT LastErrorResult() noexcept
{
    const DWORD error = GetLastError();
    return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_GEN_FAILURE : error);
}

HRESULT FinalPathFromHandle(HANDLE handle, std::wstring& path)
{
    const DWORD required = GetFinalPathNameByHandleW(handle, nullptr, 0, FILE_NAME_NORMALIZED);
    if (required == 0)
        return LastErrorResult();

    std::wstring buffer(required, L'\0');
    const DWORD written = GetFinalPathNameByHandleW(
        handle,
        buffer.data(),
        static_cast<DWORD>(buffer.size()),
        FILE_NAME_NORMALIZED);
    if (written == 0 || written >= buffer.size())
        return LastErrorResult();

    buffer.resize(written);
    path = std::move(buffer);
    return S_OK;
}

bool IsPathInsideRoot(const std::wstring& path, std::wstring root)
{
    if (root.empty())
        return false;
    if (root.back() != L'\\')
        root.push_back(L'\\');
    return path.size() > root.size() &&
           _wcsnicmp(path.c_str(), root.c_str(), root.size()) == 0;
}

bool HasPngSignature(const std::vector<BYTE>& bytes) noexcept
{
    static constexpr BYTE signature[] = {0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a};
    return bytes.size() >= std::size(signature) &&
           std::equal(std::begin(signature), std::end(signature), bytes.begin());
}

bool IsValidEmf(const std::vector<BYTE>& bytes) noexcept
{
    if (bytes.empty() || bytes.size() > std::numeric_limits<UINT>::max())
        return false;
    HENHMETAFILE metafile = SetEnhMetaFileBits(static_cast<UINT>(bytes.size()), bytes.data());
    if (metafile == nullptr)
        return false;
    DeleteEnhMetaFile(metafile);
    return true;
}

WORD ReadLittleEndianWord(const BYTE* bytes) noexcept
{
    return static_cast<WORD>(bytes[0] | (static_cast<WORD>(bytes[1]) << 8));
}

DWORD ReadLittleEndianDword(const BYTE* bytes) noexcept
{
    return static_cast<DWORD>(bytes[0]) |
           (static_cast<DWORD>(bytes[1]) << 8) |
           (static_cast<DWORD>(bytes[2]) << 16) |
           (static_cast<DWORD>(bytes[3]) << 24);
}

bool ValidateEmfPlusRecords(
    const BYTE* data,
    size_t dataSize,
    bool& sawVectorRecord) noexcept
{
    if (dataSize < 4 ||
        data[0] != static_cast<BYTE>('E') ||
        data[1] != static_cast<BYTE>('M') ||
        data[2] != static_cast<BYTE>('F') ||
        data[3] != static_cast<BYTE>('+'))
        return true;

    size_t offset = 4;
    while (offset + 12 <= dataSize)
    {
        const WORD type = ReadLittleEndianWord(data + offset);
        const WORD flags = ReadLittleEndianWord(data + offset + 2);
        const DWORD size = ReadLittleEndianDword(data + offset + 4);
        if (size < 12 || size > dataSize - offset)
            return false;

        if (type == 0x401a || type == 0x401b) // DrawImage / DrawImagePoints
            return false;
        if (type == 0x4008 && ((flags >> 8) & 0x7f) == 5) // ObjectTypeImage
            return false;
        switch (type)
        {
        case 0x400a: // FillRects
        case 0x400b: // DrawRects
        case 0x400c: // FillPolygon
        case 0x400d: // DrawLines
        case 0x400e: // FillEllipse
        case 0x400f: // DrawEllipse
        case 0x4010: // FillPie
        case 0x4011: // DrawPie
        case 0x4012: // DrawArc
        case 0x4013: // FillRegion
        case 0x4014: // FillPath
        case 0x4015: // DrawPath
        case 0x4016: // FillClosedCurve
        case 0x4017: // DrawClosedCurve
        case 0x4018: // DrawCurve
        case 0x4019: // DrawBeziers
        case 0x401c: // DrawString
            sawVectorRecord = true;
            break;
        default:
            break;
        }
        offset += size;
    }
    return offset == dataSize;
}

bool IsVectorEmf(const std::vector<BYTE>& bytes) noexcept
{
    if (!IsValidEmf(bytes))
    {
        TraceOleCall(L"IsVectorEmf SetEnhMetaFileBits failed");
        return false;
    }
    if (bytes.size() < sizeof(ENHMETAHEADER))
    {
        TraceOleCall(L"IsVectorEmf header truncated");
        return false;
    }

    bool sawVectorRecord = false;
    size_t offset = 0;
    while (offset + sizeof(ENHMETARECORD) <= bytes.size())
    {
        const auto* record = reinterpret_cast<const ENHMETARECORD*>(bytes.data() + offset);
        if (record->nSize < sizeof(DWORD) * 2 || record->nSize > bytes.size() - offset)
        {
            TraceOleCall(L"IsVectorEmf invalid record size");
            return false;
        }

        switch (record->iType)
        {
        case EMR_BITBLT:
        case EMR_STRETCHBLT:
        case EMR_MASKBLT:
        case EMR_PLGBLT:
        case EMR_SETDIBITSTODEVICE:
        case EMR_STRETCHDIBITS:
        case EMR_CREATEMONOBRUSH:
        case EMR_CREATEDIBPATTERNBRUSHPT:
        case EMR_ALPHABLEND:
        case EMR_TRANSPARENTBLT:
            return false;
        case EMR_POLYBEZIER:
        case EMR_POLYGON:
        case EMR_POLYLINE:
        case EMR_POLYBEZIERTO:
        case EMR_POLYLINETO:
        case EMR_POLYPOLYLINE:
        case EMR_POLYPOLYGON:
        case EMR_ELLIPSE:
        case EMR_RECTANGLE:
        case EMR_ROUNDRECT:
        case EMR_ARC:
        case EMR_CHORD:
        case EMR_PIE:
        case EMR_ANGLEARC:
        case EMR_LINETO:
        case EMR_ARCTO:
        case EMR_POLYDRAW:
        case EMR_FILLPATH:
        case EMR_STROKEANDFILLPATH:
        case EMR_STROKEPATH:
        case EMR_EXTTEXTOUTA:
        case EMR_EXTTEXTOUTW:
        case EMR_POLYBEZIER16:
        case EMR_POLYGON16:
        case EMR_POLYLINE16:
        case EMR_POLYBEZIERTO16:
        case EMR_POLYLINETO16:
        case EMR_POLYPOLYLINE16:
        case EMR_POLYPOLYGON16:
        case EMR_POLYDRAW16:
            sawVectorRecord = true;
            break;
        case EMR_GDICOMMENT:
        {
            if (record->nSize < 12)
                return false;
            const DWORD commentSize = ReadLittleEndianDword(bytes.data() + offset + 8);
            if (commentSize > record->nSize - 12)
                return false;
            if (!ValidateEmfPlusRecords(
                    bytes.data() + offset + 12,
                    static_cast<size_t>(commentSize),
                    sawVectorRecord))
                return false;
            break;
        }
        default:
            break;
        }

        offset += record->nSize;
        if (record->iType == EMR_EOF)
        {
            TraceOleCall(sawVectorRecord
                ? L"IsVectorEmf accepted"
                : L"IsVectorEmf no recognized vector record");
            return sawVectorRecord;
        }
    }
    TraceOleCall(L"IsVectorEmf missing EOF");
    return false;
}

HRESULT CopyCoTaskString(const wchar_t* value, LPOLESTR* output)
{
    if (output == nullptr)
        return E_POINTER;
    *output = nullptr;
    const size_t characters = std::wcslen(value) + 1;
    auto* copy = static_cast<wchar_t*>(CoTaskMemAlloc(characters * sizeof(wchar_t)));
    if (copy == nullptr)
        return E_OUTOFMEMORY;
    memcpy(copy, value, characters * sizeof(wchar_t));
    *output = copy;
    return S_OK;
}
} // namespace

void TraceOleFactoryCall(const wchar_t* method) noexcept
{
    TraceOleCall(method);
}

CFormulaOleObject::CFormulaOleObject() noexcept
    : extent_(kDefaultExtent), naturalExtent_(kDefaultExtent)
{
    TraceOleCall(L"CFormulaOleObject::ctor");
}

HRESULT CFormulaOleObject::FinalConstruct() noexcept
{
    TraceOleCall(L"FinalConstruct");
    return S_OK;
}

void CFormulaOleObject::FinalRelease() noexcept
{
    TraceOleCall(L"FinalRelease");
    viewAdviseSink_.Release();
    dataAdviseHolder_.Release();
    oleAdviseHolder_.Release();
    storage_.Release();
    clientSite_.Release();
}

HRESULT CFormulaOleObject::SetClientSite(IOleClientSite* clientSite)
{
    TraceOleCall(L"IOleObject::SetClientSite");
    clientSite_ = clientSite;
    return S_OK;
}

HRESULT CFormulaOleObject::GetClientSite(IOleClientSite** clientSite)
{
    TraceOleCall(L"IOleObject::GetClientSite");
    if (clientSite == nullptr)
        return E_POINTER;
    *clientSite = clientSite_;
    if (*clientSite != nullptr)
        (*clientSite)->AddRef();
    return S_OK;
}

HRESULT CFormulaOleObject::SetHostNames(LPCOLESTR, LPCOLESTR)
{
    TraceOleCall(L"IOleObject::SetHostNames");
    return S_OK;
}

HRESULT CFormulaOleObject::Close(DWORD saveOption)
{
    TraceOleCall(L"IOleObject::Close");
    if (dirty_ && clientSite_ != nullptr && saveOption != OLECLOSE_NOSAVE)
    {
        const HRESULT saveResult = clientSite_->SaveObject();
        if (FAILED(saveResult) && saveOption == OLECLOSE_SAVEIFDIRTY)
            return saveResult;
    }
    if (oleAdviseHolder_ != nullptr)
        oleAdviseHolder_->SendOnClose();
    clientSite_.Release();
    return S_OK;
}

HRESULT CFormulaOleObject::SetMoniker(DWORD, IMoniker*)
{
    return E_NOTIMPL;
}

HRESULT CFormulaOleObject::GetMoniker(DWORD, DWORD, IMoniker** moniker)
{
    if (moniker == nullptr)
        return E_POINTER;
    *moniker = nullptr;
    return E_NOTIMPL;
}

HRESULT CFormulaOleObject::InitFromData(IDataObject*, BOOL, DWORD)
{
    return E_NOTIMPL;
}

HRESULT CFormulaOleObject::GetClipboardData(DWORD, IDataObject** dataObject)
{
    if (dataObject == nullptr)
        return E_POINTER;
    *dataObject = nullptr;
    return E_NOTIMPL;
}

HRESULT CFormulaOleObject::DoVerb(
    LONG verb,
    LPMSG,
    IOleClientSite* activeSite,
    LONG,
    HWND,
    LPCRECT)
{
    TraceOleCall(L"IOleObject::DoVerb");
    if (activeSite != nullptr)
        clientSite_ = activeSite;

    switch (verb)
    {
    case OLEIVERB_PRIMARY:
    case OLEIVERB_OPEN:
        return LaunchVisualTeX();
    case OLEIVERB_SHOW:
    case OLEIVERB_HIDE:
        return S_OK;
    default:
        return OLEOBJ_S_INVALIDVERB;
    }
}

HRESULT CFormulaOleObject::EnumVerbs(IEnumOLEVERB** enumerator)
{
    if (enumerator == nullptr)
        return E_POINTER;
    *enumerator = nullptr;
    return OLE_S_USEREG;
}

HRESULT CFormulaOleObject::Update()
{
    return S_OK;
}

HRESULT CFormulaOleObject::IsUpToDate()
{
    return S_OK;
}

HRESULT CFormulaOleObject::GetUserClassID(CLSID* classId)
{
    return GetClassID(classId);
}

HRESULT CFormulaOleObject::GetUserType(DWORD, LPOLESTR* userType)
{
    return CopyCoTaskString(L"VisualTeX Formula", userType);
}

HRESULT CFormulaOleObject::SetExtent(DWORD drawAspect, SIZEL* size)
{
    TraceOleCall(L"IOleObject::SetExtent");
    if (size == nullptr)
        return E_POINTER;
    if (drawAspect != DVASPECT_CONTENT)
        return DV_E_DVASPECT;
    if (size->cx <= 0 || size->cy <= 0)
        return E_INVALIDARG;
    extent_ = *size;
    dirty_ = true;
    NotifyChanged();
    return S_OK;
}

HRESULT CFormulaOleObject::GetExtent(DWORD drawAspect, SIZEL* size)
{
    TraceOleCall(L"IOleObject::GetExtent");
    if (size == nullptr)
        return E_POINTER;
    if (drawAspect != DVASPECT_CONTENT)
        return DV_E_DVASPECT;
    *size = extent_;
    return S_OK;
}

HRESULT CFormulaOleObject::Advise(IAdviseSink* adviseSink, DWORD* connection)
{
    TraceOleCall(L"IOleObject::Advise");
    if (adviseSink == nullptr || connection == nullptr)
        return E_POINTER;
    if (oleAdviseHolder_ == nullptr)
    {
        const HRESULT result = CreateOleAdviseHolder(&oleAdviseHolder_);
        if (FAILED(result))
            return result;
    }
    return oleAdviseHolder_->Advise(adviseSink, connection);
}

HRESULT CFormulaOleObject::Unadvise(DWORD connection)
{
    return oleAdviseHolder_ == nullptr ? OLE_E_NOCONNECTION : oleAdviseHolder_->Unadvise(connection);
}

HRESULT CFormulaOleObject::EnumAdvise(IEnumSTATDATA** enumerator)
{
    if (enumerator == nullptr)
        return E_POINTER;
    *enumerator = nullptr;
    return oleAdviseHolder_ == nullptr ? S_OK : oleAdviseHolder_->EnumAdvise(enumerator);
}

HRESULT CFormulaOleObject::GetMiscStatus(DWORD drawAspect, DWORD* status)
{
    TraceOleCall(L"IOleObject::GetMiscStatus");
    if (status == nullptr)
        return E_POINTER;
    if (drawAspect != DVASPECT_CONTENT)
        return DV_E_DVASPECT;
    *status = OLEMISC_RECOMPOSEONRESIZE |
              OLEMISC_CANTLINKINSIDE |
              OLEMISC_NOUIACTIVATE |
              OLEMISC_SETCLIENTSITEFIRST;
    return S_OK;
}

HRESULT CFormulaOleObject::SetColorScheme(LOGPALETTE*)
{
    return S_OK;
}

HRESULT CFormulaOleObject::GetData(FORMATETC* format, STGMEDIUM* medium)
{
    TraceOleCall(L"IDataObject::GetData");
    if (medium == nullptr)
        return E_POINTER;
    ZeroMemory(medium, sizeof(*medium));
    const HRESULT query = QueryGetData(format);
    if (FAILED(query))
        return query;

    if (format->cfFormat == CF_ENHMETAFILE)
    {
        HENHMETAFILE metafile = SetEnhMetaFileBits(static_cast<UINT>(emfBytes_.size()), emfBytes_.data());
        if (metafile == nullptr)
            return DV_E_FORMATETC;
        medium->tymed = TYMED_ENHMF;
        medium->hEnhMetaFile = metafile;
        return S_OK;
    }

    if (format->cfFormat == CF_METAFILEPICT)
    {
        HENHMETAFILE enhanced = SetEnhMetaFileBits(
            static_cast<UINT>(emfBytes_.size()),
            emfBytes_.data());
        if (enhanced == nullptr)
            return DV_E_FORMATETC;
        HDC reference = GetDC(nullptr);
        const UINT byteCount = reference == nullptr
            ? 0
            : GetWinMetaFileBits(enhanced, 0, nullptr, MM_ANISOTROPIC, reference);
        std::vector<BYTE> metafileBytes(byteCount);
        const UINT copied = byteCount == 0 || reference == nullptr
            ? 0
            : GetWinMetaFileBits(
                enhanced,
                byteCount,
                metafileBytes.data(),
                MM_ANISOTROPIC,
                reference);
        if (reference != nullptr)
            ReleaseDC(nullptr, reference);
        DeleteEnhMetaFile(enhanced);
        if (copied != byteCount || byteCount == 0)
            return DV_E_FORMATETC;

        HMETAFILE metafile = SetMetaFileBitsEx(byteCount, metafileBytes.data());
        if (metafile == nullptr)
            return DV_E_FORMATETC;
        HGLOBAL metafilePicture = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, sizeof(METAFILEPICT));
        if (metafilePicture == nullptr)
        {
            DeleteMetaFile(metafile);
            return E_OUTOFMEMORY;
        }
        auto* picture = static_cast<METAFILEPICT*>(GlobalLock(metafilePicture));
        if (picture == nullptr)
        {
            GlobalFree(metafilePicture);
            DeleteMetaFile(metafile);
            return E_OUTOFMEMORY;
        }
        picture->mm = MM_ANISOTROPIC;
        // METAFILEPICT describes the preview's intrinsic physical size.
        // Using the mutable host extent here makes PowerPoint apply the host
        // resize twice when rebuilding its cache, which visibly squeezes or
        // flattens the formula after picture→OLE conversion.
        picture->xExt = naturalExtent_.cx;
        picture->yExt = naturalExtent_.cy;
        picture->hMF = metafile;
        GlobalUnlock(metafilePicture);
        medium->tymed = TYMED_MFPICT;
        medium->hMetaFilePict = metafilePicture;
        return S_OK;
    }

    HGLOBAL global = nullptr;
    const HRESULT copyResult = CopyBytesToGlobal(pngBytes_, &global);
    if (FAILED(copyResult))
        return copyResult;
    medium->tymed = TYMED_HGLOBAL;
    medium->hGlobal = global;
    return S_OK;
}

HRESULT CFormulaOleObject::GetDataHere(FORMATETC*, STGMEDIUM*)
{
    return DATA_E_FORMATETC;
}

HRESULT CFormulaOleObject::QueryGetData(FORMATETC* format)
{
    TraceOleCall(L"IDataObject::QueryGetData");
    if (format == nullptr)
        return E_POINTER;
    if (format->dwAspect != DVASPECT_CONTENT)
        return DV_E_DVASPECT;
    if (format->lindex != -1)
        return DV_E_LINDEX;
    if (format->cfFormat == CF_ENHMETAFILE &&
        (format->tymed & TYMED_ENHMF) != 0 &&
        !emfBytes_.empty())
        return S_OK;
    if (format->cfFormat == CF_METAFILEPICT &&
        (format->tymed & TYMED_MFPICT) != 0 &&
        !emfBytes_.empty())
        return S_OK;
    if (format->cfFormat == static_cast<CLIPFORMAT>(PngClipboardFormat()) &&
        (format->tymed & TYMED_HGLOBAL) != 0 &&
        !pngBytes_.empty())
        return S_OK;
    return DV_E_FORMATETC;
}

HRESULT CFormulaOleObject::GetCanonicalFormatEtc(FORMATETC* input, FORMATETC* output)
{
    if (input == nullptr || output == nullptr)
        return E_POINTER;
    *output = *input;
    output->ptd = nullptr;
    return DATA_S_SAMEFORMATETC;
}

HRESULT CFormulaOleObject::SetData(FORMATETC*, STGMEDIUM*, BOOL)
{
    return E_NOTIMPL;
}

HRESULT CFormulaOleObject::EnumFormatEtc(DWORD direction, IEnumFORMATETC** enumerator)
{
    if (enumerator == nullptr)
        return E_POINTER;
    *enumerator = nullptr;
    if (direction != DATADIR_GET)
        return E_NOTIMPL;

    FORMATETC formats[3] = {};
    formats[0].cfFormat = CF_ENHMETAFILE;
    formats[0].dwAspect = DVASPECT_CONTENT;
    formats[0].lindex = -1;
    formats[0].tymed = TYMED_ENHMF;
    formats[1].cfFormat = CF_METAFILEPICT;
    formats[1].dwAspect = DVASPECT_CONTENT;
    formats[1].lindex = -1;
    formats[1].tymed = TYMED_MFPICT;
    formats[2].cfFormat = static_cast<CLIPFORMAT>(PngClipboardFormat());
    formats[2].dwAspect = DVASPECT_CONTENT;
    formats[2].lindex = -1;
    formats[2].tymed = TYMED_HGLOBAL;
    return SHCreateStdEnumFmtEtc(static_cast<UINT>(std::size(formats)), formats, enumerator);
}

HRESULT CFormulaOleObject::DAdvise(
    FORMATETC* format,
    DWORD flags,
    IAdviseSink* adviseSink,
    DWORD* connection)
{
    TraceOleCall(L"IDataObject::DAdvise");
    if (format == nullptr || adviseSink == nullptr || connection == nullptr)
        return E_POINTER;
    if (dataAdviseHolder_ == nullptr)
    {
        const HRESULT result = CreateDataAdviseHolder(&dataAdviseHolder_);
        if (FAILED(result))
            return result;
    }
    return dataAdviseHolder_->Advise(this, format, flags, adviseSink, connection);
}

HRESULT CFormulaOleObject::DUnadvise(DWORD connection)
{
    return dataAdviseHolder_ == nullptr ? OLE_E_NOCONNECTION : dataAdviseHolder_->Unadvise(connection);
}

HRESULT CFormulaOleObject::EnumDAdvise(IEnumSTATDATA** enumerator)
{
    if (enumerator == nullptr)
        return E_POINTER;
    *enumerator = nullptr;
    return dataAdviseHolder_ == nullptr ? OLE_E_ADVISENOTSUPPORTED : dataAdviseHolder_->EnumAdvise(enumerator);
}

HRESULT CFormulaOleObject::GetClassID(CLSID* classId)
{
    if (classId == nullptr)
        return E_POINTER;
    *classId = CLSID_VisualTeXFormula;
    return S_OK;
}

HRESULT CFormulaOleObject::IsDirty()
{
    return dirty_ ? S_OK : S_FALSE;
}

HRESULT CFormulaOleObject::InitNew(IStorage* storage)
{
    TraceOleCall(L"IPersistStorage::InitNew");
    if (storage == nullptr)
        return E_POINTER;
    storage_ = storage;
    metadataJson_.clear();
    emfBytes_.clear();
    pngBytes_.clear();
    extent_ = kDefaultExtent;
    naturalExtent_ = kDefaultExtent;
    initialized_ = false;
    dirty_ = false;
    HRESULT result = storage->SetClass(CLSID_VisualTeXFormula);
    if (FAILED(result))
        return result;
    result = CreatePlaceholderPreview();
    if (FAILED(result))
        storage_.Release();
    return result;
}

HRESULT CFormulaOleObject::Load(IStorage* storage)
{
    TraceOleCall(L"IPersistStorage::Load");
    if (storage == nullptr)
        return E_POINTER;

    std::vector<BYTE> metadataBytes;
    std::vector<BYTE> emfBytes;
    std::vector<BYTE> pngBytes;
    std::wstring metadataJson;

    HRESULT result = ReadStorageStream(storage, kVisualTeXMetadataStream, metadataBytes, true);
    if (FAILED(result))
        return result;
    result = ReadStorageStream(storage, kVisualTeXEmfPreviewStream, emfBytes, true);
    if (FAILED(result))
        return result;
    result = ReadStorageStream(storage, kVisualTeXPngPreviewStream, pngBytes, true);
    if (FAILED(result))
        return result;
    result = Utf8ToWide(metadataBytes, metadataJson);
    if (FAILED(result) || metadataJson.empty() ||
        metadataJson.find(L"\"schemaVersion\"") == std::wstring::npos ||
        metadataJson.find(L"\"formulaId\"") == std::wstring::npos ||
        !IsVectorEmf(emfBytes) || !HasPngSignature(pngBytes))
        return STG_E_INVALIDHEADER;

    const bool placeholder = metadataJson == kPlaceholderMetadataJson;
    metadataJson_ = placeholder ? std::wstring() : std::move(metadataJson);
    emfBytes_ = std::move(emfBytes);
    pngBytes_ = std::move(pngBytes);
    storage_ = storage;
    initialized_ = !placeholder;
    dirty_ = false;
    UpdateExtentFromEmf(true);
    return S_OK;
}

HRESULT CFormulaOleObject::Save(IStorage* storage, BOOL sameAsLoad)
{
    TraceOleCall(L"IPersistStorage::Save");
    if (storage == nullptr)
        return E_POINTER;
    if (!initialized_ && (!IsVectorEmf(emfBytes_) || !HasPngSignature(pngBytes_)))
        return OLE_E_BLANK;

    std::vector<BYTE> metadataBytes;
    const std::wstring metadataToPersist = initialized_
        ? metadataJson_
        : std::wstring(kPlaceholderMetadataJson);
    HRESULT result = WideToUtf8(metadataToPersist, metadataBytes);
    if (FAILED(result))
        return result;
    result = storage->SetClass(CLSID_VisualTeXFormula);
    if (FAILED(result))
        return result;
    result = WriteStorageStream(storage, kVisualTeXMetadataStream, metadataBytes);
    if (FAILED(result))
        return result;
    result = WriteStorageStream(storage, kVisualTeXEmfPreviewStream, emfBytes_);
    if (FAILED(result))
        return result;
    result = WriteStorageStream(storage, kVisualTeXPngPreviewStream, pngBytes_);
    if (FAILED(result))
        return result;
    result = storage->Commit(STGC_DEFAULT);
    if (FAILED(result))
        return result;
    if (sameAsLoad)
        storage_ = storage;
    return S_OK;
}

HRESULT CFormulaOleObject::SaveCompleted(IStorage* storage)
{
    TraceOleCall(L"IPersistStorage::SaveCompleted");
    if (storage != nullptr)
        storage_ = storage;
    dirty_ = false;
    if (oleAdviseHolder_ != nullptr)
        oleAdviseHolder_->SendOnSave();
    return S_OK;
}

HRESULT CFormulaOleObject::HandsOffStorage()
{
    storage_.Release();
    return S_OK;
}

HRESULT CFormulaOleObject::Draw(
    DWORD drawAspect,
    LONG index,
    void*,
    DVTARGETDEVICE*,
    HDC,
    HDC drawContext,
    LPCRECTL bounds,
    LPCRECTL,
    BOOL(CALLBACK* continueDrawing)(ULONG_PTR),
    ULONG_PTR continueCookie)
{
    if (drawAspect != DVASPECT_CONTENT)
        return DV_E_DVASPECT;
    if (index != -1)
        return DV_E_LINDEX;
    if (drawContext == nullptr || bounds == nullptr)
        return E_INVALIDARG;
    if (continueDrawing != nullptr && !continueDrawing(continueCookie))
        return E_ABORT;

    RECT rectangle = {
        static_cast<LONG>(bounds->left),
        static_cast<LONG>(bounds->top),
        static_cast<LONG>(bounds->right),
        static_cast<LONG>(bounds->bottom),
    };
    if (!DrawEmf(drawContext, rectangle) && !DrawPng(drawContext, rectangle))
        DrawPlaceholder(drawContext, rectangle);

    if (continueDrawing != nullptr && !continueDrawing(continueCookie))
        return E_ABORT;
    return S_OK;
}

HRESULT CFormulaOleObject::GetColorSet(
    DWORD,
    LONG,
    void*,
    DVTARGETDEVICE*,
    HDC,
    LOGPALETTE** colorSet)
{
    if (colorSet == nullptr)
        return E_POINTER;
    *colorSet = nullptr;
    return E_NOTIMPL;
}

HRESULT CFormulaOleObject::Freeze(DWORD, LONG, void*, DWORD* freezeToken)
{
    if (freezeToken == nullptr)
        return E_POINTER;
    *freezeToken = 0;
    return E_NOTIMPL;
}

HRESULT CFormulaOleObject::Unfreeze(DWORD)
{
    return E_NOTIMPL;
}

HRESULT CFormulaOleObject::SetAdvise(DWORD aspects, DWORD flags, IAdviseSink* adviseSink)
{
    TraceOleCall(L"IViewObject2::SetAdvise");
    viewAspects_ = aspects;
    viewFlags_ = flags;
    viewAdviseSink_ = adviseSink;
    if ((flags & ADVF_PRIMEFIRST) != 0 && viewAdviseSink_ != nullptr)
        viewAdviseSink_->OnViewChange(DVASPECT_CONTENT, -1);
    return S_OK;
}

HRESULT CFormulaOleObject::GetAdvise(DWORD* aspects, DWORD* flags, IAdviseSink** adviseSink)
{
    if (aspects != nullptr)
        *aspects = viewAspects_;
    if (flags != nullptr)
        *flags = viewFlags_;
    if (adviseSink != nullptr)
    {
        *adviseSink = viewAdviseSink_;
        if (*adviseSink != nullptr)
            (*adviseSink)->AddRef();
    }
    return S_OK;
}

HRESULT CFormulaOleObject::GetExtent(DWORD drawAspect, LONG index, DVTARGETDEVICE*, LPSIZEL size)
{
    if (size == nullptr)
        return E_POINTER;
    if (drawAspect != DVASPECT_CONTENT)
        return DV_E_DVASPECT;
    if (index != -1)
        return DV_E_LINDEX;
    *size = extent_;
    return S_OK;
}

HRESULT CFormulaOleObject::InitializeFromFiles(BSTR metadataJson, BSTR emfPath, BSTR pngPath)
{
    return InitializeOrUpdate(metadataJson, emfPath, pngPath, true);
}

HRESULT CFormulaOleObject::UpdateFromFiles(BSTR metadataJson, BSTR emfPath, BSTR pngPath)
{
    return InitializeOrUpdate(metadataJson, emfPath, pngPath, false);
}

HRESULT CFormulaOleObject::GetFormulaJson(BSTR* metadataJson)
{
    if (metadataJson == nullptr)
        return E_POINTER;
    if (!initialized_)
    {
        *metadataJson = nullptr;
        return CO_E_NOTINITIALIZED;
    }
    *metadataJson = SysAllocStringLen(metadataJson_.data(), static_cast<UINT>(metadataJson_.size()));
    return *metadataJson == nullptr && !metadataJson_.empty() ? E_OUTOFMEMORY : S_OK;
}

HRESULT CFormulaOleObject::InitializeOrUpdate(
    BSTR metadataJson,
    BSTR emfPath,
    BSTR pngPath,
    bool requireUninitialized)
{
    TraceOleCall(requireUninitialized
        ? L"InitializeOrUpdate initialize enter"
        : L"InitializeOrUpdate update enter");
    if (metadataJson == nullptr || emfPath == nullptr || pngPath == nullptr)
    {
        TraceOleCall(L"InitializeOrUpdate null argument");
        return E_INVALIDARG;
    }
    if (requireUninitialized && initialized_)
        return HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);
    if (!requireUninitialized && !initialized_)
        return CO_E_NOTINITIALIZED;

    const UINT metadataLength = SysStringLen(metadataJson);
    if (metadataLength == 0 || metadataLength > kMaximumMetadataCharacters)
    {
        TraceOleCall(L"InitializeOrUpdate metadata length invalid");
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    std::wstring nextMetadata(metadataJson, metadataLength);
    if (nextMetadata.find(L"\"schemaVersion\"") == std::wstring::npos ||
        nextMetadata.find(L"\"formulaId\"") == std::wstring::npos)
    {
        TraceOleCall(L"InitializeOrUpdate metadata fields invalid");
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    std::vector<BYTE> nextEmf;
    std::vector<BYTE> nextPng;
    HRESULT result = ReadOfficeTempFile(emfPath, L".emf", nextEmf);
    if (FAILED(result))
    {
        TraceOleCall(L"InitializeOrUpdate EMF read failed");
        return result;
    }
    result = ReadOfficeTempFile(pngPath, L".png", nextPng);
    if (FAILED(result))
    {
        TraceOleCall(L"InitializeOrUpdate PNG read failed");
        return result;
    }
    if (!IsVectorEmf(nextEmf))
    {
        TraceOleCall(L"InitializeOrUpdate EMF validation failed");
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (!HasPngSignature(nextPng))
    {
        TraceOleCall(L"InitializeOrUpdate PNG validation failed");
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    std::wstring previousMetadata = std::move(metadataJson_);
    std::vector<BYTE> previousEmf = std::move(emfBytes_);
    std::vector<BYTE> previousPng = std::move(pngBytes_);
    const SIZEL previousExtent = extent_;
    const SIZEL previousNaturalExtent = naturalExtent_;
    const bool previousInitialized = initialized_;
    const bool previousDirty = dirty_;

    metadataJson_ = std::move(nextMetadata);
    emfBytes_ = std::move(nextEmf);
    pngBytes_ = std::move(nextPng);
    initialized_ = true;
    dirty_ = true;
    // A new object starts at the preview's natural size. Updating an existing
    // object must preserve the host extent; PowerPoint/Word own the outer box.
    UpdateExtentFromEmf(requireUninitialized);
    NotifyChanged();

    if (clientSite_ != nullptr)
    {
        TraceOleCall(L"InitializeOrUpdate requesting container save");
        const HRESULT saveResult = clientSite_->SaveObject();
        if (FAILED(saveResult))
        {
            TraceOleCall(L"InitializeOrUpdate container save failed");
            metadataJson_ = std::move(previousMetadata);
            emfBytes_ = std::move(previousEmf);
            pngBytes_ = std::move(previousPng);
            extent_ = previousExtent;
            naturalExtent_ = previousNaturalExtent;
            initialized_ = previousInitialized;
            dirty_ = previousDirty;
            NotifyChanged();
            return saveResult;
        }
        TraceOleCall(L"InitializeOrUpdate container save succeeded");
        // PowerPoint can establish or refresh its presentation cache while
        // SaveObject is running. Send a second view/data notification after
        // the durable save so the host does not retain the initial placeholder.
        NotifyChanged();
    }

    TraceOleCall(L"InitializeOrUpdate succeeded");
    return S_OK;
}

HRESULT CFormulaOleObject::ReadOfficeTempFile(
    BSTR path,
    const wchar_t* expectedExtension,
    std::vector<BYTE>& bytes) const
{
    if (path == nullptr || SysStringLen(path) == 0)
        return E_INVALIDARG;

    PWSTR localApplicationData = nullptr;
    HRESULT result = SHGetKnownFolderPath(FOLDERID_LocalAppData, KF_FLAG_DEFAULT, nullptr, &localApplicationData);
    if (FAILED(result))
        return result;

    std::wstring root(localApplicationData);
    CoTaskMemFree(localApplicationData);
    root.append(L"\\VisualTeX\\office\\temp");

    HANDLE rootHandle = CreateFileW(
        root.c_str(),
        FILE_READ_ATTRIBUTES,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr,
        OPEN_EXISTING,
        FILE_FLAG_BACKUP_SEMANTICS,
        nullptr);
    if (rootHandle == INVALID_HANDLE_VALUE)
        return LastErrorResult();

    HANDLE fileHandle = CreateFileW(
        path,
        GENERIC_READ,
        FILE_SHARE_READ,
        nullptr,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN,
        nullptr);
    if (fileHandle == INVALID_HANDLE_VALUE)
    {
        CloseHandle(rootHandle);
        return LastErrorResult();
    }

    std::wstring finalRoot;
    std::wstring finalFile;
    result = FinalPathFromHandle(rootHandle, finalRoot);
    if (SUCCEEDED(result))
        result = FinalPathFromHandle(fileHandle, finalFile);
    CloseHandle(rootHandle);
    if (FAILED(result))
    {
        CloseHandle(fileHandle);
        return result;
    }
    if (!IsPathInsideRoot(finalFile, finalRoot) ||
        _wcsicmp(PathFindExtensionW(finalFile.c_str()), expectedExtension) != 0)
    {
        CloseHandle(fileHandle);
        return E_ACCESSDENIED;
    }

    BY_HANDLE_FILE_INFORMATION information = {};
    if (!GetFileInformationByHandle(fileHandle, &information))
    {
        CloseHandle(fileHandle);
        return LastErrorResult();
    }
    if ((information.dwFileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_REPARSE_POINT)) != 0)
    {
        CloseHandle(fileHandle);
        return E_ACCESSDENIED;
    }

    LARGE_INTEGER size = {};
    if (!GetFileSizeEx(fileHandle, &size))
    {
        CloseHandle(fileHandle);
        return LastErrorResult();
    }
    if (size.QuadPart <= 0 || static_cast<ULONGLONG>(size.QuadPart) > kMaximumPreviewBytes)
    {
        CloseHandle(fileHandle);
        return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
    }

    std::vector<BYTE> loaded(static_cast<size_t>(size.QuadPart));
    size_t offset = 0;
    while (offset < loaded.size())
    {
        const DWORD requested = static_cast<DWORD>(
            std::min<size_t>(loaded.size() - offset, std::numeric_limits<DWORD>::max()));
        DWORD read = 0;
        if (!ReadFile(fileHandle, loaded.data() + offset, requested, &read, nullptr))
        {
            CloseHandle(fileHandle);
            return LastErrorResult();
        }
        if (read == 0)
        {
            CloseHandle(fileHandle);
            return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
        }
        offset += read;
    }
    CloseHandle(fileHandle);
    bytes = std::move(loaded);
    return S_OK;
}

HRESULT CFormulaOleObject::ReadStorageStream(
    IStorage* storage,
    const wchar_t* name,
    std::vector<BYTE>& bytes,
    bool required) const
{
    CComPtr<IStream> stream;
    HRESULT result = storage->OpenStream(name, nullptr, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &stream);
    if (result == STG_E_FILENOTFOUND && !required)
    {
        bytes.clear();
        return S_OK;
    }
    if (FAILED(result))
        return result;

    STATSTG stat = {};
    result = stream->Stat(&stat, STATFLAG_NONAME);
    if (FAILED(result))
        return result;
    if (stat.cbSize.QuadPart < 0 || static_cast<ULONGLONG>(stat.cbSize.QuadPart) > kMaximumPreviewBytes)
        return STG_E_INVALIDHEADER;

    std::vector<BYTE> loaded(static_cast<size_t>(stat.cbSize.QuadPart));
    ULONG read = 0;
    if (!loaded.empty())
    {
        result = stream->Read(loaded.data(), static_cast<ULONG>(loaded.size()), &read);
        if (FAILED(result) || read != loaded.size())
            return FAILED(result) ? result : STG_E_READFAULT;
    }
    bytes = std::move(loaded);
    return S_OK;
}

HRESULT CFormulaOleObject::WriteStorageStream(
    IStorage* storage,
    const wchar_t* name,
    const std::vector<BYTE>& bytes) const
{
    CComPtr<IStream> stream;
    HRESULT result = storage->CreateStream(
        name,
        STGM_CREATE | STGM_WRITE | STGM_SHARE_EXCLUSIVE,
        0,
        0,
        &stream);
    if (FAILED(result))
        return result;
    if (!bytes.empty())
    {
        ULONG written = 0;
        result = stream->Write(bytes.data(), static_cast<ULONG>(bytes.size()), &written);
        if (FAILED(result) || written != bytes.size())
            return FAILED(result) ? result : STG_E_WRITEFAULT;
    }
    return stream->Commit(STGC_DEFAULT);
}

HRESULT CFormulaOleObject::LaunchVisualTeX() const
{
    HINSTANCE launched = ShellExecuteW(
        nullptr,
        L"open",
        L"visualtex://office/ole",
        nullptr,
        nullptr,
        SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(launched) > 32)
        return S_OK;

    wchar_t modulePath[MAX_PATH] = {};
    const DWORD moduleLength = GetModuleFileNameW(nullptr, modulePath, ARRAYSIZE(modulePath));
    if (moduleLength > 0 && moduleLength < ARRAYSIZE(modulePath))
    {
        std::filesystem::path executable(modulePath);
        executable = executable.parent_path().parent_path().parent_path() / L"visualtex.exe";
        launched = ShellExecuteW(
            nullptr,
            L"open",
            executable.c_str(),
            nullptr,
            nullptr,
            SW_SHOWNORMAL);
        if (reinterpret_cast<INT_PTR>(launched) > 32)
            return S_OK;
    }

    PWSTR localApplicationData = nullptr;
    HRESULT result = SHGetKnownFolderPath(
        FOLDERID_LocalAppData,
        KF_FLAG_DEFAULT,
        nullptr,
        &localApplicationData);
    if (FAILED(result))
        return result;
    std::filesystem::path localRoot(localApplicationData);
    CoTaskMemFree(localApplicationData);

    const std::filesystem::path candidates[] = {
        localRoot / L"VisualTeX" / L"visualtex.exe",
        localRoot / L"Programs" / L"VisualTeX" / L"VisualTeX.exe",
    };
    for (const auto& executable : candidates)
    {
        launched = ShellExecuteW(
            nullptr,
            L"open",
            executable.c_str(),
            nullptr,
            nullptr,
            SW_SHOWNORMAL);
        if (reinterpret_cast<INT_PTR>(launched) > 32)
            return S_OK;
    }

    return HRESULT_FROM_WIN32(
        static_cast<DWORD>(reinterpret_cast<INT_PTR>(launched)));
}

HRESULT CFormulaOleObject::CreatePlaceholderPreview() noexcept
{
    static constexpr BYTE kTransparentPng[] = {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x04, 0x00, 0x00, 0x00, 0xB5, 0x1C, 0x0C,
        0x02, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41,
        0x54, 0x78, 0xDA, 0x63, 0xFC, 0xFF, 0x1F, 0x00,
        0x02, 0xEB, 0x01, 0xF5, 0x8F, 0x59, 0x42, 0x67,
        0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44,
        0xAE, 0x42, 0x60, 0x82,
    };

    HDC reference = GetDC(nullptr);
    if (reference == nullptr)
        return LastErrorResult();

    RECT frame = {0, 0, kDefaultExtent.cx, kDefaultExtent.cy};
    HDC metafileDc = CreateEnhMetaFileW(
        reference,
        nullptr,
        &frame,
        L"VisualTeX\0Formula placeholder\0\0");
    ReleaseDC(nullptr, reference);
    if (metafileDc == nullptr)
        return LastErrorResult();

    HGDIOBJ previousBrush = SelectObject(metafileDc, GetStockObject(WHITE_BRUSH));
    HPEN borderPen = CreatePen(PS_SOLID, 2, RGB(128, 128, 128));
    HGDIOBJ previousPen = borderPen != nullptr
        ? SelectObject(metafileDc, borderPen)
        : nullptr;
    Rectangle(metafileDc, 1, 1, 399, 99);

    HPEN formulaPen = CreatePen(PS_SOLID, 4, RGB(70, 70, 70));
    if (formulaPen != nullptr)
        SelectObject(metafileDc, formulaPen);
    MoveToEx(metafileDc, 70, 72, nullptr);
    LineTo(metafileDc, 105, 28);
    LineTo(metafileDc, 140, 72);
    MoveToEx(metafileDc, 165, 50, nullptr);
    LineTo(metafileDc, 235, 50);
    MoveToEx(metafileDc, 270, 30, nullptr);
    LineTo(metafileDc, 330, 30);
    MoveToEx(metafileDc, 270, 70, nullptr);
    LineTo(metafileDc, 330, 70);

    if (previousPen != nullptr)
        SelectObject(metafileDc, previousPen);
    if (previousBrush != nullptr)
        SelectObject(metafileDc, previousBrush);
    if (formulaPen != nullptr)
        DeleteObject(formulaPen);
    if (borderPen != nullptr)
        DeleteObject(borderPen);

    HENHMETAFILE metafile = CloseEnhMetaFile(metafileDc);
    if (metafile == nullptr)
        return LastErrorResult();

    const UINT byteCount = GetEnhMetaFileBits(metafile, 0, nullptr);
    if (byteCount == 0)
    {
        DeleteEnhMetaFile(metafile);
        return LastErrorResult();
    }

    std::vector<BYTE> placeholderEmf(byteCount);
    if (GetEnhMetaFileBits(metafile, byteCount, placeholderEmf.data()) != byteCount)
    {
        DeleteEnhMetaFile(metafile);
        return LastErrorResult();
    }
    DeleteEnhMetaFile(metafile);
    if (!IsVectorEmf(placeholderEmf))
        return DV_E_FORMATETC;

    emfBytes_ = std::move(placeholderEmf);
    pngBytes_.assign(std::begin(kTransparentPng), std::end(kTransparentPng));
    return S_OK;
}

HRESULT CFormulaOleObject::CopyBytesToGlobal(const std::vector<BYTE>& bytes, HGLOBAL* global) const
{
    if (global == nullptr)
        return E_POINTER;
    *global = nullptr;
    HGLOBAL allocation = GlobalAlloc(GMEM_MOVEABLE, bytes.size());
    if (allocation == nullptr)
        return E_OUTOFMEMORY;
    void* destination = GlobalLock(allocation);
    if (destination == nullptr)
    {
        GlobalFree(allocation);
        return E_OUTOFMEMORY;
    }
    memcpy(destination, bytes.data(), bytes.size());
    GlobalUnlock(allocation);
    *global = allocation;
    return S_OK;
}

void CFormulaOleObject::UpdateExtentFromEmf(bool resetHostExtent) noexcept
{
    if (!IsValidEmf(emfBytes_))
    {
        naturalExtent_ = kDefaultExtent;
        if (resetHostExtent) extent_ = naturalExtent_;
        return;
    }
    HENHMETAFILE metafile = SetEnhMetaFileBits(static_cast<UINT>(emfBytes_.size()), emfBytes_.data());
    if (metafile == nullptr)
    {
        naturalExtent_ = kDefaultExtent;
        if (resetHostExtent) extent_ = naturalExtent_;
        return;
    }
    ENHMETAHEADER header = {};
    header.nSize = sizeof(header);
    if (GetEnhMetaFileHeader(metafile, sizeof(header), &header) == 0)
    {
        naturalExtent_ = kDefaultExtent;
    }
    else
    {
        const LONG width = header.rclFrame.right - header.rclFrame.left;
        const LONG height = header.rclFrame.bottom - header.rclFrame.top;
        naturalExtent_.cx = width > 0 ? width : kDefaultExtent.cx;
        naturalExtent_.cy = height > 0 ? height : kDefaultExtent.cy;
    }
    if (resetHostExtent) extent_ = naturalExtent_;
    DeleteEnhMetaFile(metafile);
}

void CFormulaOleObject::NotifyChanged() noexcept
{
    if (viewAdviseSink_ != nullptr)
        viewAdviseSink_->OnViewChange(DVASPECT_CONTENT, -1);
    if (dataAdviseHolder_ != nullptr)
        dataAdviseHolder_->SendOnDataChange(this, 0, 0);
}

bool CFormulaOleObject::DrawEmf(HDC drawContext, const RECT& bounds) const noexcept
{
    if (!IsValidEmf(emfBytes_))
        return false;
    HENHMETAFILE metafile = SetEnhMetaFileBits(static_cast<UINT>(emfBytes_.size()), emfBytes_.data());
    if (metafile == nullptr)
        return false;
    const BOOL drawn = PlayEnhMetaFile(drawContext, metafile, &bounds);
    DeleteEnhMetaFile(metafile);
    return drawn != FALSE;
}

bool CFormulaOleObject::DrawPng(HDC drawContext, const RECT& bounds) const noexcept
{
    if (!HasPngSignature(pngBytes_))
        return false;
    HGLOBAL global = nullptr;
    if (FAILED(CopyBytesToGlobal(pngBytes_, &global)))
        return false;
    CComPtr<IStream> stream;
    if (FAILED(CreateStreamOnHGlobal(global, TRUE, &stream)))
    {
        GlobalFree(global);
        return false;
    }

    Gdiplus::Bitmap image(stream, FALSE);
    if (image.GetLastStatus() != Gdiplus::Ok)
        return false;
    Gdiplus::Graphics graphics(drawContext);
    const Gdiplus::Rect destination(
        bounds.left,
        bounds.top,
        std::max<LONG>(1, bounds.right - bounds.left),
        std::max<LONG>(1, bounds.bottom - bounds.top));
    return graphics.DrawImage(&image, destination) == Gdiplus::Ok;
}

void CFormulaOleObject::DrawPlaceholder(HDC drawContext, const RECT& bounds) const noexcept
{
    FillRect(drawContext, &bounds, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
    FrameRect(drawContext, &bounds, static_cast<HBRUSH>(GetStockObject(GRAY_BRUSH)));
    RECT textBounds = bounds;
    SetBkMode(drawContext, TRANSPARENT);
    SetTextColor(drawContext, RGB(80, 80, 80));
    DrawTextW(
        drawContext,
        initialized_ ? L"VisualTeX formula preview unavailable" : L"VisualTeX formula",
        -1,
        &textBounds,
        DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
}

HRESULT CFormulaOleObject::Utf8ToWide(const std::vector<BYTE>& bytes, std::wstring& value)
{
    if (bytes.empty())
    {
        value.clear();
        return S_OK;
    }
    if (bytes.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
        return E_INVALIDARG;
    const int characters = MultiByteToWideChar(
        CP_UTF8,
        MB_ERR_INVALID_CHARS,
        reinterpret_cast<const char*>(bytes.data()),
        static_cast<int>(bytes.size()),
        nullptr,
        0);
    if (characters <= 0)
        return LastErrorResult();
    std::wstring decoded(characters, L'\0');
    if (MultiByteToWideChar(
            CP_UTF8,
            MB_ERR_INVALID_CHARS,
            reinterpret_cast<const char*>(bytes.data()),
            static_cast<int>(bytes.size()),
            decoded.data(),
            characters) != characters)
        return LastErrorResult();
    value = std::move(decoded);
    return S_OK;
}

HRESULT CFormulaOleObject::WideToUtf8(const std::wstring& value, std::vector<BYTE>& bytes)
{
    if (value.empty())
    {
        bytes.clear();
        return S_OK;
    }
    if (value.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
        return E_INVALIDARG;
    const int byteCount = WideCharToMultiByte(
        CP_UTF8,
        WC_ERR_INVALID_CHARS,
        value.data(),
        static_cast<int>(value.size()),
        nullptr,
        0,
        nullptr,
        nullptr);
    if (byteCount <= 0)
        return LastErrorResult();
    std::vector<BYTE> encoded(byteCount);
    if (WideCharToMultiByte(
            CP_UTF8,
            WC_ERR_INVALID_CHARS,
            value.data(),
            static_cast<int>(value.size()),
            reinterpret_cast<char*>(encoded.data()),
            byteCount,
            nullptr,
            nullptr) != byteCount)
        return LastErrorResult();
    bytes = std::move(encoded);
    return S_OK;
}
