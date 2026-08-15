#include <atlbase.h>
#include <atlcom.h>
#include <objbase.h>
#include <ole2.h>
#include <shlobj.h>
#include <shellscalingapi.h>
#include <tlhelp32.h>
#include <wincrypt.h>

#include <filesystem>
#include <fstream>
#include <iostream>
#include <stdexcept>
#include <string>
#include <vector>

#include "../VisualTeX.FormulaOleServer/FormulaOleContract.h"

#pragma comment(lib, "crypt32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "shcore.lib")
#pragma comment(lib, "uuid.lib")

namespace
{
class ComApartment final
{
public:
    ComApartment()
    {
        const HRESULT result = OleInitialize(nullptr);
        if (FAILED(result))
            throw std::runtime_error("OleInitialize failed");
        initialized_ = true;
    }

    ~ComApartment()
    {
        if (initialized_)
            OleUninitialize();
    }

private:
    bool initialized_ = false;
};

void Check(HRESULT result, const char* operation)
{
    if (FAILED(result))
    {
        std::cerr << operation << " failed: 0x" << std::hex << static_cast<unsigned long>(result) << std::endl;
        throw std::runtime_error(operation);
    }
}

DWORD RunServerCommand(const std::filesystem::path& server, const wchar_t* argument)
{
    std::wstring command = L"\"" + server.wstring() + L"\" " + argument;
    std::vector<wchar_t> mutableCommand(command.begin(), command.end());
    mutableCommand.push_back(L'\0');

    STARTUPINFOW startup = {};
    startup.cb = sizeof(startup);
    PROCESS_INFORMATION process = {};
    if (!CreateProcessW(
            nullptr,
            mutableCommand.data(),
            nullptr,
            nullptr,
            FALSE,
            CREATE_NO_WINDOW,
            nullptr,
            server.parent_path().c_str(),
            &startup,
            &process))
        throw std::runtime_error("CreateProcessW failed");

    WaitForSingleObject(process.hProcess, 30000);
    DWORD exitCode = ERROR_GEN_FAILURE;
    GetExitCodeProcess(process.hProcess, &exitCode);
    CloseHandle(process.hThread);
    CloseHandle(process.hProcess);
    return exitCode;
}

class EmbeddedServerProcess final
{
public:
    explicit EmbeddedServerProcess(const std::filesystem::path& server)
    {
        std::wstring command = L"\"" + server.wstring() + L"\" -Embedding";
        std::vector<wchar_t> mutableCommand(command.begin(), command.end());
        mutableCommand.push_back(L'\0');

        STARTUPINFOW startup = {};
        startup.cb = sizeof(startup);
        if (!CreateProcessW(
                nullptr,
                mutableCommand.data(),
                nullptr,
                nullptr,
                FALSE,
                CREATE_NO_WINDOW,
                nullptr,
                server.parent_path().c_str(),
                &startup,
                &process_))
            throw std::runtime_error("Unable to start acceptance-owned OLE LocalServer -Embedding process");

        CloseHandle(process_.hThread);
        process_.hThread = nullptr;

        // ATL registers its class objects during PreMessageLoop. Give the
        // explicitly-started test server a bounded startup window before any
        // CoCreate/OleCreate call can fall back to a machine-installed server
        // carrying the same production CLSID.
        const DWORD deadline = GetTickCount() + 3000;
        while (GetTickCount() < deadline)
        {
            if (WaitForSingleObject(process_.hProcess, 0) == WAIT_OBJECT_0)
                throw std::runtime_error("Acceptance-owned OLE LocalServer exited before registering its class factory");
            if (GetTickCount() + 500 >= deadline)
                break;
            Sleep(100);
        }
        Sleep(500);
    }

    ~EmbeddedServerProcess()
    {
        if (process_.hProcess == nullptr)
            return;

        // Once all COM references in RunSmoke have been released, the server
        // should leave naturally after its bounded startup-grace lock. Wait for
        // that normal ATL shutdown first. Only if the acceptance-owned process
        // fails to exit do we terminate that exact process handle so it cannot
        // contaminate the next architecture's same-CLSID smoke.
        if (WaitForSingleObject(process_.hProcess, 20000) != WAIT_OBJECT_0)
            TerminateProcess(process_.hProcess, ERROR_TIMEOUT);
        CloseHandle(process_.hProcess);
        process_.hProcess = nullptr;
    }

    DWORD ProcessId() const noexcept
    {
        return process_.dwProcessId;
    }

private:
    PROCESS_INFORMATION process_ = {};
};

class ServerRegistration final
{
public:
    explicit ServerRegistration(std::filesystem::path server) : server_(std::move(server))
    {
        if (RunServerCommand(server_, L"/RegServerPerUser") != ERROR_SUCCESS)
            throw std::runtime_error("LocalServer registration failed");
        registered_ = true;
    }

    ~ServerRegistration()
    {
        if (registered_)
            RunServerCommand(server_, L"/UnregServerPerUser");
    }

private:
    std::filesystem::path server_;
    bool registered_ = false;
};

std::filesystem::path OfficeTempDirectory()
{
    PWSTR localApplicationData = nullptr;
    Check(
        SHGetKnownFolderPath(FOLDERID_LocalAppData, KF_FLAG_DEFAULT, nullptr, &localApplicationData),
        "SHGetKnownFolderPath");
    std::filesystem::path directory(localApplicationData);
    CoTaskMemFree(localApplicationData);
    directory /= L"VisualTeX";
    directory /= L"office";
    directory /= L"temp";
    std::filesystem::create_directories(directory);
    return directory;
}

void WritePng(const std::filesystem::path& path)
{
    constexpr char encoded[] =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9ZQmcAAAAASUVORK5CYII=";
    DWORD byteCount = 0;
    if (!CryptStringToBinaryA(encoded, 0, CRYPT_STRING_BASE64, nullptr, &byteCount, nullptr, nullptr))
        throw std::runtime_error("CryptStringToBinaryA size failed");
    std::vector<BYTE> bytes(byteCount);
    if (!CryptStringToBinaryA(
            encoded,
            0,
            CRYPT_STRING_BASE64,
            bytes.data(),
            &byteCount,
            nullptr,
            nullptr))
        throw std::runtime_error("CryptStringToBinaryA decode failed");
    bytes.resize(byteCount);
    std::ofstream output(path, std::ios::binary | std::ios::trunc);
    output.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
    if (!output)
        throw std::runtime_error("PNG write failed");
}

void WriteEmf(const std::filesystem::path& path)
{
    HDC screen = GetDC(nullptr);
    RECT frame = {0, 0, 5000, 1200};
    HDC metafileDc = CreateEnhMetaFileW(
        screen,
        path.c_str(),
        &frame,
        L"VisualTeX\0Formula OLE smoke preview\0\0");
    ReleaseDC(nullptr, screen);
    if (metafileDc == nullptr)
        throw std::runtime_error("CreateEnhMetaFileW failed");

    SetBkMode(metafileDc, TRANSPARENT);
    TextOutW(metafileDc, 20, 20, L"x = y + 1", 9);
    MoveToEx(metafileDc, 20, 80, nullptr);
    LineTo(metafileDc, 420, 80);
    HENHMETAFILE metafile = CloseEnhMetaFile(metafileDc);
    if (metafile == nullptr)
        throw std::runtime_error("CloseEnhMetaFile failed");
    DeleteEnhMetaFile(metafile);
}

void WriteRasterEmf(const std::filesystem::path& path)
{
    HDC screen = GetDC(nullptr);
    HDC source = CreateCompatibleDC(screen);
    HBITMAP bitmap = CreateCompatibleBitmap(screen, 32, 32);
    if (source == nullptr || bitmap == nullptr)
    {
        if (bitmap != nullptr) DeleteObject(bitmap);
        if (source != nullptr) DeleteDC(source);
        ReleaseDC(nullptr, screen);
        throw std::runtime_error("Raster EMF source allocation failed");
    }
    HGDIOBJ previous = SelectObject(source, bitmap);
    PatBlt(source, 0, 0, 32, 32, WHITENESS);
    SetPixel(source, 4, 4, RGB(255, 0, 0));

    RECT frame = {0, 0, 3200, 3200};
    HDC metafileDc = CreateEnhMetaFileW(
        screen,
        path.c_str(),
        &frame,
        L"VisualTeX\0Forbidden raster OLE smoke preview\0\0");
    if (metafileDc == nullptr)
    {
        SelectObject(source, previous);
        DeleteObject(bitmap);
        DeleteDC(source);
        ReleaseDC(nullptr, screen);
        throw std::runtime_error("CreateEnhMetaFileW for raster preview failed");
    }
    StretchBlt(metafileDc, 0, 0, 320, 320, source, 0, 0, 32, 32, SRCCOPY);
    HENHMETAFILE metafile = CloseEnhMetaFile(metafileDc);

    SelectObject(source, previous);
    DeleteObject(bitmap);
    DeleteDC(source);
    ReleaseDC(nullptr, screen);
    if (metafile == nullptr)
        throw std::runtime_error("CloseEnhMetaFile for raster preview failed");
    DeleteEnhMetaFile(metafile);
}

void VerifyStream(IStorage* storage, const wchar_t* name)
{
    CComPtr<IStream> stream;
    Check(storage->OpenStream(name, nullptr, STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &stream), "OpenStream");
    STATSTG stat = {};
    Check(stream->Stat(&stat, STATFLAG_NONAME), "IStream::Stat");
    if (stat.cbSize.QuadPart <= 0)
        throw std::runtime_error("Persisted OLE stream is empty");
}

void VerifyOleCreateProtocol(const std::filesystem::path& temp, const std::wstring& suffix)
{
    for (const auto renderOption : {OLERENDER_NONE, OLERENDER_DRAW})
    {
        const wchar_t* label = renderOption == OLERENDER_NONE
            ? L"OLERENDER_NONE"
            : L"OLERENDER_DRAW";
        const std::filesystem::path storagePath =
            temp / (L"ole-create-" + std::wstring(label) + L"-" + suffix + L".ole");
        CComPtr<IStorage> storage;
        Check(
            StgCreateDocfile(
                storagePath.c_str(),
                STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE,
                0,
                &storage),
            "StgCreateDocfile(OleCreate)");

        CComPtr<IOleObject> object;
        const HRESULT result = OleCreate(
            CLSID_VisualTeXFormula,
            IID_IOleObject,
            renderOption,
            nullptr,
            nullptr,
            storage,
            reinterpret_cast<void**>(&object));
        std::wcout << L"OleCreate " << label << L" -> 0x"
                   << std::hex << static_cast<unsigned long>(result) << std::dec << std::endl;
        if (renderOption == OLERENDER_NONE)
            Check(result, "OleCreate(OLERENDER_NONE)");
        if (SUCCEEDED(result) && object == nullptr)
            throw std::runtime_error("OleCreate succeeded without returning IOleObject");

        object.Release();
        storage.Release();
        std::error_code ignored;
        std::filesystem::remove(storagePath, ignored);
    }
}

void VerifyServerDpiAwareness(const std::filesystem::path& expectedServer)
{
    const std::wstring expected = std::filesystem::weakly_canonical(expectedServer).wstring();
    bool found = false;

    // COM can return the LocalServer proxy just before the newly-created process
    // becomes visible in a Toolhelp snapshot. Retry only the process-discovery
    // portion for a short bounded interval; the DPI assertion itself remains
    // strict once the expected server executable is observed.
    for (int attempt = 0; attempt < 20 && !found; ++attempt)
    {
        HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        if (snapshot == INVALID_HANDLE_VALUE)
            throw std::runtime_error("CreateToolhelp32Snapshot failed while checking OLE server DPI awareness");

        PROCESSENTRY32W entry = {};
        entry.dwSize = sizeof(entry);
        if (Process32FirstW(snapshot, &entry))
        {
            do
            {
                HANDLE process = OpenProcess(
                    PROCESS_QUERY_LIMITED_INFORMATION,
                    FALSE,
                    entry.th32ProcessID);
                if (process == nullptr)
                    continue;
                wchar_t pathBuffer[32768] = {};
                DWORD pathLength = static_cast<DWORD>(std::size(pathBuffer));
                if (QueryFullProcessImageNameW(process, 0, pathBuffer, &pathLength))
                {
                    std::error_code ignored;
                    const std::wstring candidate = std::filesystem::weakly_canonical(
                        std::filesystem::path(std::wstring(pathBuffer, pathLength)),
                        ignored).wstring();
                    if (!ignored && _wcsicmp(candidate.c_str(), expected.c_str()) == 0)
                    {
                        PROCESS_DPI_AWARENESS awareness = PROCESS_DPI_UNAWARE;
                        Check(GetProcessDpiAwareness(process, &awareness), "GetProcessDpiAwareness(OLE server)");
                        if (awareness != PROCESS_PER_MONITOR_DPI_AWARE)
                        {
                            CloseHandle(process);
                            CloseHandle(snapshot);
                            throw std::runtime_error(
                                "Formula OLE LocalServer is not per-monitor DPI aware; legacy MFPICT can shrink into the upper-left corner");
                        }
                        found = true;
                        CloseHandle(process);
                        break;
                    }
                }
                CloseHandle(process);
            } while (Process32NextW(snapshot, &entry));
        }
        CloseHandle(snapshot);
        if (!found)
            Sleep(100);
    }

    if (!found)
        throw std::runtime_error("Running Formula OLE LocalServer process was not found for DPI verification");
}

CComPtr<IVisualTeXFormulaObject> CreateFormulaObject()
{
    CComPtr<IUnknown> unknown;
    Check(
        CoCreateInstance(
            CLSID_VisualTeXFormula,
            nullptr,
            CLSCTX_LOCAL_SERVER,
            IID_IUnknown,
            reinterpret_cast<void**>(&unknown)),
        "CoCreateInstance(IUnknown)");

    CComQIPtr<IOleObject> oleObject(unknown);
    if (oleObject == nullptr)
        throw std::runtime_error("IOleObject is unavailable after activation");

    CComQIPtr<IVisualTeXFormulaObject> formulaObject(unknown);
    if (formulaObject == nullptr)
        throw std::runtime_error("IVisualTeXFormulaObject is unavailable after activation");
    return formulaObject;
}

void VerifyDataAndView(IUnknown* object)
{
    CComQIPtr<IDataObject> dataObject(object);
    if (dataObject == nullptr)
        throw std::runtime_error("IDataObject is unavailable");

    FORMATETC emfFormat = {};
    emfFormat.cfFormat = CF_ENHMETAFILE;
    emfFormat.dwAspect = DVASPECT_CONTENT;
    emfFormat.lindex = -1;
    emfFormat.tymed = TYMED_ENHMF;
    Check(dataObject->QueryGetData(&emfFormat), "QueryGetData(CF_ENHMETAFILE)");
    STGMEDIUM emfMedium = {};
    Check(dataObject->GetData(&emfFormat, &emfMedium), "GetData(CF_ENHMETAFILE)");
    ReleaseStgMedium(&emfMedium);

    FORMATETC metafilePictureFormat = {};
    metafilePictureFormat.cfFormat = CF_METAFILEPICT;
    metafilePictureFormat.dwAspect = DVASPECT_CONTENT;
    metafilePictureFormat.lindex = -1;
    metafilePictureFormat.tymed = TYMED_MFPICT;
    Check(
        dataObject->QueryGetData(&metafilePictureFormat),
        "QueryGetData(CF_METAFILEPICT)");
    STGMEDIUM metafilePictureMedium = {};
    Check(
        dataObject->GetData(&metafilePictureFormat, &metafilePictureMedium),
        "GetData(CF_METAFILEPICT)");
    if (metafilePictureMedium.tymed != TYMED_MFPICT
        || metafilePictureMedium.hMetaFilePict == nullptr)
        throw std::runtime_error("CF_METAFILEPICT returned an invalid medium");
    ReleaseStgMedium(&metafilePictureMedium);

    const UINT pngFormatId = RegisterClipboardFormatW(L"PNG");
    FORMATETC pngFormat = {};
    pngFormat.cfFormat = static_cast<CLIPFORMAT>(pngFormatId);
    pngFormat.dwAspect = DVASPECT_CONTENT;
    pngFormat.lindex = -1;
    pngFormat.tymed = TYMED_HGLOBAL;
    Check(dataObject->QueryGetData(&pngFormat), "QueryGetData(PNG)");
    STGMEDIUM pngMedium = {};
    Check(dataObject->GetData(&pngFormat, &pngMedium), "GetData(PNG)");
    ReleaseStgMedium(&pngMedium);

    CComQIPtr<IViewObject2> viewObject(object);
    if (viewObject == nullptr)
        throw std::runtime_error("IViewObject2 is unavailable");
    SIZEL extent = {};
    Check(viewObject->GetExtent(DVASPECT_CONTENT, -1, nullptr, &extent), "IViewObject2::GetExtent");
    if (extent.cx <= 0 || extent.cy <= 0)
        throw std::runtime_error("OLE extent is invalid");

    HDC screen = GetDC(nullptr);
    HDC memory = CreateCompatibleDC(screen);
    HBITMAP bitmap = CreateCompatibleBitmap(screen, 640, 180);
    HGDIOBJ previous = SelectObject(memory, bitmap);
    RECTL bounds = {0, 0, 640, 180};
    Check(
        viewObject->Draw(
            DVASPECT_CONTENT,
            -1,
            nullptr,
            nullptr,
            screen,
            memory,
            &bounds,
            nullptr,
            nullptr,
            0),
        "IViewObject2::Draw");
    SelectObject(memory, previous);
    DeleteObject(bitmap);
    DeleteDC(memory);
    ReleaseDC(nullptr, screen);
}

void VerifyPlaceholderPersistence(
    const std::filesystem::path& storagePath,
    const std::filesystem::path& emf,
    const std::filesystem::path& png,
    const std::wstring& metadata)
{
    CComPtr<IStorage> storage;
    Check(
        StgCreateDocfile(
            storagePath.c_str(),
            STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE,
            0,
            &storage),
        "StgCreateDocfile(placeholder)");

    CComPtr<IVisualTeXFormulaObject> formula = CreateFormulaObject();
    CComQIPtr<IPersistStorage> persist(formula);
    Check(persist->InitNew(storage), "IPersistStorage::InitNew(placeholder)");
    Check(persist->Save(storage, TRUE), "IPersistStorage::Save(placeholder)");
    Check(persist->SaveCompleted(storage), "IPersistStorage::SaveCompleted(placeholder)");
    Check(storage->Commit(STGC_DEFAULT), "IStorage::Commit(placeholder)");
    formula.Release();
    persist.Release();
    storage.Release();

    Check(
        StgOpenStorage(
            storagePath.c_str(),
            nullptr,
            STGM_READWRITE | STGM_SHARE_EXCLUSIVE,
            nullptr,
            0,
            &storage),
        "StgOpenStorage(placeholder)");
    formula = CreateFormulaObject();
    persist = formula;
    Check(persist->Load(storage), "IPersistStorage::Load(placeholder)");
    CComBSTR placeholderJson;
    const HRESULT placeholderRead = formula->GetFormulaJson(&placeholderJson);
    if (placeholderRead != CO_E_NOTINITIALIZED)
        throw std::runtime_error("Placeholder object unexpectedly exposed formula metadata");
    Check(
        formula->InitializeFromFiles(
            CComBSTR(metadata.c_str()),
            CComBSTR(emf.c_str()),
            CComBSTR(png.c_str())),
        "InitializeFromFiles(after placeholder reload)");
    Check(persist->Save(storage, TRUE), "IPersistStorage::Save(after placeholder reload)");
    Check(persist->SaveCompleted(storage), "IPersistStorage::SaveCompleted(after placeholder reload)");
    formula.Release();
    persist.Release();
    storage.Release();
    std::error_code ignored;
    std::filesystem::remove(storagePath, ignored);
}

void VerifyPreinitializeHostExtentPreserved(
    const std::filesystem::path& storagePath,
    const std::filesystem::path& emf,
    const std::filesystem::path& png,
    const std::wstring& metadata)
{
    CComPtr<IStorage> storage;
    Check(
        StgCreateDocfile(
            storagePath.c_str(),
            STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE,
            0,
            &storage),
        "StgCreateDocfile(preinitialize host extent)");

    CComPtr<IVisualTeXFormulaObject> formula = CreateFormulaObject();
    CComQIPtr<IPersistStorage> persist(formula);
    CComQIPtr<IOleObject> oleObject(formula);
    if (persist == nullptr || oleObject == nullptr)
        throw std::runtime_error("Required OLE interfaces are unavailable for host extent acceptance");
    Check(persist->InitNew(storage), "IPersistStorage::InitNew(preinitialize host extent)");

    // PowerPoint 2021 Home can size the newly allocated OLE object before
    // VisualTeX initializes its EMF/PNG preview. InitializeFromFiles must not
    // replace this explicit container-owned extent with the preview's natural
    // extent, otherwise the formula is painted at natural size in the upper-left
    // corner of a larger PowerPoint Shape.
    const SIZEL requested = {8467, 3387}; // approximately 240pt x 96pt
    SIZEL actual = {};
    Check(oleObject->SetExtent(DVASPECT_CONTENT, const_cast<SIZEL*>(&requested)),
        "IOleObject::SetExtent(preinitialize host extent)");
    Check(
        formula->InitializeFromFiles(
            CComBSTR(metadata.c_str()),
            CComBSTR(emf.c_str()),
            CComBSTR(png.c_str())),
        "InitializeFromFiles(preinitialize host extent)");
    Check(oleObject->GetExtent(DVASPECT_CONTENT, &actual),
        "IOleObject::GetExtent(preinitialize host extent)");
    if (actual.cx != requested.cx || actual.cy != requested.cy)
        throw std::runtime_error("InitializeFromFiles replaced the explicit PowerPoint host extent");

    CComQIPtr<IDataObject> dataObject(formula);
    if (dataObject == nullptr)
        throw std::runtime_error("IDataObject is unavailable for host extent acceptance");
    FORMATETC metafilePictureFormat = {};
    metafilePictureFormat.cfFormat = CF_METAFILEPICT;
    metafilePictureFormat.dwAspect = DVASPECT_CONTENT;
    metafilePictureFormat.lindex = -1;
    metafilePictureFormat.tymed = TYMED_MFPICT;
    STGMEDIUM metafilePictureMedium = {};
    Check(
        dataObject->GetData(&metafilePictureFormat, &metafilePictureMedium),
        "GetData(CF_METAFILEPICT preinitialize host extent)");
    auto* picture = static_cast<METAFILEPICT*>(GlobalLock(metafilePictureMedium.hMetaFilePict));
    if (picture == nullptr)
    {
        ReleaseStgMedium(&metafilePictureMedium);
        throw std::runtime_error("CF_METAFILEPICT could not be locked for host extent acceptance");
    }
    const LONG presentationWidth = picture->xExt;
    const LONG presentationHeight = picture->yExt;
    GlobalUnlock(metafilePictureMedium.hMetaFilePict);
    ReleaseStgMedium(&metafilePictureMedium);
    if (presentationWidth != requested.cx || presentationHeight != requested.cy)
        throw std::runtime_error("Legacy PowerPoint OLE presentation did not preserve the explicit host extent");

    dataObject.Release();
    formula.Release();
    oleObject.Release();
    persist.Release();
    storage.Release();
    std::error_code ignored;
    std::filesystem::remove(storagePath, ignored);
}

void VerifyPostinitializeHostExtentLocksPresentation(
    const std::filesystem::path& storagePath,
    const std::filesystem::path& emf,
    const std::filesystem::path& png,
    const std::wstring& metadata)
{
    CComPtr<IStorage> storage;
    Check(
        StgCreateDocfile(
            storagePath.c_str(),
            STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE,
            0,
            &storage),
        "StgCreateDocfile(postinitialize host extent)");

    CComPtr<IVisualTeXFormulaObject> formula = CreateFormulaObject();
    CComQIPtr<IPersistStorage> persist(formula);
    CComQIPtr<IOleObject> oleObject(formula);
    CComQIPtr<IDataObject> dataObject(formula);
    if (persist == nullptr || oleObject == nullptr || dataObject == nullptr)
        throw std::runtime_error("Required OLE interfaces are unavailable for postinitialize host extent acceptance");
    Check(persist->InitNew(storage), "IPersistStorage::InitNew(postinitialize host extent)");
    Check(
        formula->InitializeFromFiles(
            CComBSTR(metadata.c_str()),
            CComBSTR(emf.c_str()),
            CComBSTR(png.c_str())),
        "InitializeFromFiles(postinitialize host extent)");

    // Other PowerPoint builds initialize first and call SetExtent afterwards.
    // The first explicit container extent must become the intrinsic presentation
    // extent used by legacy MFPICT caches. Later layout/resizing calls may change
    // the live host extent but must not compound the cached presentation size.
    SIZEL initialHost = {8467, 3387}; // approximately 240pt x 96pt
    Check(oleObject->SetExtent(DVASPECT_CONTENT, &initialHost),
        "IOleObject::SetExtent(postinitialize initial host extent)");

    auto readPresentationExtent = [&]() -> SIZEL
    {
        FORMATETC format = {};
        format.cfFormat = CF_METAFILEPICT;
        format.dwAspect = DVASPECT_CONTENT;
        format.lindex = -1;
        format.tymed = TYMED_MFPICT;
        STGMEDIUM medium = {};
        Check(dataObject->GetData(&format, &medium),
            "GetData(CF_METAFILEPICT postinitialize host extent)");
        auto* picture = static_cast<METAFILEPICT*>(GlobalLock(medium.hMetaFilePict));
        if (picture == nullptr)
        {
            ReleaseStgMedium(&medium);
            throw std::runtime_error("CF_METAFILEPICT could not be locked for postinitialize host extent acceptance");
        }
        SIZEL result = {picture->xExt, picture->yExt};
        GlobalUnlock(medium.hMetaFilePict);
        ReleaseStgMedium(&medium);
        return result;
    };

    SIZEL presentation = readPresentationExtent();
    if (presentation.cx != initialHost.cx || presentation.cy != initialHost.cy)
        throw std::runtime_error("First post-initialize PowerPoint host extent was not adopted by the legacy presentation");

    SIZEL resizedHost = {12700, 5080}; // approximately 360pt x 144pt
    Check(oleObject->SetExtent(DVASPECT_CONTENT, &resizedHost),
        "IOleObject::SetExtent(postinitialize resized host extent)");
    SIZEL liveExtent = {};
    Check(oleObject->GetExtent(DVASPECT_CONTENT, &liveExtent),
        "IOleObject::GetExtent(postinitialize resized host extent)");
    if (liveExtent.cx != resizedHost.cx || liveExtent.cy != resizedHost.cy)
        throw std::runtime_error("Live OLE host extent did not follow the later PowerPoint resize");
    presentation = readPresentationExtent();
    if (presentation.cx != initialHost.cx || presentation.cy != initialHost.cy)
        throw std::runtime_error("Later PowerPoint resizing compounded the intrinsic legacy presentation extent");

    dataObject.Release();
    formula.Release();
    oleObject.Release();
    persist.Release();
    storage.Release();
    std::error_code ignored;
    std::filesystem::remove(storagePath, ignored);
}

void RunSmoke(const std::filesystem::path& server)
{
    ServerRegistration registration(server);
    ComApartment apartment;
    EmbeddedServerProcess embeddedServer(server);

    // Keep the explicitly-started server locked for the full smoke so transient
    // gaps between individual OleCreate/CoCreate calls cannot let SCM fall back
    // to a machine-installed instance with the same CLSID.
    CComPtr<IVisualTeXFormulaObject> serverKeepAlive = CreateFormulaObject();
    VerifyServerDpiAwareness(server);

    const std::filesystem::path temp = OfficeTempDirectory();
    const std::wstring suffix = std::to_wstring(GetCurrentProcessId());
    VerifyOleCreateProtocol(temp, suffix);
    const std::filesystem::path emf = temp / (L"ole-smoke-" + suffix + L".emf");
    const std::filesystem::path rasterEmf = temp / (L"ole-smoke-raster-" + suffix + L".emf");
    const std::filesystem::path png = temp / (L"ole-smoke-" + suffix + L".png");
    const std::filesystem::path storagePath = temp / (L"ole-smoke-" + suffix + L".ole");
    const std::filesystem::path placeholderStoragePath = temp / (L"ole-placeholder-" + suffix + L".ole");
    const std::filesystem::path hostExtentStoragePath = temp / (L"ole-host-extent-" + suffix + L".ole");
    const std::filesystem::path postHostExtentStoragePath = temp / (L"ole-post-host-extent-" + suffix + L".ole");
    WriteEmf(emf);
    WriteRasterEmf(rasterEmf);
    WritePng(png);

    const std::wstring metadata =
        LR"({"schemaVersion":1,"formulaId":"11111111-2222-4333-8444-555555555555","title":"Smoke","latex":"x=y+1","lines":[{"id":"aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee","latex":"x=y+1"}],"codeFormat":"raw","displayMode":"inline","numbered":false,"renderWidthPx":320,"renderHeightPx":80,"baseline":62})";

    VerifyPlaceholderPersistence(placeholderStoragePath, emf, png, metadata);
    VerifyPreinitializeHostExtentPreserved(hostExtentStoragePath, emf, png, metadata);
    VerifyPostinitializeHostExtentLocksPresentation(postHostExtentStoragePath, emf, png, metadata);

    CComPtr<IStorage> storage;
    Check(
        StgCreateDocfile(
            storagePath.c_str(),
            STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE,
            0,
            &storage),
        "StgCreateDocfile");

    CComPtr<IVisualTeXFormulaObject> formula = CreateFormulaObject();
    CComQIPtr<IPersistStorage> persist(formula);
    if (persist == nullptr)
        throw std::runtime_error("IPersistStorage is unavailable");
    Check(persist->InitNew(storage), "IPersistStorage::InitNew");
    Check(
        formula->InitializeFromFiles(
            CComBSTR(metadata.c_str()),
            CComBSTR(emf.c_str()),
            CComBSTR(png.c_str())),
        "InitializeFromFiles");
    VerifyDataAndView(formula);

    CComQIPtr<IVisualTeXFormulaMetadata> metadataWriter(formula);
    if (metadataWriter == nullptr)
        throw std::runtime_error("IVisualTeXFormulaMetadata is unavailable");
    const std::wstring numberedMetadata =
        LR"({"schemaVersion":1,"formulaId":"11111111-2222-4333-8444-555555555555","title":"Smoke","latex":"x=y+1","lines":[{"id":"aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee","latex":"x=y+1"}],"codeFormat":"raw","displayMode":"block","numbered":true,"renderWidthPx":320,"renderHeightPx":80,"baseline":62})";
    Check(
        metadataWriter->SetFormulaJson(CComBSTR(numberedMetadata.c_str())),
        "SetFormulaJson");
    CComBSTR afterMetadataOnlyUpdate;
    Check(
        formula->GetFormulaJson(&afterMetadataOnlyUpdate),
        "GetFormulaJson(after metadata-only update)");
    if (std::wstring(afterMetadataOnlyUpdate, afterMetadataOnlyUpdate.Length()) != numberedMetadata)
        throw std::runtime_error("Metadata-only update did not persist in memory");

    const HRESULT invalidMetadataUpdate = metadataWriter->SetFormulaJson(CComBSTR(L"{}"));
    if (SUCCEEDED(invalidMetadataUpdate))
        throw std::runtime_error("Invalid metadata-only update unexpectedly succeeded");
    CComBSTR afterFailedMetadataUpdate;
    Check(
        formula->GetFormulaJson(&afterFailedMetadataUpdate),
        "GetFormulaJson(after failed metadata-only update)");
    if (std::wstring(afterFailedMetadataUpdate, afterFailedMetadataUpdate.Length()) != numberedMetadata)
        throw std::runtime_error("Failed metadata-only update mutated the formula");

    CComBSTR beforeFailedUpdate;
    Check(formula->GetFormulaJson(&beforeFailedUpdate), "GetFormulaJson(before failed update)");
    const HRESULT invalidUpdate = formula->UpdateFromFiles(
        CComBSTR(metadata.c_str()),
        CComBSTR(rasterEmf.c_str()),
        CComBSTR(png.c_str()));
    if (SUCCEEDED(invalidUpdate))
        throw std::runtime_error("Raster EMF update unexpectedly succeeded");
    CComBSTR afterFailedUpdate;
    Check(formula->GetFormulaJson(&afterFailedUpdate), "GetFormulaJson(after failed update)");
    if (std::wstring(beforeFailedUpdate, beforeFailedUpdate.Length()) !=
        std::wstring(afterFailedUpdate, afterFailedUpdate.Length()))
        throw std::runtime_error("Failed update mutated the formula");

    Check(persist->Save(storage, TRUE), "IPersistStorage::Save");
    Check(persist->SaveCompleted(storage), "IPersistStorage::SaveCompleted");
    Check(storage->Commit(STGC_DEFAULT), "IStorage::Commit");
    VerifyStream(storage, kVisualTeXMetadataStream);
    VerifyStream(storage, kVisualTeXEmfPreviewStream);
    VerifyStream(storage, kVisualTeXPngPreviewStream);

    metadataWriter.Release();
    formula.Release();
    persist.Release();
    storage.Release();

    Check(
        StgOpenStorage(
            storagePath.c_str(),
            nullptr,
            STGM_READ | STGM_SHARE_EXCLUSIVE,
            nullptr,
            0,
            &storage),
        "StgOpenStorage");
    formula = CreateFormulaObject();
    persist = formula;
    Check(persist->Load(storage), "IPersistStorage::Load");
    if (persist->IsDirty() != S_FALSE)
        throw std::runtime_error("Loaded object should be clean");

    CComBSTR loadedMetadata;
    Check(formula->GetFormulaJson(&loadedMetadata), "GetFormulaJson(loaded)");
    if (std::wstring(loadedMetadata, loadedMetadata.Length()) != numberedMetadata)
        throw std::runtime_error("Metadata-only update did not round-trip through structured storage");
    VerifyDataAndView(formula);

    formula.Release();
    persist.Release();
    storage.Release();
    std::error_code ignored;
    std::filesystem::remove(storagePath, ignored);
    std::filesystem::remove(emf, ignored);
    std::filesystem::remove(rasterEmf, ignored);
    std::filesystem::remove(png, ignored);
}
} // namespace

int wmain(int argc, wchar_t** argv)
{
    if (argc != 2)
    {
        std::wcerr << L"Usage: VisualTeX.FormulaOleServer.Tests.exe <FormulaOleServer.exe>" << std::endl;
        return 2;
    }

    try
    {
        const std::filesystem::path server = std::filesystem::absolute(argv[1]);
        if (!std::filesystem::is_regular_file(server))
            throw std::runtime_error("LocalServer executable does not exist");
        RunSmoke(server);
        std::wcout << L"VisualTeX Formula OLE LocalServer smoke test passed" << std::endl;
        return 0;
    }
    catch (const std::exception& error)
    {
        std::cerr << "VisualTeX Formula OLE LocalServer smoke test failed: " << error.what() << std::endl;
        return 1;
    }
}
