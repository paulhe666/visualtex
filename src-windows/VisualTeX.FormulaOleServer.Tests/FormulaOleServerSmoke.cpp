#include <atlbase.h>
#include <atlcom.h>
#include <objbase.h>
#include <ole2.h>
#include <shlobj.h>
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

void RunSmoke(const std::filesystem::path& server)
{
    ServerRegistration registration(server);
    ComApartment apartment;

    const std::filesystem::path temp = OfficeTempDirectory();
    const std::wstring suffix = std::to_wstring(GetCurrentProcessId());
    VerifyOleCreateProtocol(temp, suffix);
    const std::filesystem::path emf = temp / (L"ole-smoke-" + suffix + L".emf");
    const std::filesystem::path rasterEmf = temp / (L"ole-smoke-raster-" + suffix + L".emf");
    const std::filesystem::path png = temp / (L"ole-smoke-" + suffix + L".png");
    const std::filesystem::path storagePath = temp / (L"ole-smoke-" + suffix + L".ole");
    const std::filesystem::path placeholderStoragePath = temp / (L"ole-placeholder-" + suffix + L".ole");
    WriteEmf(emf);
    WriteRasterEmf(rasterEmf);
    WritePng(png);

    const std::wstring metadata =
        LR"({"schemaVersion":1,"formulaId":"11111111-2222-4333-8444-555555555555","title":"Smoke","latex":"x=y+1","lines":[{"id":"aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee","latex":"x=y+1"}],"codeFormat":"raw","displayMode":"inline","numbered":false,"renderWidthPx":320,"renderHeightPx":80,"baseline":62})";

    VerifyPlaceholderPersistence(placeholderStoragePath, emf, png, metadata);

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
    if (std::wstring(loadedMetadata, loadedMetadata.Length()) != metadata)
        throw std::runtime_error("Metadata did not round-trip through structured storage");
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
