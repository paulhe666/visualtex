#define NOMINMAX
#include <windows.h>
#include <dbghelp.h>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>
#include <vector>

#pragma comment(lib, "dbghelp.lib")

static std::wstring utf8_to_wide(const std::string& value) {
    if (value.empty()) return {};
    int count = MultiByteToWideChar(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), nullptr, 0);
    std::wstring result(count, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), result.data(), count);
    return result;
}

static std::string wide_to_utf8(const std::wstring& value) {
    if (value.empty()) return {};
    int count = WideCharToMultiByte(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), nullptr, 0, nullptr, nullptr);
    std::string result(count, '\0');
    WideCharToMultiByte(CP_UTF8, 0, value.c_str(), static_cast<int>(value.size()), result.data(), count, nullptr, nullptr);
    return result;
}

static std::filesystem::path temp_log_path() {
    wchar_t buffer[MAX_PATH]{};
    DWORD count = GetTempPathW(MAX_PATH, buffer);
    if (count == 0 || count >= MAX_PATH) return L"p13-native-crash-capture.log";
    return std::filesystem::path(buffer) / L"p13-native-crash-capture.log";
}

static std::filesystem::path temp_pid_path() {
    wchar_t buffer[MAX_PATH]{};
    DWORD count = GetTempPathW(MAX_PATH, buffer);
    if (count == 0 || count >= MAX_PATH) return L"p13-native-crash-child.pid";
    return std::filesystem::path(buffer) / L"p13-native-crash-child.pid";
}

static void log_line(std::ofstream& log, const std::string& text) {
    SYSTEMTIME time{};
    GetLocalTime(&time);
    log << std::setfill('0')
        << std::setw(2) << time.wHour << ':'
        << std::setw(2) << time.wMinute << ':'
        << std::setw(2) << time.wSecond << '.'
        << std::setw(3) << time.wMilliseconds << ' '
        << text << '\n';
    log.flush();
}

static std::string symbol_for(HANDLE process, DWORD64 address) {
    std::vector<unsigned char> storage(sizeof(SYMBOL_INFO) + MAX_SYM_NAME + 1);
    auto* symbol = reinterpret_cast<SYMBOL_INFO*>(storage.data());
    symbol->SizeOfStruct = sizeof(SYMBOL_INFO);
    symbol->MaxNameLen = MAX_SYM_NAME;
    DWORD64 displacement = 0;
    std::ostringstream out;
    out << "0x" << std::hex << address;
    if (SymFromAddr(process, address, &displacement, symbol)) {
        out << " " << symbol->Name;
        if (displacement) out << "+0x" << std::hex << displacement;
    }
    IMAGEHLP_LINE64 line{};
    line.SizeOfStruct = sizeof(line);
    DWORD line_displacement = 0;
    if (SymGetLineFromAddr64(process, address, &line_displacement, &line)) {
        out << " [" << line.FileName << ':' << std::dec << line.LineNumber << ']';
    }
    return out.str();
}

static void capture_stack(std::ofstream& log, HANDLE process, DWORD thread_id, const EXCEPTION_DEBUG_INFO& info) {
    HANDLE thread = OpenThread(THREAD_GET_CONTEXT | THREAD_QUERY_INFORMATION, FALSE, thread_id);
    if (!thread) {
        log_line(log, "STACK_ERROR OpenThread=" + std::to_string(GetLastError()));
        return;
    }
    CONTEXT context{};
    context.ContextFlags = CONTEXT_FULL;
    if (!GetThreadContext(thread, &context)) {
        log_line(log, "STACK_ERROR GetThreadContext=" + std::to_string(GetLastError()));
        CloseHandle(thread);
        return;
    }

    std::ostringstream header;
    header << "CRASH exception=0x" << std::hex << info.ExceptionRecord.ExceptionCode
           << " address=0x" << reinterpret_cast<uintptr_t>(info.ExceptionRecord.ExceptionAddress)
           << " firstChance=" << std::dec << info.dwFirstChance
           << " thread=" << thread_id;
    log_line(log, header.str());

#if defined(_M_X64)
    STACKFRAME64 frame{};
    frame.AddrPC.Offset = context.Rip;
    frame.AddrPC.Mode = AddrModeFlat;
    frame.AddrStack.Offset = context.Rsp;
    frame.AddrStack.Mode = AddrModeFlat;
    frame.AddrFrame.Offset = context.Rbp;
    frame.AddrFrame.Mode = AddrModeFlat;
    for (int index = 0; index < 80; ++index) {
        if (frame.AddrPC.Offset == 0) break;
        log_line(log, "FRAME " + std::to_string(index) + " " + symbol_for(process, frame.AddrPC.Offset));
        if (!StackWalk64(
                IMAGE_FILE_MACHINE_AMD64,
                process,
                thread,
                &frame,
                &context,
                nullptr,
                SymFunctionTableAccess64,
                SymGetModuleBase64,
                nullptr)) {
            break;
        }
    }
#else
#error This helper must be built as x64.
#endif
    CloseHandle(thread);
}

int wmain(int argc, wchar_t** argv) {
    const std::wstring default_exe = L"C:\\Users\\pojian_liao\\Desktop\\devspace\\visualtex1.2.3-reference-3d207d7\\apps\\windows\\src-tauri\\target\\debug\\visualtex.exe";
    const std::wstring executable = argc > 1 ? argv[1] : default_exe;
    const auto log_path = temp_log_path();
    const auto pid_path = temp_pid_path();
    std::ofstream log(log_path, std::ios::binary | std::ios::trunc);
    if (!log) return 10;
    log_line(log, "CAPTURE_START exe=" + wide_to_utf8(executable));

    std::wstring command = L"\"" + executable + L"\" --office-background";
    std::vector<wchar_t> mutable_command(command.begin(), command.end());
    mutable_command.push_back(L'\0');
    STARTUPINFOW startup{};
    startup.cb = sizeof(startup);
    PROCESS_INFORMATION process_info{};
    const std::wstring working_directory = std::filesystem::path(executable).parent_path().wstring();
    DWORD flags = DEBUG_ONLY_THIS_PROCESS | CREATE_NO_WINDOW;
    if (!CreateProcessW(
            executable.c_str(),
            mutable_command.data(),
            nullptr,
            nullptr,
            FALSE,
            flags,
            nullptr,
            working_directory.c_str(),
            &startup,
            &process_info)) {
        log_line(log, "CREATE_PROCESS_FAILED error=" + std::to_string(GetLastError()));
        return 11;
    }
    {
        std::ofstream pid_file(pid_path, std::ios::binary | std::ios::trunc);
        pid_file << process_info.dwProcessId << '\n';
    }
    CloseHandle(process_info.hThread);
    log_line(log, "CHILD_STARTED pid=" + std::to_string(process_info.dwProcessId));

    HANDLE process = process_info.hProcess;
    SymSetOptions(SYMOPT_UNDNAME | SYMOPT_DEFERRED_LOADS | SYMOPT_LOAD_LINES | SYMOPT_FAIL_CRITICAL_ERRORS);
    if (!SymInitialize(process, nullptr, TRUE)) {
        log_line(log, "SYM_INITIALIZE_FAILED error=" + std::to_string(GetLastError()));
    } else {
        log_line(log, "SYM_INITIALIZED");
    }

    bool done = false;
    while (!done) {
        DEBUG_EVENT event{};
        if (!WaitForDebugEvent(&event, 1000)) {
            if (GetLastError() == ERROR_SEM_TIMEOUT) continue;
            log_line(log, "WAIT_DEBUG_FAILED error=" + std::to_string(GetLastError()));
            break;
        }
        DWORD continue_status = DBG_CONTINUE;
        switch (event.dwDebugEventCode) {
            case CREATE_PROCESS_DEBUG_EVENT:
                if (event.u.CreateProcessInfo.hFile) CloseHandle(event.u.CreateProcessInfo.hFile);
                break;
            case CREATE_THREAD_DEBUG_EVENT:
                if (event.u.CreateThread.hThread) CloseHandle(event.u.CreateThread.hThread);
                break;
            case LOAD_DLL_DEBUG_EVENT:
                if (event.u.LoadDll.hFile) CloseHandle(event.u.LoadDll.hFile);
                break;
            case EXCEPTION_DEBUG_EVENT: {
                const DWORD code = event.u.Exception.ExceptionRecord.ExceptionCode;
                if (code == EXCEPTION_BREAKPOINT || code == EXCEPTION_SINGLE_STEP) {
                    continue_status = DBG_CONTINUE;
                } else {
                    if (code == 0xC000001D) {
                        capture_stack(log, process, event.dwThreadId, event.u.Exception);
                    }
                    continue_status = DBG_EXCEPTION_NOT_HANDLED;
                }
                break;
            }
            case EXIT_PROCESS_DEBUG_EVENT:
                log_line(log, "CHILD_EXIT code=" + std::to_string(event.u.ExitProcess.dwExitCode));
                done = true;
                break;
            default:
                break;
        }
        ContinueDebugEvent(event.dwProcessId, event.dwThreadId, continue_status);
    }

    SymCleanup(process);
    CloseHandle(process);
    log_line(log, "CAPTURE_END");
    return 0;
}
