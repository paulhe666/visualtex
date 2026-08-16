using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const uint ThreadSuspendResume = 0x0002;
    private const uint ThreadGetContext = 0x0008;
    private const uint ThreadQueryInformation = 0x0040;
    private const uint ContextAmd64 = 0x00100000;
    private const uint ContextControl = ContextAmd64 | 0x00000001;
    private const int ContextFlagsOffset = 48;
    private const int RipOffset = 248;
    private const int ContextBufferSize = 1232;
    private const int RspOffset = 152;
    private const int RbpOffset = 160;
    private const uint ImageFileMachineAmd64 = 0x8664;
    private const int AddrModeFlat = 3;

    private static void RunOleClipboardFlushProbe()
    {
        var initialized = OleInitialize(IntPtr.Zero);
        if (initialized < 0)
            Marshal.ThrowExceptionForHR(initialized);
        try
        {
            var result = OleFlushClipboard();
            Console.WriteLine($"OleFlushClipboard HRESULT=0x{result:X8}");
            if (result < 0)
                Marshal.ThrowExceptionForHR(result);
        }
        finally
        {
            OleUninitialize();
        }
    }

    private static void RunWin32ThreadExecutionProbe()
    {
        if (!int.TryParse(Environment.GetEnvironmentVariable("VISUALTEX_THREAD_PROBE_PID"), out var processId)
            || processId <= 0)
            throw new InvalidOperationException("VISUALTEX_THREAD_PROBE_PID is required.");
        if (!int.TryParse(Environment.GetEnvironmentVariable("VISUALTEX_THREAD_PROBE_TID"), out var threadId)
            || threadId <= 0)
            throw new InvalidOperationException("VISUALTEX_THREAD_PROBE_TID is required.");

        using var process = Process.GetProcessById(processId);
        var modules = process.Modules
            .Cast<ProcessModule>()
            .Select(module => new
            {
                module.ModuleName,
                module.FileName,
                Start = unchecked((ulong)module.BaseAddress.ToInt64()),
                End = unchecked((ulong)module.BaseAddress.ToInt64()) + unchecked((ulong)module.ModuleMemorySize),
            })
            .OrderBy(module => module.Start)
            .ToArray();

        var thread = OpenThread(
            ThreadSuspendResume | ThreadGetContext | ThreadQueryInformation,
            false,
            unchecked((uint)threadId));
        if (thread == IntPtr.Zero)
            throw new InvalidOperationException($"OpenThread({threadId}) failed: {Marshal.GetLastWin32Error()}.");

        var context = Marshal.AllocHGlobal(ContextBufferSize + 16);
        try
        {
            var alignedAddress = (context.ToInt64() + 15L) & ~15L;
            var aligned = new IntPtr(alignedAddress);
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var sample = 1; sample <= 60; sample++)
            {
                var suspend = SuspendThread(thread);
                if (suspend == uint.MaxValue)
                    throw new InvalidOperationException($"SuspendThread({threadId}) failed: {Marshal.GetLastWin32Error()}.");
                ulong rip;
                try
                {
                    for (var offset = 0; offset < ContextBufferSize; offset += sizeof(long))
                        Marshal.WriteInt64(aligned, offset, 0L);
                    Marshal.WriteInt32(aligned, ContextFlagsOffset, unchecked((int)ContextControl));
                    if (!GetThreadContext(thread, aligned))
                        throw new InvalidOperationException($"GetThreadContext({threadId}) failed: {Marshal.GetLastWin32Error()}.");
                    rip = unchecked((ulong)Marshal.ReadInt64(aligned, RipOffset));
                }
                finally
                {
                    ResumeThread(thread);
                }

                var module = modules.FirstOrDefault(item => rip >= item.Start && rip < item.End);
                var label = module is null
                    ? "<unmapped>"
                    : module.ModuleName;
                counts[label] = counts.TryGetValue(label, out var count) ? count + 1 : 1;
                var offsetText = module is null ? string.Empty : $"+0x{rip - module.Start:X}";
                Console.WriteLine($"sample={sample:D2} tid={threadId} rip=0x{rip:X16} module={label}{offsetText}");
                Thread.Sleep(20);
            }

            Console.WriteLine("--- RIP module summary ---");
            foreach (var pair in counts.OrderByDescending(pair => pair.Value))
                Console.WriteLine($"{pair.Key}: {pair.Value}/60");

            Console.WriteLine("--- native stack snapshot ---");
            CaptureNativeStack(process.Handle, thread, aligned, threadId, modules);
        }
        finally
        {
            Marshal.FreeHGlobal(context);
            CloseHandle(thread);
        }
    }

    private static void CaptureNativeStack(
        IntPtr processHandle,
        IntPtr threadHandle,
        IntPtr context,
        int threadId,
        dynamic[] modules)
    {
        var suspend = SuspendThread(threadHandle);
        if (suspend == uint.MaxValue)
            throw new InvalidOperationException($"SuspendThread({threadId}) failed: {Marshal.GetLastWin32Error()}.");
        try
        {
            for (var offset = 0; offset < ContextBufferSize; offset += sizeof(long))
                Marshal.WriteInt64(context, offset, 0L);
            Marshal.WriteInt32(context, ContextFlagsOffset, unchecked((int)ContextControl));
            if (!GetThreadContext(threadHandle, context))
                throw new InvalidOperationException($"GetThreadContext({threadId}) failed: {Marshal.GetLastWin32Error()}.");

            if (!SymInitialize(processHandle, null, true))
                throw new InvalidOperationException($"SymInitialize failed: {Marshal.GetLastWin32Error()}.");
            try
            {
                var dbghelp = GetModuleHandleW("dbghelp.dll");
                if (dbghelp == IntPtr.Zero)
                    throw new InvalidOperationException("dbghelp.dll is not loaded after SymInitialize.");
                var functionTableAccess = GetProcAddress(dbghelp, "SymFunctionTableAccess64");
                var getModuleBase = GetProcAddress(dbghelp, "SymGetModuleBase64");
                if (functionTableAccess == IntPtr.Zero || getModuleBase == IntPtr.Zero)
                    throw new InvalidOperationException("DbgHelp stack-walk callbacks are unavailable.");

                var rip = unchecked((ulong)Marshal.ReadInt64(context, RipOffset));
                var rsp = unchecked((ulong)Marshal.ReadInt64(context, RspOffset));
                var rbp = unchecked((ulong)Marshal.ReadInt64(context, RbpOffset));
                var frame = new StackFrame64
                {
                    AddrPC = new Address64 { Offset = rip, Mode = AddrModeFlat },
                    AddrFrame = new Address64 { Offset = rbp, Mode = AddrModeFlat },
                    AddrStack = new Address64 { Offset = rsp, Mode = AddrModeFlat },
                    AddrReturn = new Address64 { Mode = AddrModeFlat },
                    AddrBStore = new Address64 { Mode = AddrModeFlat },
                };

                for (var index = 0; index < 32; index++)
                {
                    var address = index == 0 ? frame.AddrPC.Offset : 0UL;
                    if (index > 0)
                    {
                        if (!StackWalk64(
                                ImageFileMachineAmd64,
                                processHandle,
                                threadHandle,
                                ref frame,
                                context,
                                IntPtr.Zero,
                                functionTableAccess,
                                getModuleBase,
                                IntPtr.Zero))
                            break;
                        address = frame.AddrPC.Offset;
                    }
                    if (address == 0) break;
                    var module = modules.FirstOrDefault(item =>
                        address >= (ulong)item.Start && address < (ulong)item.End);
                    var label = module is null ? "<unmapped>" : (string)module.ModuleName;
                    var offsetText = module is null
                        ? string.Empty
                        : $"+0x{address - (ulong)module.Start:X}";
                    Console.WriteLine($"frame={index:D2} pc=0x{address:X16} module={label}{offsetText}");
                }
            }
            finally
            {
                SymCleanup(processHandle);
            }
        }
        finally
        {
            ResumeThread(threadHandle);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Address64
    {
        public ulong Offset;
        public ushort Segment;
        public int Mode;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KdHelp64
    {
        public ulong Thread;
        public uint ThCallbackStack;
        public uint ThCallbackBStore;
        public uint NextCallback;
        public uint FramePointer;
        public ulong KiCallUserMode;
        public ulong KeUserCallbackDispatcher;
        public ulong SystemRangeStart;
        public ulong KiUserExceptionDispatcher;
        public ulong StackBase;
        public ulong StackLimit;
        public uint BuildVersion;
        public uint RetpolineStubFunctionTableSize;
        public ulong RetpolineStubFunctionTable;
        public uint RetpolineStubOffset;
        public uint RetpolineStubSize;
        public ulong Reserved0A;
        public ulong Reserved0B;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct StackFrame64
    {
        public Address64 AddrPC;
        public Address64 AddrReturn;
        public Address64 AddrFrame;
        public Address64 AddrStack;
        public Address64 AddrBStore;
        public IntPtr FuncTableEntry;
        public ulong Params0;
        public ulong Params1;
        public ulong Params2;
        public ulong Params3;
        [MarshalAs(UnmanagedType.Bool)] public bool Far;
        [MarshalAs(UnmanagedType.Bool)] public bool Virtual;
        public ulong Reserved0;
        public ulong Reserved1;
        public ulong Reserved2;
        public KdHelp64 KdHelp;
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr OpenThread(uint desiredAccess, bool inheritHandle, uint threadId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint SuspendThread(IntPtr threadHandle);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint ResumeThread(IntPtr threadHandle);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetThreadContext(IntPtr threadHandle, IntPtr context);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr handle);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr GetModuleHandleW(string moduleName);

    [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
    private static extern IntPtr GetProcAddress(IntPtr module, string procName);

    [DllImport("dbghelp.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SymInitialize(IntPtr processHandle, string? searchPath, bool invadeProcess);

    [DllImport("dbghelp.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SymCleanup(IntPtr processHandle);

    [DllImport("ole32.dll")]
    private static extern int OleInitialize(IntPtr reserved);

    [DllImport("ole32.dll")]
    private static extern void OleUninitialize();

    [DllImport("ole32.dll")]
    private static extern int OleFlushClipboard();

    [DllImport("dbghelp.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool StackWalk64(
        uint machineType,
        IntPtr processHandle,
        IntPtr threadHandle,
        ref StackFrame64 stackFrame,
        IntPtr contextRecord,
        IntPtr readMemoryRoutine,
        IntPtr functionTableAccessRoutine,
        IntPtr getModuleBaseRoutine,
        IntPtr translateAddress);
}
