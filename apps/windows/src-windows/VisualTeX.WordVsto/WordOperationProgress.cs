using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace VisualTeX.WordVsto;

// Office mutations must remain on their STA without pumping nested Office
// events. This window owns a separate UI thread and receives only plain text.
internal sealed class WordOperationProgress : IDisposable
{
    // This is a status surface, not an input dialog. Activating an independent
    // UI thread here changes Windows' foreground history; closing it can then
    // surface the previously active application instead of the Word document.
    private sealed class ProgressForm : Form
    {
        protected override bool ShowWithoutActivation => true;
        protected override CreateParams CreateParams
        {
            get
            {
                var parameters = base.CreateParams;
                parameters.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE (including mouse clicks)
                return parameters;
            }
        }
    }

    private sealed class Snapshot
    {
        internal readonly string Message;
        internal readonly int Completed, Total;
        internal Snapshot(string message, int completed, int total)
        { Message = message; Completed = completed; Total = total; }
    }
    private Snapshot _snapshot = new("正在准备…", 0, 0);
    private readonly Stopwatch _elapsed = Stopwatch.StartNew();
    private readonly Action _onDispose;
    private int _closed;

    internal WordOperationProgress(string title, Action onDispose, IntPtr owner)
    {
        _onDispose = onDispose;
        Owner = new NativeWindowOwner(owner);
        var thread = new Thread(() => Run(title, owner)) { IsBackground = true, Name = "VisualTeX operation progress" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    internal void Report(string message, int completed = 0, int total = 0)
    {
        Volatile.Write(ref _snapshot, new Snapshot(message, completed, total));
        WordDoubleClickHook.TraceMessage($"operation-progress stage={message} completed={completed} total={total} elapsedMs={_elapsed.ElapsedMilliseconds}");
    }

    internal IWin32Window Owner { get; }

    private void Run(string title, IntPtr owner)
    {
        try
        {
            using var form = new ProgressForm
            {
                Text = title, ClientSize = new Size(500, 126),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false,
                ControlBox = false, ShowInTaskbar = false,
                Font = SystemFonts.MessageBoxFont,
            };
            var message = new Label { Left = 18, Top = 16, Width = 464, Height = 40 };
            var bar = new ProgressBar { Left = 18, Top = 62, Width = 464, Height = 18 };
            var elapsed = new Label { Left = 18, Top = 91, Width = 464, Height = 24 };
            form.Controls.AddRange(new Control[] { message, bar, elapsed });
            using var timer = new System.Windows.Forms.Timer { Interval = 150 };
            timer.Tick += (_, _) =>
            {
                if (Volatile.Read(ref _closed) != 0) { form.Close(); return; }
                var state = Volatile.Read(ref _snapshot);
                message.Text = state.Message;
                var determinate = state.Total > 0;
                bar.Style = determinate ? ProgressBarStyle.Continuous : ProgressBarStyle.Marquee;
                if (determinate)
                {
                    bar.Maximum = state.Total;
                    bar.Value = Math.Max(0, Math.Min(state.Completed, state.Total));
                }
                elapsed.Text = (determinate ? $"{state.Completed}/{state.Total}    " : "")
                    + $"已用时 {_elapsed.Elapsed:mm\\:ss}";
            };
            timer.Start();
            // A modeless owned window stays above its Word window without
            // covering unrelated applications or pumping the Office STA.
            if (Volatile.Read(ref _closed) != 0) return;
            if (owner != IntPtr.Zero) form.Show(Owner);
            Application.Run(form);
        }
        catch (Exception error)
        {
            // A progress-window failure must not abort or re-enter a document edit.
            WordDoubleClickHook.TraceMessage("operation-progress-window-failed " + error);
        }
    }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref _closed, 1) == 0) _onDispose();
    }
}

internal sealed partial class WordFormulaService
{
    internal Action<string, int, int>? OperationProgress { get; set; }
    private void ReportOperationProgress(string stage, int completed, int total) =>
        OperationProgress?.Invoke(stage, completed, total);
}

public sealed partial class ThisAddIn
{
    private WordOperationProgress? _operationProgress;
    private WordOperationProgress BeginOperationProgress(string title, WordFormulaService service,
        string initialStage = "正在读取公式范围…")
    {
        var window = _application!.ActiveWindow;
        IntPtr owner;
        try { owner = new IntPtr(window.Hwnd); }
        finally { ReleaseComObject(window); }
        WordOperationProgress? progress = null;
        progress = new WordOperationProgress(title, () =>
        {
            if (ReferenceEquals(_operationProgress, progress))
            {
                service.OperationProgress = null;
                _operationProgress = null;
            }
        }, owner);
        _operationProgress = progress;
        service.OperationProgress = progress.Report;
        progress.Report(initialStage);
        return progress;
    }
}
