using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace VisualTeX.WindowsOffice.VstoShared;

// Progress for long Office batch operations. The window lives on its own STA
// thread so it keeps painting (and its Stop button keeps working) while the
// Office thread is busy, and it never calls back into Office.
internal sealed class OfficeOperationProgress : IDisposable
{
    // A status surface, not an input dialog: it never takes activation, so
    // closing it cannot surface a different application. Its Stop button still
    // receives clicks.
    private sealed class ProgressForm : Form
    {
        protected override bool ShowWithoutActivation => true;
        protected override CreateParams CreateParams
        {
            get
            {
                var parameters = base.CreateParams;
                parameters.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE
                return parameters;
            }
        }
    }

    private sealed class OwnerWindow : IWin32Window
    {
        internal OwnerWindow(IntPtr handle) => Handle = handle;
        public IntPtr Handle { get; }
    }

    private sealed class Snapshot
    {
        internal Snapshot(string message, int completed, int total)
        {
            Message = message;
            Completed = completed;
            Total = total;
        }

        internal string Message { get; }
        internal int Completed { get; }
        internal int Total { get; }
    }

    private readonly Stopwatch _elapsed = Stopwatch.StartNew();
    private readonly CancellationTokenSource _cancellation = new();
    private readonly Action<string>? _trace;
    private readonly bool _cancellable;
    private Snapshot _snapshot;
    private string? _completionMessage;
    private int _lingerMilliseconds;
    private int _closed;

    internal OfficeOperationProgress(
        string title,
        IntPtr owner,
        bool cancellable,
        string initialMessage,
        Action<string>? trace = null)
    {
        _cancellable = cancellable;
        _trace = trace;
        _snapshot = new Snapshot(initialMessage, 0, 0);
        Owner = new OwnerWindow(owner);
        var thread = new Thread(() => Run(title, owner)) { IsBackground = true, Name = "VisualTeX operation progress" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    internal IWin32Window Owner { get; }

    // Cancelled by the Stop button. Operations check it between items, so the
    // item being written always completes.
    internal CancellationToken CancellationToken => _cancellation.Token;

    internal bool IsCancellationRequested => _cancellation.IsCancellationRequested;

    // Same effect as the Stop button.
    internal void RequestCancel()
    {
        try
        {
            if (!_cancellation.IsCancellationRequested) _cancellation.Cancel();
        }
        catch (ObjectDisposedException) { }
    }

    internal void Report(string message, int completed = 0, int total = 0)
    {
        Volatile.Write(ref _snapshot, new Snapshot(message, completed, total));
        _trace?.Invoke($"operation-progress stage={message} completed={completed} total={total} elapsedMs={_elapsed.ElapsedMilliseconds}");
    }

    // Shows the outcome in place and closes after a short delay, so a result is
    // visible even where Office has no usable status bar.
    internal void Complete(string message, int lingerMilliseconds = 1800)
    {
        Volatile.Write(ref _lingerMilliseconds, Math.Max(0, lingerMilliseconds));
        Volatile.Write(ref _completionMessage, message);
        _trace?.Invoke($"operation-progress-complete message={message} elapsedMs={_elapsed.ElapsedMilliseconds}");
    }

    // Closes the window; after Complete it closes once the outcome has lingered.
    public void Dispose()
    {
        Interlocked.Exchange(ref _closed, 1);
    }

    private void Run(string title, IntPtr owner)
    {
        try
        {
            var english = OfficePluginLanguage.IsEnglish;
            using var form = new ProgressForm
            {
                Text = title,
                ClientSize = new Size(520, _cancellable ? 150 : 126),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ControlBox = false,
                ShowInTaskbar = false,
                Font = SystemFonts.MessageBoxFont,
            };
            var message = new Label { Left = 18, Top = 14, Width = 484, Height = 42, AutoEllipsis = true };
            var bar = new ProgressBar { Left = 18, Top = 60, Width = 484, Height = 18 };
            var elapsed = new Label { Left = 18, Top = 88, Width = 330, Height = 24 };
            var stop = new Button
            {
                Text = english ? "Stop" : "停止",
                Left = 402,
                Top = 110,
                Width = 100,
                Height = 28,
                Visible = _cancellable,
            };
            stop.Click += (_, _) =>
            {
                if (_cancellation.IsCancellationRequested) return;
                _cancellation.Cancel();
                stop.Enabled = false;
                stop.Text = english ? "Stopping…" : "正在停止…";
                _trace?.Invoke($"operation-progress-stop-requested elapsedMs={_elapsed.ElapsedMilliseconds}");
            };
            form.Controls.AddRange(new Control[] { message, bar, elapsed, stop });

            var completionShownAt = -1L;
            using var timer = new System.Windows.Forms.Timer { Interval = 150 };
            timer.Tick += (_, _) =>
            {
                var completion = Volatile.Read(ref _completionMessage);
                // Dispose after Complete lets the outcome linger; otherwise close now.
                if (Volatile.Read(ref _closed) != 0 && completion is null) { form.Close(); return; }
                if (completion is not null)
                {
                    if (completionShownAt < 0)
                    {
                        completionShownAt = _elapsed.ElapsedMilliseconds;
                        message.Text = completion;
                        bar.Style = ProgressBarStyle.Continuous;
                        bar.Maximum = 1;
                        bar.Value = 1;
                        stop.Visible = false;
                        elapsed.Text = (english ? "Elapsed " : "已用时 ") + _elapsed.Elapsed.ToString(@"mm\:ss");
                    }
                    else if (_elapsed.ElapsedMilliseconds - completionShownAt >= Volatile.Read(ref _lingerMilliseconds))
                    {
                        form.Close();
                    }
                    return;
                }
                var state = Volatile.Read(ref _snapshot);
                message.Text = state.Message;
                var determinate = state.Total > 0;
                bar.Style = determinate ? ProgressBarStyle.Continuous : ProgressBarStyle.Marquee;
                if (determinate)
                {
                    bar.Maximum = state.Total;
                    bar.Value = Math.Max(0, Math.Min(state.Completed, state.Total));
                }
                elapsed.Text = (determinate ? $"{state.Completed}/{state.Total}    " : string.Empty)
                    + (english ? "Elapsed " : "已用时 ") + _elapsed.Elapsed.ToString(@"mm\:ss");
            };
            timer.Start();
            if (Volatile.Read(ref _closed) != 0) return;
            // A modeless owned window stays above its Office window without
            // covering unrelated applications or pumping the Office STA.
            if (owner != IntPtr.Zero) form.Show(Owner);
            else form.Show();
            Application.Run(form);
        }
        catch (Exception error)
        {
            // A progress-window failure must never abort or re-enter an Office edit.
            _trace?.Invoke("operation-progress-window-failed " + error);
        }
    }
}
