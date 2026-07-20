using System.Collections.Concurrent;
using System.Runtime.ExceptionServices;
using System.Windows.Forms;

namespace VisualTeX.WindowsOleBridge;

internal sealed class OfficeStaDispatcher : IDisposable
{
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _ready = new(false);
    private readonly BlockingCollection<Action> _queue = new();
    private Control? _control;
    private ApplicationContext? _context;
    private bool _disposed;

    public OfficeStaDispatcher(FileLogger logger)
    {
        _thread = new Thread(() => RunMessageLoop(logger))
        {
            Name = "VisualTeX Office COM STA",
            IsBackground = true,
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
        if (!_ready.Wait(TimeSpan.FromSeconds(10)))
            throw new TimeoutException("Office STA dispatcher did not start in time.");
    }

    public int ManagedThreadId => _thread.ManagedThreadId;

    public Task<T> InvokeAsync<T>(Func<T> operation, CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var completion = new TaskCompletionSource<T>(
            TaskCreationOptions.RunContinuationsAsynchronously);
        cancellationToken.ThrowIfCancellationRequested();
        cancellationToken.Register(() => completion.TrySetCanceled(cancellationToken));

        void Execute()
        {
            if (completion.Task.IsCompleted) return;
            try
            {
                if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    throw new InvalidOperationException("Office COM operation left the STA thread.");
                completion.TrySetResult(operation());
            }
            catch (Exception error)
            {
                completion.TrySetException(error);
            }
        }

        _queue.Add(Execute, cancellationToken);
        _control?.BeginInvoke(new Action(DrainQueue));
        return completion.Task;
    }

    public Task InvokeAsync(Action operation, CancellationToken cancellationToken) =>
        InvokeAsync(
            () =>
            {
                operation();
                return true;
            },
            cancellationToken);

    private void RunMessageLoop(FileLogger logger)
    {
        try
        {
            _context = new ApplicationContext();
            _control = new Control();
            _control.CreateControl();
            _ready.Set();
            using var timer = new System.Windows.Forms.Timer { Interval = 15 };
            timer.Tick += (_, _) => DrainQueue();
            timer.Start();
            Application.Run(_context);
        }
        catch (Exception error)
        {
            logger.Error("Office STA message loop failed.", error);
            _ready.Set();
            while (_queue.TryTake(out var work))
            {
                try { work(); } catch { }
            }
        }
        finally
        {
            _control?.Dispose();
            _control = null;
        }
    }

    private void DrainQueue()
    {
        while (_queue.TryTake(out var work)) work();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        if (_control is not null && !_control.IsDisposed)
        {
            try
            {
                _control.BeginInvoke(new Action(() => _context?.ExitThread()));
            }
            catch { }
        }
        if (!_thread.Join(TimeSpan.FromSeconds(3)))
        {
            // Background thread will terminate with the process if Office is stuck.
        }
        _queue.Dispose();
        _ready.Dispose();
    }
}
