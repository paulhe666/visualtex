using System.Windows.Forms;

namespace VisualTeX.PowerPointVsto;

internal sealed class OfficeUiDispatcher : IDisposable
{
    private readonly Control _control;
    private readonly HashSet<System.Windows.Forms.Timer> _delayedTimers = new();

    private static void RunFireAndForgetSafely(Action operation)
    {
        try { operation(); }
        catch (Exception error)
        {
            System.Diagnostics.Trace.WriteLine(
                $"VisualTeX PowerPoint UI callback failed: {error}");
        }
    }

    public OfficeUiDispatcher()
    {
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            throw new InvalidOperationException("The PowerPoint add-in must initialize on the Office STA thread.");
        _control = new Control();
        _control.CreateControl();
    }

    public void Post(Action operation)
    {
        if (_control.IsDisposed || _control.Disposing) return;
        try
        {
            _control.BeginInvoke(new Action(() => RunFireAndForgetSafely(operation)));
        }
        catch (InvalidOperationException) { }
    }

    public void PostDelayed(Action operation, int delayMilliseconds)
    {
        if (_control.IsDisposed || _control.Disposing) return;

        void Schedule()
        {
            if (_control.IsDisposed || _control.Disposing) return;
            var timer = new System.Windows.Forms.Timer
            {
                Interval = Math.Max(1, delayMilliseconds),
            };
            EventHandler? onTick = null;
            onTick = (_, _) =>
            {
                timer.Stop();
                if (onTick is not null) timer.Tick -= onTick;
                _delayedTimers.Remove(timer);
                timer.Dispose();
                if (_control.IsDisposed || _control.Disposing) return;
                RunFireAndForgetSafely(operation);
            };
            timer.Tick += onTick;
            _delayedTimers.Add(timer);
            timer.Start();
        }

        try
        {
            if (_control.InvokeRequired) _control.BeginInvoke(new Action(Schedule));
            else Schedule();
        }
        catch (InvalidOperationException) { }
    }

    public Task<T> InvokeAsync<T>(Func<T> operation)
    {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        void Execute()
        {
            try { completion.TrySetResult(operation()); }
            catch (Exception error) { completion.TrySetException(error); }
        }
        try
        {
            if (_control.IsDisposed || _control.Disposing)
                completion.TrySetException(new ObjectDisposedException(nameof(OfficeUiDispatcher)));
            else if (_control.InvokeRequired)
                _control.BeginInvoke(new Action(Execute));
            else
                Execute();
        }
        catch (InvalidOperationException error)
        {
            completion.TrySetException(error);
        }
        return completion.Task;
    }

    public void Dispose()
    {
        foreach (var timer in _delayedTimers.ToArray())
        {
            try { timer.Stop(); } catch { }
            timer.Dispose();
        }
        _delayedTimers.Clear();
        _control.Dispose();
    }
}
