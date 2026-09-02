using System.Windows.Forms;

namespace VisualTeX.WordVsto;

internal sealed class OfficeUiDispatcher : IDisposable
{
    private readonly Control _control;

    private static void RunFireAndForgetSafely(Action operation)
    {
        try { operation(); }
        catch (Exception error)
        {
            System.Diagnostics.Trace.WriteLine(
                $"VisualTeX Word UI callback failed: {error}");
        }
    }

    public OfficeUiDispatcher()
    {
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            throw new InvalidOperationException("The Word add-in must initialize on the Office STA thread.");
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

    public Task<T> InvokeAsync<T>(Func<T> operation)
    {
        var completion = new TaskCompletionSource<T>(
            TaskCreationOptions.RunContinuationsAsynchronously);
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

    public void Dispose() => _control.Dispose();
}
