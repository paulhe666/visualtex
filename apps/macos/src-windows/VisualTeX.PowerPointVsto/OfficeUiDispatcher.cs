using System.Windows.Forms;

namespace VisualTeX.PowerPointVsto;

internal sealed class OfficeUiDispatcher : IDisposable
{
    private readonly Control _control;

    public OfficeUiDispatcher()
    {
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            throw new InvalidOperationException("The PowerPoint add-in must initialize on the Office STA thread.");
        _control = new Control();
        _control.CreateControl();
    }

    public Task<T> InvokeAsync<T>(Func<T> operation)
    {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        void Execute()
        {
            try { completion.TrySetResult(operation()); }
            catch (Exception error) { completion.TrySetException(error); }
        }
        if (_control.InvokeRequired) _control.BeginInvoke(new Action(Execute));
        else Execute();
        return completion.Task;
    }

    public void Dispose() => _control.Dispose();
}
