namespace VisualTeX.WindowsOleBridge;

internal static class ReplacementTransaction
{
    public static T Execute<T>(
        Func<T> createCandidate,
        Action<T> configureCandidate,
        Action deleteOriginal,
        Action<T> deleteCandidate)
        where T : class
    {
        T? candidate = null;
        try
        {
            candidate = createCandidate();
            configureCandidate(candidate);
            deleteOriginal();
            return candidate;
        }
        catch
        {
            if (candidate is not null)
            {
                try { deleteCandidate(candidate); } catch { }
            }
            throw;
        }
    }
}
