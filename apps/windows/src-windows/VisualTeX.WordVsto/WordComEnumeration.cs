using System.Collections;
using System.Runtime.InteropServices;

namespace VisualTeX.WordVsto;

// InlineShapes.Count and Item(i) both walk the collection, so an index loop over
// n inline shapes costs O(n^2) COM work in a long document; the collection's own
// enumerator visits them in one pass (measured 1.8 ms per shape at 100 and 500).
// Not for OMaths: its enumerator is far slower than index access.
// Each yielded element belongs to the caller, which releases it; the collection
// must not change while it is enumerated.
internal static class WordComEnumeration
{
    internal static IEnumerable<T> Each<T>(IEnumerable collection) where T : class
    {
        var enumerator = collection.GetEnumerator();
        try
        {
            while (enumerator.MoveNext())
            {
                var current = enumerator.Current;
                if (current is T item)
                    yield return item;
                else if (current is not null && Marshal.IsComObject(current))
                    Marshal.ReleaseComObject(current);
            }
        }
        finally
        {
            if (enumerator is ICustomAdapter adapter
                && adapter.GetUnderlyingObject() is { } native
                && Marshal.IsComObject(native))
                Marshal.ReleaseComObject(native);
        }
    }
}
