using System.Collections.Generic;
using System.Diagnostics;

namespace EPPlus.Core.Extensions.Helpers
{
    internal static class EnumerableExtensions
    {
        [DebuggerStepThrough]
        internal static bool IsGreaterThanOne<T>(this IEnumerable<T> source)
        {
            using (IEnumerator<T> enumerator = source.GetEnumerator())
            {
                return enumerator.MoveNext() && enumerator.MoveNext();
            }
        }
    }
}