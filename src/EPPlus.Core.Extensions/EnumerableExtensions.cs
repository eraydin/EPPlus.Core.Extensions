using System.Collections.Generic;

namespace EPPlus.Core.Extensions
{
    internal static class EnumerableExtensions
    {
        internal static bool IsGreaterThanOne<T>(this IEnumerable<T> source)
        {
            using (IEnumerator<T> enumerator = source.GetEnumerator())
            {
                return enumerator.MoveNext() && enumerator.MoveNext();
            }
        }
    }
}
