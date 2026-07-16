using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Core.Extensions.Results
{
    /// <summary>
    ///     Contains imported items and any errors captured while mapping them.
    /// </summary>
    public sealed class ExcelReadResult<T>
    {
        internal ExcelReadResult(IEnumerable<T> items, IEnumerable<ExcelReadError> errors)
        {
            Items = items.ToList().AsReadOnly();
            Errors = errors.ToList().AsReadOnly();
        }

        public IReadOnlyList<T> Items { get; }

        public IReadOnlyList<ExcelReadError> Errors { get; }

        public bool HasErrors => Errors.Count > 0;

        public bool IsSuccess => !HasErrors;
    }
}
