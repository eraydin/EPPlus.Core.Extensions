using System;

using static EPPlus.Core.Extensions.Helpers.Guard;

namespace EPPlus.Core.Extensions
{
    /// <inheritdoc />
    /// <summary>
    ///     Attribute used to map property to Excel table column
    /// </summary>
    /// <remarks>Can not map by both name and index. It will map to the property name if none is specified</remarks>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelTableColumnAttribute : Attribute
    {
        private int _columnIndex;

        private string _columnName;

        public ExcelTableColumnAttribute()
        {
        }

        /// <inheritdoc />
        /// <summary>
        ///     Set this property to map by 1-based index
        /// </summary>
        public ExcelTableColumnAttribute(int columnIndex) => ColumnIndex = columnIndex;

        /// <inheritdoc />
        /// <summary>
        ///     Set this property to map by name
        /// </summary>
        public ExcelTableColumnAttribute(string columnName) => ColumnName = columnName;

        /// <summary>
        ///     Set this property to map by name
        /// </summary>
        public string ColumnName
        {
            get => _columnName;
            set
            {
                ThrowIfConditionMet(_columnIndex > 0, "Cannot set both {0} and {1}!", nameof(ColumnName), nameof(ColumnIndex));
                NotNullOrWhiteSpace(value, nameof(ColumnName));

                _columnName = value;
            }
        }

        /// <summary>
        ///     Use this property to map by 1-based index
        /// </summary>
        public int ColumnIndex
        {
            get => _columnIndex;
            set
            {
                ThrowIfConditionMet(_columnName != null, "Cannot set both {0} and {1}!", nameof(ColumnName), nameof(ColumnIndex));
                ThrowIfConditionMet(value <= 0, "{0} cannot be zero or negative!", nameof(ColumnIndex));
                _columnIndex = value;
            }
        }
    }
}