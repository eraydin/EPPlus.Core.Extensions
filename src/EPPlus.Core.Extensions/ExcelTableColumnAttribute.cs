using System;

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
                if (_columnIndex > 0)
                {
                    throw new ArgumentException($"Cannot set both {nameof(ColumnName)} and {nameof(ColumnIndex)}!");
                }

                if (string.IsNullOrWhiteSpace(value))
                {
                    throw new ArgumentException($"{nameof(ColumnName)} cannot be null or empty!");
                }

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
                if (_columnName != null)
                {
                    throw new ArgumentException($"Cannot set both {nameof(ColumnName)} and {nameof(ColumnIndex)}!");
                }
                if (value <= 0)
                {
                    throw new ArgumentException($"{nameof(ColumnIndex)} cannot be zero or negative!");
                }

                _columnIndex = value;
            }
        }
    }
}
