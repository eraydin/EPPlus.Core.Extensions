using System;

namespace EPPlus.Core.Extensions
{
    /// <inheritdoc />
    /// <summary>
    /// Attribute used to map property to Excel table column
    /// </summary>
    /// <remarks>Can not map by both name and index. It will map to the property name if none is specified</remarks>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelTableColumnAttribute : Attribute
    {
        private string columnName = null;

        /// <summary>
        /// Set this property to map by name
        /// </summary>
        public string ColumnName
        {
            get => columnName;
            set
            {
                if (columnIndex > 0) throw new ArgumentException("Can not set both Column Name and Column Index!");
                if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Column Name can't be empty!");

                columnName = value;
            }
        }

        private int columnIndex = 0;

        /// <summary>
        /// Use this property to map by 1-based index
        /// </summary>
        public int ColumnIndex
        {
            get => columnIndex;
            set
            {
                if (columnName != null) throw new ArgumentException("Can not set both Column Name and Column Index!");
                if (value <= 0) throw new ArgumentException("Column Index can't be zero or negative!");

                columnIndex = value;
            }
        }
    }
}
