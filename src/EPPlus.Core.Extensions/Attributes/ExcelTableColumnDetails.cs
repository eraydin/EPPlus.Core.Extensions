using System.Reflection;

namespace EPPlus.Core.Extensions.Attributes
{
    internal class ExcelTableColumnDetails
    {
        public ExcelTableColumnDetails(int columnPosition, PropertyInfo propertyInfo, ExcelTableColumnAttribute columnAttribute)
        {
            ColumnPosition = columnPosition;
            PropertyInfo = propertyInfo;
            ColumnAttribute = columnAttribute;
        }

        public PropertyInfo PropertyInfo { get; }

        public ExcelTableColumnAttribute ColumnAttribute { get; }

        public int ColumnPosition { get; }

        public override string ToString() => !string.IsNullOrEmpty(ColumnAttribute.ColumnName) ? ColumnAttribute.ColumnName : PropertyInfo.Name;
    }
}