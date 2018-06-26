using System.Reflection;

namespace EPPlus.Core.Extensions
{
    internal class ExcelTableColumnAttributeAndPropertyInfo
    {
        public ExcelTableColumnAttributeAndPropertyInfo(PropertyInfo propertyInfo, ExcelTableColumnAttribute columnAttribute)
        {
            PropertyInfo = propertyInfo;
            ColumnAttribute = columnAttribute;
        }

        public PropertyInfo PropertyInfo { get; }

        public ExcelTableColumnAttribute ColumnAttribute { get; }

        public override string ToString() => !string.IsNullOrEmpty(ColumnAttribute.ColumnName) ? ColumnAttribute.ColumnName : PropertyInfo.Name;
    }
}
