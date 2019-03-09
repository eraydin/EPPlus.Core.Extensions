using System.Reflection;

using EPPlus.Core.Extensions.Attributes;

namespace EPPlus.Core.Extensions.Enrichments
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
