using System.Reflection;
using System.Text.RegularExpressions;

namespace EPPlus.Core.Extensions
{
    internal class ExcelTableColumnAttributeAndProperyInfo
    {
        public ExcelTableColumnAttributeAndProperyInfo(PropertyInfo propertyInfo, ExcelTableColumnAttribute columnAttribute)
        {
            PropertyInfo = propertyInfo;
            ColumnAttribute = columnAttribute;
        }

        public PropertyInfo PropertyInfo { get; }

        public ExcelTableColumnAttribute ColumnAttribute { get; }

        public override string ToString()
        {
            return !string.IsNullOrEmpty(ColumnAttribute.ColumnName) ? ColumnAttribute.ColumnName : Regex.Replace(PropertyInfo.Name, "[a-z][A-Z]", m => $"{m.Value[0]} {m.Value[1]}");
        }
    }
}
