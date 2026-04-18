using System.Reflection;

namespace EPPlus.Core.Extensions.Attributes
{
    internal class ExcelTableColumnDetails
    {
        public ExcelTableColumnDetails(int columnPosition, PropertyInfo propertyInfo, ExcelTableColumnAttribute columnAttribute, PropertyInfo ownerPropertyInfo = null)
        {
            ColumnPosition = columnPosition;
            PropertyInfo = propertyInfo;
            ColumnAttribute = columnAttribute;
            OwnerPropertyInfo = ownerPropertyInfo;
        }

        public PropertyInfo PropertyInfo { get; }

        public ExcelTableColumnAttribute ColumnAttribute { get; }

        public int ColumnPosition { get; }

        /// <summary>
        ///     When non-null, this column belongs to a nested object reached via this property on the root type.
        /// </summary>
        public PropertyInfo OwnerPropertyInfo { get; }

        public override string ToString() => !string.IsNullOrEmpty(ColumnAttribute.ColumnName) ? ColumnAttribute.ColumnName : PropertyInfo.Name;
    }
}