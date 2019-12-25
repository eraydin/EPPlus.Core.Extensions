using System.Reflection;

namespace EPPlus.Core.Extensions.Enrichments
{
    internal class ColumnPositionAndPropertyInfoAndIsOptional
    {
        public ColumnPositionAndPropertyInfoAndIsOptional(int columnPosition, PropertyInfo propertyInfo, bool isOptional)
        {
            ColumnPosition = columnPosition;
            PropertyInfo = propertyInfo;
            IsOptional = isOptional;
        }

        public int ColumnPosition { get; }

        public PropertyInfo PropertyInfo { get; }

        public bool IsOptional { get; }
    }
}