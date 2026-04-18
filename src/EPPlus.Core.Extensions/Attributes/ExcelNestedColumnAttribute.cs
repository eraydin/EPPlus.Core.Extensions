using System;

namespace EPPlus.Core.Extensions.Attributes
{
    /// <summary>
    ///     Marks a complex-type property so that its own <see cref="ExcelTableColumnAttribute"/>-decorated
    ///     properties are mapped as flat columns on the parent worksheet row.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelNestedColumnAttribute : Attribute { }
}
