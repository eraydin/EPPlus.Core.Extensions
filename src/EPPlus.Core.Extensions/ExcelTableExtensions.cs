
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.Core.Extensions
{
    public static class ExcelTableExtensions
    {
        /// <summary>
        /// Method returns table data bounds with regards to header and totals row visibility
        /// </summary>
        /// <param name="table">Extended object</param>
        /// <returns>Address range</returns>
        public static ExcelAddress GetDataBounds(this ExcelTable table)
        {
            return new ExcelAddress(
                table.Address.Start.Row + (table.ShowHeader ? 1 : 0),
                table.Address.Start.Column,
                table.Address.End.Row - (table.ShowTotal ? 1 : 0),
                table.Address.End.Column
            );
        }
    }
}
