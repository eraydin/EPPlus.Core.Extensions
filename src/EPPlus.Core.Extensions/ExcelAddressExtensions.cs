using OfficeOpenXml;

namespace EPPlus.Core.Extensions
{
    public static class ExcelAddressExtensions
    {
        /// <summary>
        ///     Checks whether the given range is empty or not
        /// </summary>
        /// <param name="address">Excel cell range</param>
        /// <param name="hasHeaderRow">'false' as default</param>
        /// <returns>'true' or 'false'</returns>
        public static bool IsEmptyRange(this ExcelAddressBase address, bool hasHeaderRow = false) => !hasHeaderRow ? address.Start.Row == 0 : address.Start.Row == address.End.Row;
    }
}
