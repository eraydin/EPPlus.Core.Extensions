namespace EPPlus.Core.Extensions.Results
{
    /// <summary>
    ///     Describes the stage at which an Excel import error occurred.
    /// </summary>
    public enum ExcelReadErrorKind
    {
        Mapping,
        Casting,
        Validation
    }
}
