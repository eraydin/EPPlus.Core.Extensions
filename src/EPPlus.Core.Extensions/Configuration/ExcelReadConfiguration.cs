using EPPlus.Core.Extensions.Events;

namespace EPPlus.Core.Extensions.Configuration
{
    public abstract class ExcelReadConfiguration<T>
    {
        internal string CastingExceptionMessage { get; private set; } = "The expected type of '{0}' property is '{1}', but the cell [{2}] contains an invalid value.";

        internal string ColumnValidationExceptionMessage { get; private set; } = "'{0}' column could not found on the worksheet.";

        internal bool HasHeaderRow { get; private set; } = true;

        internal bool ThrowValidationExceptions { get; private set; } = true;

        internal bool ThrowCastingExceptions { get; private set; } = true;

        internal OnCaught<T> OnCaught { get; private set; }

        public ExcelReadConfiguration<T> WithCastingExceptionMessage(string message)
        {
            CastingExceptionMessage = message;
            return this;
        }

        public ExcelReadConfiguration<T> WithHeaderValidationExceptionMessage(string message)
        {
            ColumnValidationExceptionMessage = message;
            return this;
        }

        public ExcelReadConfiguration<T> SkipValidationErrors()
        {
            ThrowValidationExceptions = false;
            return this;
        }

        public ExcelReadConfiguration<T> SkipCastingErrors()
        {
            ThrowCastingExceptions = false;
            return this;
        }  

        public ExcelReadConfiguration<T> WithoutHeaderRow()
        {
            HasHeaderRow = false;
            return this;
        }

        public ExcelReadConfiguration<T> Intercept(OnCaught<T> onCaught)
        {
            OnCaught = onCaught;
            return this;
        }
    }

    public class DefaultExcelReadConfiguration<T> : ExcelReadConfiguration<T>
    {
        public static ExcelReadConfiguration<T> Instance => new DefaultExcelReadConfiguration<T>();
    }
}  