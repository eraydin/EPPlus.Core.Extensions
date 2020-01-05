using EPPlus.Core.Extensions.Events;

namespace EPPlus.Core.Extensions.Configuration
{
    public class ExcelReadConfiguration<T>
    {
        public static ExcelReadConfiguration<T> Instance => new ExcelReadConfiguration<T>();

        public virtual string CastingExceptionMessage { get; private set; } = "The expected type of '{0}' property is '{1}', but the cell [{2}] contains an invalid value.";

        public virtual string ColumnValidationExceptionMessage { get; private set; } = "'{0}' column could not found on the worksheet.";

        public virtual bool HasHeaderRow { get; private set; } = true;

        public virtual bool ThrowValidationExceptions { get; private set; } = true;

        public virtual bool ThrowCastingExceptions { get; private set; } = true;

        public virtual OnCaught<T> OnCaught { get; private set; }

        public virtual ExcelReadConfiguration<T> WithCastingExceptionMessage(string message)
        {
            CastingExceptionMessage = message;
            return this;
        }

        public virtual ExcelReadConfiguration<T> WithHeaderValidationExceptionMessage(string message)
        {
            ColumnValidationExceptionMessage = message;
            return this;
        }

        public virtual ExcelReadConfiguration<T> SkipValidationErrors()
        {
            ThrowValidationExceptions = false;
            return this;
        }

        public virtual ExcelReadConfiguration<T> SkipCastingErrors()
        {
            ThrowCastingExceptions = false;
            return this;
        }  

        public virtual ExcelReadConfiguration<T> WithoutHeaderRow()
        {
            HasHeaderRow = false;
            return this;
        }

        public virtual ExcelReadConfiguration<T> Intercept(OnCaught<T> onCaught)
        {
            OnCaught = onCaught;
            return this;
        }
    }
}  