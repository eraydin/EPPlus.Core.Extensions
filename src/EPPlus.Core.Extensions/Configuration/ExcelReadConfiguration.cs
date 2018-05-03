namespace EPPlus.Core.Extensions.Configuration
{
    public interface IExcelReadConfiguration<T>
    {
        bool HasHeaderRow { get; }       
     
        bool ThrowValidationExceptions { get; }

        bool ThrowCastingExceptions { get; }

        string CastingExceptionMessage { get; }     

        string ColumnValidationExceptionMessage { get; }

        OnCaught<T> OnCaught { get; }

        IExcelReadConfiguration<T> WithCastingExceptionMessage(string message);    

        IExcelReadConfiguration<T> WithHeaderValidationExceptionMessage(string message);

        IExcelReadConfiguration<T> SkipValidationErrors();

        IExcelReadConfiguration<T> SkipCastingErrors();

        IExcelReadConfiguration<T> WithHeaderRow();

        IExcelReadConfiguration<T> WithoutHeaderRow();

        IExcelReadConfiguration<T> Intercept(OnCaught<T> onCaught);
    }

    /// <summary>
    ///     Default configurations
    /// </summary>
    public class DefaultExcelReadConfiguration<T> : IExcelReadConfiguration<T>
    {
        public static IExcelReadConfiguration<T> Instance => new DefaultExcelReadConfiguration<T>();

        public string CastingExceptionMessage { get; protected set; } = $"The expected type of '{0}' property is '{1}', but the cell [{2}] contains an invalid value.";
              
        public string ColumnValidationExceptionMessage { get; protected set; } = $"{0} column could not found on the worksheet";

        public bool HasHeaderRow { get; protected set; } = true;

        public bool ThrowValidationExceptions { get; protected set; } = true;

        public bool ThrowCastingExceptions { get; protected set; } = true;

        public OnCaught<T> OnCaught { get; protected set; }

        public IExcelReadConfiguration<T> WithCastingExceptionMessage(string message)
        {
            CastingExceptionMessage = message;
            return this;
        }  

        public IExcelReadConfiguration<T> WithHeaderValidationExceptionMessage(string message)
        {
            ColumnValidationExceptionMessage = message;
            return this;
        }
      
        public IExcelReadConfiguration<T> SkipValidationErrors()
        {
            ThrowValidationExceptions = false;
            return this;
        }

        public IExcelReadConfiguration<T> SkipCastingErrors()
        {
            ThrowCastingExceptions = false;
            return this;
        }

        public IExcelReadConfiguration<T> WithHeaderRow()
        {
            HasHeaderRow = true;
            return this;
        }

        public IExcelReadConfiguration<T> WithoutHeaderRow()
        {
            HasHeaderRow = false;
            return this;
        }

        public IExcelReadConfiguration<T> Intercept(OnCaught<T> onCaught)
        {
            OnCaught = onCaught;
            return this;
        }
    }
}
