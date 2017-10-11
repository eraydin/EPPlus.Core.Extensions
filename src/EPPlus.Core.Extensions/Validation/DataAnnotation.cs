namespace EPPlus.Core.Extensions.Validation
{
    internal class DataAnnotation
    {
        public static EntityValidationResult ValidateEntity<T>(T entity) where T : class
        {
            return new EntityValidator<T>().Validate(entity);
        }
    }
}