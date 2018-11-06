namespace EPPlus.Core.Extensions.Events
{
    public delegate void OnCaught<in T>(T current, int rowIndex);
}
