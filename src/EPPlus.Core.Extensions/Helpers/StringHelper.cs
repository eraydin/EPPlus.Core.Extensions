using System;

namespace EPPlus.Core.Extensions.Helpers
{
    internal static class StringHelper
    {
        internal static string GenerateRandomTableName()
        {
            return $"Table{new Random(Guid.NewGuid().GetHashCode()).Next(99999)}";
        }
    }
}