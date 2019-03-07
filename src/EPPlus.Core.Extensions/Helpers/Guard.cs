using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

namespace EPPlus.Core.Extensions.Helpers
{
    internal static class Guard
    {
        [DebuggerStepThrough]
        public static void NotNull<T>(T value, string parameterName)
        {
            if (value == null)
            {
                throw new ArgumentNullException(parameterName);
            }
        }

        [DebuggerStepThrough]
        public static void NotNullOrEmpty<T>(IEnumerable<T> value, string name)
        {
            NotNull(value, name);

            if (!value.Any())
            {
                throw new ArgumentException("Value must not be empty", name);
            }
        }

        [DebuggerStepThrough]
        public static void NotNullOrWhiteSpace(string value, string name)
        {
            NotNullOrEmpty(value, name);

            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("Value must not be empty", name);
            }
        }

        [DebuggerStepThrough]
        public static void ThrowIfConditionMet(bool condition, string message, params object[] args)
        {
            if (condition)
            {
                throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, message, args));
            }
        }

        [DebuggerStepThrough]
        private static void NotNullOrEmpty(string value, string name)
        {
            NotNull(value, name);

            if (value.Length == 0)
            {
                throw new ArgumentException("Value must not be empty", name);
            }
        }
    }
}