using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

using JetBrains.Annotations;

namespace EPPlus.Core.Extensions.Helpers
{
    internal static class Guard
    {
        [DebuggerStepThrough]
        public static T NotNull<T>([NotNull] T value, string parameterName)
        {
            if (value == null)
            {
                throw new ArgumentNullException(parameterName);
            }

            return value;
        }

        [DebuggerStepThrough]
        public static string NotNullOrEmpty([NotNull] string value, string name)
        {
            NotNull(value, name);

            if (value.Length == 0)
            {
                throw new ArgumentException("Value must not be empty", name);
            }

            return value;
        }

        [DebuggerStepThrough]
        public static IEnumerable<T> NotNullOrEmpty<T>([NotNull] IEnumerable<T> value, string name)
        {
            NotNull(value, name);

            if (!value.Any())
            {
                throw new ArgumentException("Value must not be empty", name);
            }

            return value;
        }

        [DebuggerStepThrough]
        public static string NotNullOrWhiteSpace([NotNull] string value, string name)
        {
            NotNullOrEmpty(value, name);

            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("Value must not be empty", name);
            }

            return value;
        }

        [DebuggerStepThrough]
        public static bool ThrowIfConditionMet(bool condition, string message, params object[] args)
        {
            if (condition)
            {
                throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture, message, args));
            }

            return true;
        }
    }
}