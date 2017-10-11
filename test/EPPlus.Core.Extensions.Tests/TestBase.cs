using System;
using System.Reflection;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions.Tests
{
    public class TestBase : IDisposable
    {
        public readonly ExcelPackage excelPackage;
        private readonly string resourceName = "EPPlus.Core.Extensions.Tests.Resources.testsheets.xlsx";

        public TestBase()
        {
            excelPackage = new ExcelPackage(typeof(TestBase).GetTypeInfo().Assembly.GetManifestResourceStream(resourceName));
        }

        public void Dispose()
        {
            excelPackage.Dispose();
        }
    }
}
