using System;
using System.Reflection;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions.Tests
{
    public class TestBase : IDisposable
    {
        protected readonly ExcelPackage excelPackage1;
        protected readonly ExcelPackage excelPackage2;
        private readonly string resourceName1 = "EPPlus.Core.Extensions.Tests.Resources.testsheets1.xlsx";
        private readonly string resourceName2 = "EPPlus.Core.Extensions.Tests.Resources.testsheets2.xlsx";

        public TestBase()
        {
            excelPackage1 = new ExcelPackage(typeof(TestBase).GetTypeInfo().Assembly.GetManifestResourceStream(resourceName1));
            excelPackage2 = new ExcelPackage(typeof(TestBase).GetTypeInfo().Assembly.GetManifestResourceStream(resourceName2));
        }

        public void Dispose()
        {
            excelPackage1.Dispose();
            excelPackage2.Dispose();
        }

        public string GetRandomName() => new Random(Guid.NewGuid().GetHashCode()).Next().ToString();
    }
}
