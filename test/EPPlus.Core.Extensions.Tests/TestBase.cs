using System;
using System.Reflection;

using OfficeOpenXml;

namespace EPPlus.Core.Extensions.Tests
{
    public class TestBase : IDisposable
    {
        protected readonly ExcelPackage ExcelPackage1;
        protected readonly ExcelPackage ExcelPackage2;
        private const string ResourceName1 = "EPPlus.Core.Extensions.Tests.Resources.testsheets1.xlsx";
        private const string ResourceName2 = "EPPlus.Core.Extensions.Tests.Resources.testsheets2.xlsx";

        protected TestBase()
        {
            ExcelPackage1 = new ExcelPackage(typeof(TestBase).GetTypeInfo().Assembly.GetManifestResourceStream(ResourceName1));
            ExcelPackage2 = new ExcelPackage(typeof(TestBase).GetTypeInfo().Assembly.GetManifestResourceStream(ResourceName2));
        }

        public void Dispose()
        {
            ExcelPackage1.Dispose();
            ExcelPackage2.Dispose();
        }

        protected string GetRandomName() => new Random(Guid.NewGuid().GetHashCode()).Next().ToString();
    }
}
