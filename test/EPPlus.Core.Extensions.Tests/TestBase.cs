using OfficeOpenXml;
using System.IO;
using System.Reflection;

namespace EPPlus.Core.Extensions.Tests
{
    public class TestBase
    {
        public static ExcelPackage excelPackage;
        private readonly string resourceName = "EPPlus.Core.Extensions.Tests.Resources.testsheets.xlsx";

        public TestBase()
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            {
                excelPackage = new ExcelPackage(stream);
            }
        }

        ~TestBase()
        {
            excelPackage.Dispose();
        }
    }
}
