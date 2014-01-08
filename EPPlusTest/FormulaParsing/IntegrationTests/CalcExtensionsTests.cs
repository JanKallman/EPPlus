using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Calculation;

namespace EPPlusTest.FormulaParsing.IntegrationTests
{
    [TestClass]
    public class CalcExtensionsTests
    {
        [TestMethod]
        public void ShouldCalculateChainTest()
        {
            var package = new ExcelPackage(new FileInfo("c:\\temp\\chaintest.xlsx"));
            package.Workbook.Calculate();
        }
    }
}
