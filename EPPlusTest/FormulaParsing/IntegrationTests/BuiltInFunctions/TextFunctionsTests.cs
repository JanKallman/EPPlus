using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class TextFunctionsTests
    {
        [TestMethod]
        public void HyperlinkShouldHandleReference()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "HYPERLINK(B1)";
                sheet.Cells["B1"].Value = "http://epplus.codeplex.com";
                sheet.Calculate();
                Assert.AreEqual("http://epplus.codeplex.com", sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void HyperlinkShouldHandleReference2()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "HYPERLINK(B1, B2)";
                sheet.Cells["B1"].Value = "http://epplus.codeplex.com";
                sheet.Cells["B2"].Value = "Epplus";
                sheet.Calculate();
                Assert.AreEqual("Epplus", sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void HyperlinkShouldHandleText()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "HYPERLINK(\"testing\")";
                sheet.Calculate();
                Assert.AreEqual("testing", sheet.Cells["A1"].Value);
            }
        }
    }
}
