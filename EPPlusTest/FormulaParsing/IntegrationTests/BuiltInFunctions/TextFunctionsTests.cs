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

        [TestMethod]
        public void CharShouldReturnCharValOfNumber()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Char(A2)";
                sheet.Cells["A2"].Value = 55;
                sheet.Calculate();
                Assert.AreEqual("7", sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void FixedShouldHaveCorrectDefaultValues()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Fixed(A2)";
                sheet.Cells["A2"].Value = 1234.5678;
                sheet.Calculate();
                Assert.AreEqual(1234.5678.ToString("N2"), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void FixedShouldSetCorrectNumberOfDecimals()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Fixed(A2,4)";
                sheet.Cells["A2"].Value = 1234.56789;
                sheet.Calculate();
                Assert.AreEqual(1234.56789.ToString("N4"), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void FixedShouldSetNoCommas()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Fixed(A2,4,true)";
                sheet.Cells["A2"].Value = 1234.56789;
                sheet.Calculate();
                Assert.AreEqual(1234.56789.ToString("F4"), sheet.Cells["A1"].Value);
            }
        }

        [TestMethod]
        public void FixedShouldHandleNegativeDecimals()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Fixed(A2,-1,true)";
                sheet.Cells["A2"].Value = 1234.56789;
                sheet.Calculate();
                Assert.AreEqual(1230.ToString("F0"), sheet.Cells["A1"].Value);
            }
        }
    }
}
