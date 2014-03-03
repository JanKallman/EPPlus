using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

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

        [TestMethod]
        public void CalculateTest()
        {
            //var pck = new ExcelPackage();
            //var ws = pck.Workbook.Worksheets.Add("Calc1");

            //ws.SetValue("A1", (short)1);
            //var v = pck.Workbook.FormulaParserManager.Parse("2.5-Calc1!A1+abs(3.0)-SIN(3)");
            //Assert.AreEqual(4.358879992, Math.Round((double)v, 9));

            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");

            ws.SetValue("A1", (short)1);
            var v = pck.Workbook.FormulaParserManager.Parse("2.5-Calc1!A1+ABS(-3.0)-SIN(3)*abs(5)");
            Assert.AreEqual(3.79439996, Math.Round((double)v,9));
        }

        [TestMethod]
        public void CalculateTest2()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");

            ws.SetValue("A1", (short)1);
            var v = pck.Workbook.FormulaParserManager.Parse("3*(2+5.5*2)+2*0.5+3");
            Assert.AreEqual(43, Math.Round((double)v, 9));
        }
    }
}
