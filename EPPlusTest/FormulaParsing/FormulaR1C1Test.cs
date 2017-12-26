using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class FormulaR1C1Tests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;
        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            var s1 = _package.Workbook.Worksheets.Add("test");
            s1.Cells["A1"].Value = 1;
            s1.Cells["A2"].Value = 2;
            s1.Cells["A3"].Value = 3;
            s1.Cells["A4"].Value = 4;

            s1.Cells["B1"].Value = 5;
            s1.Cells["B2"].Value = 6;
            s1.Cells["B3"].Value = 7;
            s1.Cells["B4"].Value = 8;

            s1.Cells["C1"].Value = 5;
            s1.Cells["C2"].Value = 6;
            s1.Cells["C3"].Value = 7;
            s1.Cells["C4"].Value = 8;

            _sheet = s1;
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void RC2()
        {
            string fR1C1 = "RC2";
            _sheet.Cells[5, 1].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 1].Formula;
            _sheet.Cells[5, 1].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[5,1].FormulaR1C1);
        }
        [TestMethod]
        public void C()
        {
            string fR1C1 = "SUMIFS(C,C2,RC1)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            _sheet.Cells[5, 3].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
        }
        [TestMethod]
        public void C2Abs()
        {
            string fR1C1 = "SUM(C2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUM($B:$B)", f);
        }
        [TestMethod]
        public void C2AbsWithSheet()
        {
            string fR1C1 = "SUM(A!C2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUM(A!$B:$B)", f);
        }
        [TestMethod]
        public void C2()
        {
            string fR1C1 = "SUM(C2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            _sheet.Cells[5, 3].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
        }
        [TestMethod]
        public void R2Abs()
        {
            string fR1C1 = "SUM(R2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUM($2:$2)",f);
        }
        [TestMethod]
        public void R2()
        {
            string fR1C1 = "SUM(R2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            _sheet.Cells[5, 3].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
        }
        [TestMethod]
        public void RCRelativeToAB()
        {
            string fR1C1 = "SUMIFS(C,C2,RC1)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.AreEqual("SUMIFS(C:C,$B:$B,$A5)", f);
        }
        [TestMethod]
        public void RCRelativeToABToR1C1()
        {
            string fR1C1 = "SUMIFS(C,C2,RC1)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            _sheet.Cells[5, 3].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
        }
        [TestMethod]
        public void RCRelativeToABToR1C1_2()
        {
            string fR1C1 = "SUM(RC9:RC[-1])";
            _sheet.Cells[5, 13].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 13].Formula;
            Assert.AreEqual("SUM($I5:L5)", f);
            _sheet.Cells[5, 13].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[5, 13].FormulaR1C1);

            //"RC{colShort} - SUM(RC21:RC12)";
        }
        [TestMethod]
        public void RCFixToABToR1C1_2()
        {
            string fR1C1 = "RC28-SUM(RC12:RC21)";
            _sheet.Cells[6, 13].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[6, 13].Formula;
            Assert.AreEqual("$AB6-SUM($L6:$U6)", f);
            _sheet.Cells[6, 13].Formula = f;
            Assert.AreEqual(fR1C1, _sheet.Cells[6, 13].FormulaR1C1);

        }
    }
}
