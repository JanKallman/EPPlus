using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class DatabaseTests
    {
        [TestMethod]
        public void DgetShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                // database
                sheet.Cells["A1"].Value = "crit1";
                sheet.Cells["B1"].Value = "crit2";
                sheet.Cells["C1"].Value = "crit3";
                sheet.Cells["A2"].Value = "test";
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["C2"].Value = "output";
                sheet.Cells["A3"].Value = "test";
                sheet.Cells["B3"].Value = 3;
                sheet.Cells["C3"].Value = "aaa";
                // criteria
                sheet.Cells["D1"].Value = "crit1";
                sheet.Cells["D2"].Value = "t*t";
                sheet.Cells["E1"].Value = "crit2";
                sheet.Cells["E2"].Value = 2;
                // function
                sheet.Cells["F1"].Formula = "DGET(A1:C3,\"Crit3\",D1:E2)";

                sheet.Workbook.Calculate();

                Assert.AreEqual("output", sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void DcountShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                // database
                sheet.Cells["A1"].Value = "crit1";
                sheet.Cells["B1"].Value = "crit2";
                sheet.Cells["C1"].Value = "crit3";
                sheet.Cells["A2"].Value = "test";
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["C2"].Value = "output";
                sheet.Cells["A3"].Value = "tesst";
                sheet.Cells["B3"].Value = "2";
                sheet.Cells["C3"].Value = "aaa";
                // criteria
                sheet.Cells["D1"].Value = "crit1";
                sheet.Cells["D2"].Value = "t*t";
                sheet.Cells["E1"].Value = "crit2";
                sheet.Cells["E2"].Value = 2;
                // function
                sheet.Cells["F1"].Formula = "DCOUNT(A1:C3,\"Crit2\",D1:E2)";

                sheet.Workbook.Calculate();

                Assert.AreEqual(1, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void DcountaShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                // database
                sheet.Cells["A1"].Value = "crit1";
                sheet.Cells["B1"].Value = "crit2";
                sheet.Cells["C1"].Value = "crit3";
                sheet.Cells["A2"].Value = "test";
                sheet.Cells["B2"].Value = 2;
                sheet.Cells["C2"].Value = "output";
                sheet.Cells["A3"].Value = "tesst";
                sheet.Cells["B3"].Value = "2";
                sheet.Cells["C3"].Value = "aaa";
                // criteria
                sheet.Cells["D1"].Value = "crit1";
                sheet.Cells["D2"].Value = "t*t";
                sheet.Cells["E1"].Value = "crit2";
                sheet.Cells["E2"].Value = 2;
                // function
                sheet.Cells["F1"].Formula = "DCOUNTA(A1:C3,\"Crit2\",D1:E2)";

                sheet.Workbook.Calculate();

                Assert.AreEqual(2, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void DMaxShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                // database
                sheet.Cells["A1"].Value = "crit1";
                sheet.Cells["B1"].Value = "crit2";

                sheet.Cells["A2"].Value = "test";
                sheet.Cells["B2"].Value = 2;

                sheet.Cells["A3"].Value = "tesst";
                sheet.Cells["B3"].Value = 1;
                // criteria
                sheet.Cells["D1"].Value = "crit1";
                sheet.Cells["D2"].Value = "t*t";

                // function
                sheet.Cells["F1"].Formula = "DMAX(A1:B3,\"Crit2\",D1:E2)";

                sheet.Workbook.Calculate();

                Assert.AreEqual(2d, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void DMinShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                // database
                sheet.Cells["A1"].Value = "crit1";
                sheet.Cells["B1"].Value = "crit2";

                sheet.Cells["A2"].Value = "test";
                sheet.Cells["B2"].Value = 2;

                sheet.Cells["A3"].Value = "tesst";
                sheet.Cells["B3"].Value = 1;
                // criteria
                sheet.Cells["D1"].Value = "crit1";
                sheet.Cells["D2"].Value = "t*t";

                // function
                sheet.Cells["F1"].Formula = "DMIN(A1:B3,\"Crit2\",D1:E2)";

                sheet.Workbook.Calculate();

                Assert.AreEqual(1d, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void DSumShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                // database
                sheet.Cells["A1"].Value = "crit1";
                sheet.Cells["B1"].Value = "crit2";

                sheet.Cells["A2"].Value = "test";
                sheet.Cells["B2"].Value = 2;

                sheet.Cells["A3"].Value = "tesst";
                sheet.Cells["B3"].Value = 1;
                // criteria
                sheet.Cells["D1"].Value = "crit1";
                sheet.Cells["D2"].Value = "t*t";

                // function
                sheet.Cells["F1"].Formula = "DSUM(A1:B3,\"Crit2\",D1:E2)";

                sheet.Workbook.Calculate();

                Assert.AreEqual(3d, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void DAverageShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                // database
                sheet.Cells["A1"].Value = "crit1";
                sheet.Cells["B1"].Value = "crit2";

                sheet.Cells["A2"].Value = "test";
                sheet.Cells["B2"].Value = 2;

                sheet.Cells["A3"].Value = "tesst";
                sheet.Cells["B3"].Value = 1;
                // criteria
                sheet.Cells["D1"].Value = "crit1";
                sheet.Cells["D2"].Value = "t*t";

                // function
                sheet.Cells["F1"].Formula = "DAVERAGE(A1:B3,\"Crit2\",D1:E2)";

                sheet.Workbook.Calculate();

                Assert.AreEqual(1.5d, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void DVarShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                // database
                sheet.Cells["A1"].Value = "crit1";
                sheet.Cells["B1"].Value = "crit2";

                sheet.Cells["A2"].Value = "test";
                sheet.Cells["B2"].Value = 2;

                sheet.Cells["A3"].Value = "tesst";
                sheet.Cells["B3"].Value = 1;
                // criteria
                sheet.Cells["D1"].Value = "crit1";
                sheet.Cells["D2"].Value = "t*t";

                // function
                sheet.Cells["F1"].Formula = "DVAR(A1:B3,\"Crit2\",D1:E2)";

                sheet.Workbook.Calculate();

                Assert.AreEqual(0.5d, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void DVarpShouldReturnCorrectResult()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                // database
                sheet.Cells["A1"].Value = "crit1";
                sheet.Cells["B1"].Value = "crit2";

                sheet.Cells["A2"].Value = "test";
                sheet.Cells["B2"].Value = 2;

                sheet.Cells["A3"].Value = "tesst";
                sheet.Cells["B3"].Value = 1;
                // criteria
                sheet.Cells["D1"].Value = "crit1";
                sheet.Cells["D2"].Value = "t*t";

                // function
                sheet.Cells["F1"].Formula = "DVARP(A1:B3,\"Crit2\",D1:E2)";

                sheet.Workbook.Calculate();

                Assert.AreEqual(0.25d, sheet.Cells["F1"].Value);
            }
        }

        [TestMethod]
        public void DVarpShouldReturnByFieldIndex()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                // database
                sheet.Cells["A1"].Value = "crit1";
                sheet.Cells["B1"].Value = "crit2";

                sheet.Cells["A2"].Value = "test";
                sheet.Cells["B2"].Value = 2;

                sheet.Cells["A3"].Value = "tesst";
                sheet.Cells["B3"].Value = 1;
                // criteria
                sheet.Cells["D1"].Value = "crit1";
                sheet.Cells["D2"].Value = "t*t";

                // function
                sheet.Cells["F1"].Formula = "DVARP(A1:B3,2,D1:E2)";

                sheet.Workbook.Calculate();

                Assert.AreEqual(0.25d, sheet.Cells["F1"].Value);
            }
        }
    }
}
