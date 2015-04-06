using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Database
{
    [TestClass]
    public class CriteriaTests
    {
        [TestMethod]
        public void CriteriaShouldReadFieldsAndValues()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Crit1";
                sheet.Cells["B1"].Value = "Crit2";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = 2;

                var provider = new EpplusExcelDataProvider(package);

                var criteria = new ExcelDatabaseCriteria(provider, "A1:B2");

                Assert.AreEqual(2, criteria.Items.Count);
                Assert.AreEqual("Crit1", criteria.Items.Keys.First());
                Assert.AreEqual("Crit2", criteria.Items.Keys.Last());
                Assert.AreEqual(1, criteria.Items.Values.First());
                Assert.AreEqual(2, criteria.Items.Values.Last());
            }
        }

        [TestMethod]
        public void CriteriaShouldIgnoreEmptyFields1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Crit1";
                sheet.Cells["B1"].Value = "Crit2";
                sheet.Cells["A2"].Value = 1;

                var provider = new EpplusExcelDataProvider(package);

                var criteria = new ExcelDatabaseCriteria(provider, "A1:B2");

                Assert.AreEqual(1, criteria.Items.Count);
                Assert.AreEqual("Crit1", criteria.Items.Keys.First());
                Assert.AreEqual(1, criteria.Items.Values.Last());
            }
        }

        [TestMethod]
        public void CriteriaShouldIgnoreEmptyFields2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Crit1";
                sheet.Cells["A2"].Value = 1;

                var provider = new EpplusExcelDataProvider(package);

                var criteria = new ExcelDatabaseCriteria(provider, "A1:B2");

                Assert.AreEqual(1, criteria.Items.Count);
                Assert.AreEqual("Crit1", criteria.Items.Keys.First());
                Assert.AreEqual(1, criteria.Items.Values.Last());
            }
        }
    }
}
