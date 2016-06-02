using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Database
{
    [TestClass]
    public class ExcelDatabaseTests
    {
        [TestMethod]
        public void DatabaseShouldReadFields()
        {
            using (var package = new ExcelPackage())
            {
                var database = GetDatabase(package);

                Assert.AreEqual(2, database.Fields.Count(), "count was not 2");
                Assert.AreEqual("col1", database.Fields.First().FieldName, "first fieldname was not 'col1'");
                Assert.AreEqual("col2", database.Fields.Last().FieldName, "last fieldname was not 'col12'");
            }
        }

        [TestMethod]
        public void HasMoreRowsShouldBeTrueWhenInitialized()
        {
            using (var package = new ExcelPackage())
            {
                var database = GetDatabase(package);

                Assert.IsTrue(database.HasMoreRows);
            }
            
        }

        [TestMethod]
        public void HasMoreRowsShouldBeFalseWhenLastRowIsRead()
        {
            using (var package = new ExcelPackage())
            {
                var database = GetDatabase(package);
                database.Read();

                Assert.IsFalse(database.HasMoreRows);
            }

        }

        [TestMethod]
        public void DatabaseShouldReadFieldsInRow()
        {
            using (var package = new ExcelPackage())
            {
                var database = GetDatabase(package);
                var row = database.Read();

                Assert.AreEqual(1, row["col1"]);
                Assert.AreEqual(2, row["col2"]);
            }

        }

        private static ExcelDatabase GetDatabase(ExcelPackage package)
        {
            var provider = new EpplusExcelDataProvider(package);
            var sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Value = "col1";
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["B1"].Value = "col2";
            sheet.Cells["B2"].Value = 2;
            var database = new ExcelDatabase(provider, "A1:B2");
            return database;
        }
    }
}
