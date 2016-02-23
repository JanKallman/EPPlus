using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class CommentsTest
    {
        [Ignore]
        [TestMethod]
        public void ReadExcelComments()
        {
            var fi = new FileInfo(@"c:\temp\googleComments\Comments.excel.xlsx");
            using (var excelPackage = new ExcelPackage(fi))
            {
                var sheet1 = excelPackage.Workbook.Worksheets.First();
                Assert.AreEqual(2, sheet1.Comments.Count);
            }
        }
        [Ignore]
        [TestMethod]
        public void ReadGoogleComments()
        {
            var fi = new FileInfo(@"c:\temp\googleComments\Comments.google.xlsx");
            using (var excelPackage = new ExcelPackage(fi))
            {
                var sheet1 = excelPackage.Workbook.Worksheets.First();
                Assert.AreEqual(2, sheet1.Comments.Count);
                Assert.AreEqual("Note for column 'Address'.", sheet1.Comments[0].Text);
            }
        }
    }
}
