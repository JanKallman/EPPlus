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

        //[Ignore]
        [TestMethod]
        public void VisibilityComments()
        {
            var xlsxName = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
            try
            {
                using (var ms = File.Open(xlsxName, FileMode.OpenOrCreate))
                using (var pkg = new ExcelPackage(ms))
                {
                    var ws = pkg.Workbook.Worksheets.Add("Comment");
                    var a1 = ws.Cells["A1"];
                    a1.Value = "Justin Dearing";
                    a1.AddComment("I am A1s comment", "JD");
                    Assert.IsFalse(a1.Comment.Visible); // Comments are by default invisible 
                    a1.Comment.Visible = true;
                    a1.Comment.Visible = false;
                    Assert.IsNotNull(a1.Comment);
                    //check style attribute
                    var stylesDict = new System.Collections.Generic.Dictionary<string, string>();
                    string[] styles = a1.Comment.Style
                        .Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                    foreach(var s in styles)
                    {
                        string[] split = s.Split(':');
                        if (split.Length == 2)
                        {
                            var k = (split[0] ?? "").Trim().ToLower();
                            var v = (split[1] ?? "").Trim().ToLower();
                            stylesDict[k] = v;
                        }
                    }
                    Assert.IsTrue(stylesDict.ContainsKey("visibility"));
                    //Assert.AreEqual("visible", stylesDict["visibility"]);
                    Assert.AreEqual("hidden", stylesDict["visibility"]);
                    Assert.IsFalse(a1.Comment.Visible);
                    pkg.Save();
                    ms.Close();
                }
            }
            finally
            {
                //open results file in program for view xlsx.
                //comments of cell A1 must be hidden.
                //System.Diagnostics.Process.Start(Path.GetDirectoryName(xlsxName));
                File.Delete(xlsxName);
            }
        }
    }
}
