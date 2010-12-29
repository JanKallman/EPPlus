using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace ExcelPackageTest
{
    [TestClass]
    public class ReadTemplate
    {
        [TestMethod]
        public void ReadDrawing()
        {
            using (ExcelPackage pck = new ExcelPackage(new FileInfo(@"Test\Drawing.xlsx"))) 
            {
                var ws = pck.Workbook.Worksheets["Pyramid"];
                Assert.AreEqual(ws.Cells["V24"].Value, 104D);
            }

        }
        [TestMethod]
        public void ReadWorkSheet()
        {
            FileStream instream = new FileStream(@"Test\Worksheet.xlsx", FileMode.Open, FileAccess.ReadWrite);
            using (ExcelPackage pck = new ExcelPackage(instream))
            {
                var ws = pck.Workbook.Worksheets["Perf"];
                Assert.AreEqual(ws.Cells["H6"].Formula, "B5+B6");

                ws = pck.Workbook.Worksheets["Comment"];
                var comment = ws.Cells["B2"].Comment;

                Assert.AreNotEqual(comment, null);
                Assert.AreEqual(comment.Author, "JK");
            }
            instream.Close();
        }
        [TestMethod]
        public void ReadStreamWithTemplateWorkSheet()
        {
            FileStream instream = new FileStream(@"Test\Worksheet.xlsx", FileMode.Open, FileAccess.Read);
            MemoryStream stream = new MemoryStream();
            using (ExcelPackage pck = new ExcelPackage(stream, instream))
            {
                var ws = pck.Workbook.Worksheets["Perf"];
                Assert.AreEqual(ws.Cells["H6"].Formula, "B5+B6");
                pck.SaveAs(new FileInfo(@"Test\Worksheet2.xlsx"));
            }
            instream.Close();
        }
        [TestMethod]
        public void ReadStreamSaveAsStream()
        {
            FileStream instream = new FileStream(@"Test\Worksheet.xlsx", FileMode.Open, FileAccess.ReadWrite);
            MemoryStream stream = new MemoryStream();
            using (ExcelPackage pck = new ExcelPackage(instream))
            {
                var ws = pck.Workbook.Worksheets["Perf"];
                pck.SaveAs(stream);
            }
            instream.Close();
        }
    }
}
