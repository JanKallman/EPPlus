using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace EPPlusTest
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
                Assert.AreEqual(comment.Author, "Jan Källman");

                ws = pck.Workbook.Worksheets["Hidden"];
                Assert.AreEqual<eWorkSheetHidden>(ws.Hidden, eWorkSheetHidden.Hidden);

                ws = pck.Workbook.Worksheets["VeryHidden"];
                Assert.AreEqual<eWorkSheetHidden>(ws.Hidden, eWorkSheetHidden.VeryHidden);

                ws = pck.Workbook.Worksheets["RichText"];
                Assert.AreEqual("Room 02 & 03", ws.Cells["G1"].RichText.Text);
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

                ws = pck.Workbook.Worksheets["newsheet"];
                Assert.AreEqual(ws.GetValue<DateTime>(20 ,21),new DateTime(2010,1,1));

                ws = pck.Workbook.Worksheets["Loaded DataTable"];                
                Assert.AreEqual(ws.GetValue<string>(2 ,1),"Row1");
                Assert.AreEqual(ws.GetValue<int>(2, 2), 1);
                Assert.AreEqual(ws.GetValue<bool>(2, 3), true);
                Assert.AreEqual(ws.GetValue<double>(2, 4), 1.5);

                ws=pck.Workbook.Worksheets["RichText"];

                var r1 = ws.Cells["A1"].RichText[0];
                Assert.AreEqual(r1.Text,"Test");
                Assert.AreEqual(r1.Bold, true);
                //r1.Bold = true;
                //r1.Color = Color.Pink;

                //var r2 = rs.Add(" of");
                //r2.Size = 14;
                //r2.Italic = true;

                //var r3 = rs.Add(" rich");
                //r3.FontName = "Arial";
                //r3.Size = 18;
                //r3.Italic = true;

                //var r4 = rs.Add("text.");



                Assert.AreEqual(pck.Workbook.Worksheets["Address"].GetValue<string>(40,1),"\b\t");

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
        [TestMethod]
        public void ReadBlankStream()
        {
            MemoryStream stream = new MemoryStream();
            using (ExcelPackage pck = new ExcelPackage(stream))
            {
                var ws = pck.Workbook.Worksheets["Perf"];
                pck.SaveAs(stream);
            }
            stream.Close();
        }
        [TestMethod]
        public void ReadBug()
        {
            var file = new FileInfo(@"c:\temp\Adenoviridae Protocol.xlsx");
            using (ExcelPackage pck = new ExcelPackage(file))
            {
                pck.Workbook.Worksheets[1].Cells["G4"].Value=12;
                pck.SaveAs(new FileInfo(@"c:\temp\Adenoviridae Protocol2.xlsx"));
            }
        }
        [TestMethod]
        public void ReadBug3()
        {
            ExcelPackage xlsPack = new ExcelPackage(new FileInfo(@"c:\temp\billing_template.xlsx"));
            ExcelWorkbook xlsWb = xlsPack.Workbook;
            ExcelWorksheet xlsSheet = xlsWb.Worksheets["Billing"];
        }
        [TestMethod]
        public void ReadBug2()
        {
            var file = new FileInfo(@"c:\temp\book2.xlsx");
            using (ExcelPackage pck = new ExcelPackage(file))
            {
                Assert.AreEqual("Good", pck.Workbook.Worksheets[1].Cells["A1"].StyleName);
                Assert.AreEqual("Good 2", pck.Workbook.Worksheets[1].Cells["C1"].StyleName);
                Assert.AreEqual("Note", pck.Workbook.Worksheets[1].Cells["G11"].StyleName);
                pck.SaveAs(new FileInfo(@"c:\temp\Adenoviridae Protocol2.xlsx"));
            }
        }
    }
}
