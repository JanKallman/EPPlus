using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Security.Cryptography.X509Certificates;

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
                ws = pck.Workbook.Worksheets["Scatter"];
                var cht = ws.Drawings["ScatterChart1"] as ExcelScatterChart;
                Assert.AreEqual(cht.Title.Text, "Header  Text");
                cht.Title.Text = "Test";
                Assert.AreEqual(cht.Title.Text, "Test");
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

                ws = pck.Workbook.Worksheets["HeaderImage"];

                Assert.AreEqual(ws.HeaderFooter.Pictures.Count, 3);

                ws = pck.Workbook.Worksheets["newsheet"];
                Assert.AreEqual(ws.Cells["F2"].Style.Font.UnderLine,true);
                Assert.AreEqual(ws.Cells["F2"].Style.Font.UnderLineType, ExcelUnderLineType.Double);
                Assert.AreEqual(ws.Cells["F3"].Style.Font.UnderLineType, ExcelUnderLineType.SingleAccounting);
                Assert.AreEqual(ws.Cells["F5"].Style.Font.UnderLineType, ExcelUnderLineType.None);
                Assert.AreEqual(ws.Cells["F5"].Style.Font.UnderLine, false);
                
                //Assert.AreEqual(ws.HeaderFooter.Pictures[0].Name, "");
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
        [TestMethod]
        public void CondFormatDataValBug()
        {            
            var file = new FileInfo(@"c:\temp\condi.xlsx");
            using (ExcelPackage pck = new ExcelPackage(file))
            {
                var dv = pck.Workbook.Worksheets[1].Cells["A1"].DataValidation.AddIntegerDataValidation();
                dv.Formula.Value = 1;
                dv.Formula2.Value = 4;
                dv.Operator = OfficeOpenXml.DataValidation.ExcelDataValidationOperator.equal;
                pck.SaveAs(new FileInfo(@"c:\temp\condi2.xlsx"));
            }
        }
        [TestMethod]
        public void ReadBug4()
        {
            var lines = new List<string>();
            var package = new ExcelPackage(new FileInfo(@"c:\temp\test.xlsx"));

            ExcelWorkbook workBook = package.Workbook;
            if (workBook != null)
            {
            if (workBook.Worksheets.Count > 0) //fails on this line
            {
            // Get the first worksheet
            ExcelWorksheet currentWorksheet = workBook.Worksheets.First();

            var rowCount = 1;
            var lastRow = currentWorksheet.Dimension.End.Row;
            var lastColumn = currentWorksheet.Dimension.End.Column;
            while (rowCount <= lastRow)
            {
            var columnCount = 1;
            var line = "";
            while (columnCount <= lastColumn)
            {
            line += currentWorksheet.Cells[rowCount, columnCount].Value + "|";
            columnCount++;
            }
            lines.Add(line);
            rowCount++;
            }
            }
            }
        }
        [TestMethod]
        public void ReadBug5()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\2.9 bugs\protect.xlsx"));

            package.Workbook.Worksheets[1].Protection.AllowInsertColumns = true;
            package.Workbook.Worksheets[1].Protection.SetPassword("test");
            package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\protectnew.xlsx"));
        }
        [TestMethod]
        public void ReadBug6()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\2.9 bugs\outofrange\error.xlsx"));

            package.Workbook.Worksheets[1].Protection.AllowInsertColumns = true;
            package.Workbook.Worksheets[1].Protection.SetPassword("test");
            package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\error.xlsx"));
        }
        [TestMethod]
        public void ReadBug7()
        {
            var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("test");
            using (var rng = ws.Cells["A1"])
            {
                var rt1 = rng.RichText.Add("TEXT1\r\n");
                rt1.Bold = true;
                rng.Style.WrapText = true;
                var rt2=rng.RichText.Add("TEXT2");
                rt2.Bold = false;
            }
            
            package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\error.xlsx"));
        }
        [TestMethod]
        public void ReadBug8()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\2.9 bugs\bug\Genband SO CrossRef Phoenix.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            using (var rng = ws.Cells["A1"])
            {
                var rt1 = rng.RichText.Add("TEXT1\r\n");
                rt1.Bold = true;
                rng.Style.WrapText = true;
                var rt2 = rng.RichText.Add("TEXT2");
                rt2.Bold = false;
            }

            package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\billing_template.xlsx.error"));
        }
        [TestMethod]
        public void ReadBug9()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\CovenantsCheckReportTemplate.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            package.SaveAs(new FileInfo(@"c:\temp\2.9 bugs\new_t.xlsx"));
        }
        [TestMethod]
        public void ReadBug10()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\Model_graphes_MBW.xlsm"));

            var ws = package.Workbook.Worksheets["HTTP_data"];
            Assert.IsNotNull(ws.Cells["B4"].Style.Fill.BackgroundColor.Indexed);
            Assert.IsNotNull(ws.Cells["B5"].Style.Fill.BackgroundColor.Indexed);
        }
        [TestMethod]
        public void ReadBug11()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\test.xlsx"));
            var ws = package.Workbook.Worksheets[1];

        }
        [TestMethod]
        public void ReadBug12()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\sample.xlsx"));
            var ws = package.Workbook.Worksheets[1];
            var pck2 = new ExcelPackage();
            pck2.Workbook.Worksheets.Add("Test", ws);
            pck2.SaveAs(new FileInfo(@"c:\temp\SampleNew.xlsx"));
        }
        [TestMethod]
        public void ReadVBA()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\vba.xlsm"));
            foreach (var module in package.Workbook.VbaProject.Modules)
            {
                Assert.AreNotEqual(module, null);
            }

            List<X509Certificate2> ret = new List<X509Certificate2>();
            X509Store store = new X509Store(StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            foreach (var c in store.Certificates)
            {
                ret.Add(c);
            }
            
            package.Workbook.VbaProject.Signature.Certificate = store.Certificates[8];
            package.Workbook.VbaProject.Signature.Save(package.Workbook.VbaProject);
            package.Save();
            //Assert.AreNotEqual(package.Workbook.VbaProject.Signature.Uri.AbsolutePath, "");
        }
    }
}
