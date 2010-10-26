
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Drawing;
using System.Drawing;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Style;
using System.Data;

namespace ExcelPackageTest
{
    [TestClass]
    public class WorkSheetTest
    {
        private TestContext testContextInstance;
        private static ExcelPackage _pck;
        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        [ClassInitialize()]
        public static void MyClassInitialize(TestContext testContext)
        {
            Directory.CreateDirectory(string.Format("Test"));
            _pck = new ExcelPackage(new FileInfo("Test\\Worksheet.xlsx"));
        }

        // Use ClassCleanup to run code after all tests in a class have run
        [ClassCleanup()]
        public static void MyClassCleanup()
        {
            _pck = null;
        }
        [TestMethod]
        public void LoadData()
        {
            ExcelWorksheet ws = _pck.Workbook.Worksheets.Add("newsheet");
            ws.Cells["U19"].Value = new DateTime(2009, 12, 31);
            ws.Cells["U20"].Value = new DateTime(2010, 1, 1);
            ws.Cells["U21"].Value = new DateTime(2010, 1, 2);
            ws.Cells["U22"].Value = new DateTime(2010, 1, 3);
            ws.Cells["U23"].Value = new DateTime(2010, 1, 4);
            ws.Cells["U24"].Value = new DateTime(2010, 1, 5);
            ws.Cells["U19:U24"].Style.Numberformat.Format = "yyyy-mm-dd";

            ws.Cells["V19"].Value = 100;
            ws.Cells["V20"].Value = 102;
            ws.Cells["V21"].Value = 101;
            ws.Cells["V22"].Value = 103;
            ws.Cells["V23"].Value = 105;
            ws.Cells["V24"].Value = 104;
            ws.Cells["v19:v24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells["v19:v24"].Style.Numberformat.Format = @"$#,##0.00_);($#,##0.00)";

            ws.Cells["X19"].Value = 210;
            ws.Cells["X20"].Value = 212;
            ws.Cells["X21"].Value = 221;
            ws.Cells["X22"].Value = 123;
            ws.Cells["X23"].Value = 135;
            ws.Cells["X24"].Value = 134;

            // add autofilter
            ws.Cells["U19:X24"].AutoFilter = true;
            ExcelPicture pic = ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);
            pic.SetPosition(150, 140);

            ws.Cells["A30"].Value = "Text orientation 45";
            ws.Cells["A30"].Style.TextRotation = 45;
            ws.Cells["B30"].Value = "Text orientation 90";
            ws.Cells["B30"].Style.TextRotation = 90;
            ws.Cells["c30"].Value = "Text orientation 180";
            ws.Cells["c30"].Style.TextRotation = 180;
            ws.Cells["D30"].Value = "Text orientation 38";
            ws.Cells["D30"].Style.TextRotation = 38;
            ws.Cells["D30"].Style.Font.Bold = true;
            ws.Cells["D30"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            ws.Workbook.Names.Add("TestName", ws.Cells["B30:E30"]);
            ws.Workbook.Names["TestName"].Style.Font.Color.SetColor(Color.Red);


            ws.Workbook.Names["TestName"].Offset(1, 0).Value = "Offset test 1";
            ws.Workbook.Names["TestName"].Offset(2,-1, 2, 2).Value = "Offset test 2";

            //Test vertical align
            ws.Cells["E19"].Value = "Subscript";
            ws.Cells["E19"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.Subscript;
            ws.Cells["E20"].Value = "Subscript";
            ws.Cells["E20"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;
            ws.Cells["E21"].Value = "Superscript";
            ws.Cells["E21"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.Superscript;
            ws.Cells["E21"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.None;


            ws.Cells["E22"].Value = "Indent 2";
            ws.Cells["E22"].Style.Indent = 2;
            ws.Cells["E23"].Value = "Shrink to fit";
            ws.Cells["E23"].Style.ShrinkToFit=true;


            ws.Names.Add("SheetName", ws.Cells["A1:A2"]);
            ws.View.FreezePanes(3, 5);

            foreach (ExcelRangeBase cell in ws.Cells["A1"])
            {
                Assert.Fail("A1 is not set");
            }

            //foreach (ExcelRangeBase cell in ws.Cells["B1:G25,A1"])
            //{
            //    TestContext.WriteLine(cell.Address);
            //}

            foreach (ExcelRangeBase cell in ws.Cells[ws.Dimension.Address])
            {
                //TestContext.WriteLine(cell.Address);
                System.Diagnostics.Debug.WriteLine(cell.Address);
            }
            
            ////Linq test
            var res = from c in ws.Cells[ws.Dimension.Address] where c.Value !=null &&  c.Value.ToString() == "Offset test 1" select c;

            foreach (ExcelRangeBase cell in res)
            {
                System.Diagnostics.Debug.WriteLine(cell.Address);
            }

            _pck.Workbook.Properties.Author = "Jan Källman";
            _pck.Workbook.Properties.Category="Category";
            _pck.Workbook.Properties.Comments = "Comments";
            _pck.Workbook.Properties.Company="Adventure works";
            _pck.Workbook.Properties.Keywords = "Keywords";
            _pck.Workbook.Properties.Title = "Title";
            _pck.Workbook.Properties.Subject = "Subject";
            _pck.Workbook.Properties.Status = "Status";
            _pck.Workbook.Properties.HyperlinkBase = new Uri("http://serversideexcel.com",UriKind.Absolute );
            _pck.Workbook.Properties.Manager= "Manager";

            _pck.Workbook.Properties.SetCustomPropertyValue("DateTest", new DateTime(2008, 12, 31));
            TestContext.WriteLine(_pck.Workbook.Properties.GetCustomPropertyValue("DateTest").ToString());
            _pck.Workbook.Properties.SetCustomPropertyValue("Author", "Jan Källman");
            _pck.Workbook.Properties.SetCustomPropertyValue("Count", 1);
            _pck.Workbook.Properties.SetCustomPropertyValue("IsTested", false);
            _pck.Workbook.Properties.SetCustomPropertyValue("LargeNo", 123456789123);
            _pck.Workbook.Properties.SetCustomPropertyValue("Author", 3);
        }
        const int PERF_ROWS=5000;
        [TestMethod]
        public void Performance()
        {
            ExcelWorksheet ws=_pck.Workbook.Worksheets.Add("Perf");
            TestContext.WriteLine("StartTime {0}", DateTime.Now);

            Random r = new Random();
            for (int i = 1; i <= PERF_ROWS; i++)
            {
                ws.Cells[i,1].Value=string.Format("Row {0}\n.Test new row\"'",i);
                ws.Cells[i,2].Value=i;
                ws.Cells[i, 2].Style.WrapText = true;
                ws.Cells[i,3].Value=DateTime.Now;
                ws.Cells[i, 4].Value = r.NextDouble()*100000;                
            }            
            ws.Cells[1, 2, PERF_ROWS, 2].Style.Numberformat.Format="#,##0";
            ws.Cells[1, 3, PERF_ROWS, 3].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
            ws.Cells[1, 4, PERF_ROWS, 4].Style.Numberformat.Format = "#,##0.00";
            ws.Cells[PERF_ROWS + 1, 2].Formula = "SUM(B1:B" + PERF_ROWS.ToString() +")";
            ws.Column(1).Width = 12;
            ws.Column(2).Width = 8;
            ws.Column(3).Width = 20;
            ws.Column(4).Width = 14;
            ws.DeleteRow(1000, 3, true);
            ws.DeleteRow(2000, 1, true);

            ws.InsertRow(2001, 4);

            ws.InsertRow(2010, 1, 2010);

            ws.InsertRow(20000, 2);

            ws.DeleteRow(20005, 4, false);

            //Single formula
            ws.Cells["H3"].Formula = "B2+B3";
            ws.DeleteRow(2, 1, true);

            //Shared formula
            ws.Cells["H5:H30"].Formula = "B4+B5";
            ws.Cells["H5:H30"].Style.Numberformat.Format= "_(\"$\"* # ##0.00_);_(\"$\"* (# ##0.00);_(\"$\"* \"-\"??_);_(@_)";
            ws.InsertRow(7, 3);
            ws.InsertRow(2, 1);
            ws.DeleteRow(30, 3, true);

            ws.DeleteRow(15, 2, true);
            ws.Cells["a1:B100"].Style.Locked = false;
            ws.Cells["a1:B12"].Style.Hidden = true;
            TestContext.WriteLine("EndTime {0}", DateTime.Now);
        }
        [TestMethod]
        public void InsertDeleteTest()
        {
            ExcelWorksheet ws = _pck.Workbook.Worksheets.Add("InsertDelete");
            ws.Cells["A1:C5"].Value = 1;
            ws.Cells["A1:B3"].Merge = true;
            ws.Cells["D3"].Formula = "A2+C5";
            ws.InsertRow(2, 1);

            ws.Cells["A10:C15"].Value = 1;
            ws.Cells["A11:B13"].Merge = true;
            ws.DeleteRow(12, 1,true);

            ws.Cells["a1:B100"].Style.Locked = false;
            ws.Cells["a1:B12"].Style.Hidden = true;
            ws.Protection.IsProtected=true;
            ws.Protection.SetPassword("Password");


            var range = ws.Cells["B2:D100"];

            ws.PrinterSettings.PrintArea=null;
            ws.PrinterSettings.PrintArea=ws.Cells["B2:D99"];
            ws.PrinterSettings.PrintArea = null;
            ws.Row(15).PageBreak = true;
            ws.Column(3).PageBreak = true;
            ws.View.ShowHeaders = false;
            ws.View.PageBreakView = true;

            ws.Workbook.CalcMode = ExcelCalcMode.Automatic;

            Assert.AreEqual(range.Start.Column, 2);
            Assert.AreEqual(range.Start.Row, 2);
            Assert.AreEqual(range.Start.Address, "B2");

            Assert.AreEqual(range.End.Column, 4);
            Assert.AreEqual(range.End.Row, 100);
            Assert.AreEqual(range.End.Address, "D100");

            ExcelAddress addr = new ExcelAddress("B1:D3");

            Assert.AreEqual(addr.Start.Column, 2);
            Assert.AreEqual(addr.Start.Row, 1);
            Assert.AreEqual(addr.End.Column, 4);
            Assert.AreEqual(addr.End.Row, 3);
        }
        [TestMethod]
        public void RichTextCells()
        {
            ExcelWorksheet ws = _pck.Workbook.Worksheets.Add("RichText");
            var rs = ws.Cells["A1"].RichText;

            var r1 = rs.Add("Test");
            r1.Bold = true;
            r1.Color = Color.Pink;
            
            var r2 = rs.Add(" of");
            r2.Size = 14;
            r2.Italic = true;

            var r3 = rs.Add(" rich");
            r3.FontName = "Arial";
            r3.Size = 18;
            r3.Italic = true;

            var r4 = rs.Add("text.");
            r4.Size = 8.25f;
            r4.Italic = true;
            r4.UnderLine = true;

            rs=ws.Cells["A3:A4"].RichText;

            var r5 = rs.Add("Double");
            r5.Color = Color.PeachPuff;
            r5.FontName = "times new roman";
            r5.Size = 16;

            var r6 = rs.Add(" cells");
            r6.Color = Color.Red;
            r6.UnderLine=true;
        }
        [TestMethod]
        public void SaveWorksheet()
        {
            _pck.Save();
        }
        [TestMethod]
        public void TestComments()
        {
            var ws = _pck.Workbook.Worksheets.Add("Comment");            
            var comment = ws.Comments.Add(ws.Cells["B2"], "Jan Källman\r\nAuthor\r\n", "JK");            

            comment.RichText[0].Bold = true;
            comment.RichText[0].PreserveSpace = true;
            var rt = comment.RichText.Add("Test comment");
            comment.VerticalAlignment = eTextAlignVerticalVml.Center;
            comment.HorizontalAlignment = eTextAlignHorizontalVml.Center;
            comment.Visible = true;
            comment.BackgroundColor = Color.Green;
            comment.To.Row += 4;
            comment.To.Column += 2;
            comment.LineStyle = eLineStyleVml.LongDash;
            comment.LineColor = Color.Red;
            comment.LineWidth = (Single)2.5;
            rt.Color = Color.Red;

            var rt2=ws.Cells["C3"].AddComment("Range Added Comment test test test test test test test test test test testtesttesttesttesttesttesttesttesttesttest", "Jan Källman");
            ws.Cells["c3"].Comment.AutoFit = true;
        }
        [TestMethod]
        public void Address()
        {
            var ws = _pck.Workbook.Worksheets.Add("Address");
            ws.Cells["A1:A4,B5:B7"].Value = "AddressTest";
            ws.Cells["A1:A4,B5:B7"].Style.Font.Color.SetColor(Color.Red);
            ws.Cells["A2:A3,B4:B8"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.LightUp;
            ws.Cells["A2:A3,B4:B8"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            ws.Cells["2:2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells["2:2"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            ws.Cells["B:B"].Style.Font.Name = "Times New Roman";

            ws.Cells["C4:G4,H8:H30,B15"].FormulaR1C1 = "RC[-1]+R1C[-1]";
            ws.Cells["G1,G3"].Hyperlink = new ExcelHyperLink("Comment!$A$1","Comment");
            ws.Cells["G1,G3"].Style.Font.Color.SetColor(Color.Blue);
            ws.Cells["G1,G3"].Style.Font.UnderLine = true;

            ws.Cells["A1:G5"].Copy(ws.Cells["A50"]);
        }
        [TestMethod]
        public void WorksheetCopy()
        {
            var ws = _pck.Workbook.Worksheets.Add("Copied Address", _pck.Workbook.Worksheets["Address"]);
            var wsCopy = _pck.Workbook.Worksheets.Add("Copied Comment", _pck.Workbook.Worksheets["Comment"]);

            ExcelPackage pck2 = new ExcelPackage();
            pck2.Workbook.Worksheets.Add("Copy From other pck", _pck.Workbook.Worksheets["Address"]);
            pck2.SaveAs(new FileInfo("test\\copy.xlsx"));
            pck2=null;
            Assert.AreEqual(2, wsCopy.Comments.Count);
        }
        [TestMethod]
        public void TestDelete()
        {
            string file = "test.xlsx";

            if (File.Exists(file))
                File.Delete(file);

            Create(file);

            ExcelPackage pack = new ExcelPackage(new FileInfo (file ));
            ExcelWorksheet w = pack.Workbook.Worksheets["delete"];
            w.DeleteRow(2, 1);
            pack.Save();
        }
        static void Create(string file)
        {
            ExcelPackage pack = new ExcelPackage(new FileInfo(file));
            ExcelWorksheet w = pack.Workbook.Worksheets.Add("delete");
            w.Cells[1, 1].Value = "test";
            w.Cells[2, 1].Value = "to delete";
            pack.Save();
        }
        [TestMethod]
        public void RowStyle()
        {
            FileInfo newFile = new FileInfo(@"sample8.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(@"sample8.xlsx");
            }

            ExcelPackage package = new ExcelPackage();
            //Load the sheet with one string column, one date column and a few random numbers.
            var ws = package.Workbook.Worksheets.Add("First line test");

            ws.Cells[1, 1].Value = "1; 1";
            ws.Cells[2, 1].Value = "2; 1";
            ws.Cells[1, 2].Value = "1; 2";
            ws.Cells[2, 2].Value = "2; 2";

            ws.Row(1).Style.Font.Bold = true;
            ws.Column(1).Style.Font.Bold = true;
            package.SaveAs(newFile);

        }
        [TestMethod]
        public void HideTest()
        {
            var ws = _pck.Workbook.Worksheets.Add("Hidden");
            ws.Cells["A1"].Value = "This workbook is hidden"    ;
            ws.Hidden = eWorkSheetHidden.Hidden;
        }
        [TestMethod]
        public void VeryHideTest()
        {
            var ws = _pck.Workbook.Worksheets.Add("VeryHidden");
            ws.Cells["a1"].Value = "This workbook is hidden";
            ws.Hidden = eWorkSheetHidden.VeryHidden;
        }
        [TestMethod]
        public void PrinterSettings()
        {
            var ws = _pck.Workbook.Worksheets.Add("Sod/Hydroseed");

            ws.Cells[1, 1].Value = "1; 1";
            ws.Cells[2, 1].Value = "2; 1";
            ws.Cells[1, 2].Value = "1; 2";
            ws.Cells[2, 2].Value = "2; 2";
            ws.Cells[1, 1, 1, 2].AutoFilter = true;
            ws.PrinterSettings.BlackAndWhite = true;
            ws.PrinterSettings.ShowGridLines = true;
            ws.PrinterSettings.ShowHeaders = true;
            ws.PrinterSettings.PaperSize = ePaperSize.A4;

            ws.PrinterSettings.RepeatRows = new ExcelAddress("1:1");
            ws.PrinterSettings.RepeatColumns = new ExcelAddress("A:A");

            ws.PrinterSettings.Draft = true;

            ws.PrinterSettings.PrintArea=ws.Cells["A1:B2"];

            ws.Select(new ExcelAddress("3:4,E5:F6"));
        }
        [TestMethod]
        public void StyleNameTest()
        {
            var ws = _pck.Workbook.Worksheets.Add("StyleNameTest");

            ws.Cells[1, 1].Value = "R1 C1";
            ws.Cells[1, 2].Value = "R1 C2";
            ws.Cells[1, 3].Value = "R1 C3";
            ws.Cells[2, 1].Value = "R2 C1";
            ws.Cells[2, 2].Value = "R2 C2";
            ws.Cells[2, 3].Value = "R2 C3";

            var ns=_pck.Workbook.Styles.CreateNamedStyle("TestStyle");
            ns.Style.Font.Bold = true;

            ws.Cells["A1:C1"].StyleName = "TestStyle";
            ws.defaultRowHeight = 35;
        }
        [TestMethod]
        public void ValueError()
        {
            var ws = _pck.Workbook.Worksheets.Add("ValueError");

            ws.Cells[1, 1].Value = "Domestic Violence&#xB; and the Professional";
            var rt=ws.Cells[1, 2].RichText.Add("Domestic Violence&#xB; and the Professional 2");
            TestContext.WriteLine(rt.Bold.ToString());
            rt.Bold = true;
            TestContext.WriteLine(rt.Bold.ToString());
        }
        [TestMethod]
        public void FormulaError()
        {
            var ws = _pck.Workbook.Worksheets.Add("FormulaError");

            ws.Cells["A1:K3"].Formula = "A3+A4";
            //ws.Cells["B2:I2"].Formula = "";   //Error
        }
        [TestMethod]
        public void TableTest()
        {            
            var ws = _pck.Workbook.Worksheets.Add("Table");
            ws.Cells["B1"].Value = 123;
            var tbl = ws.Tables.Add(ws.Cells["B1:U12"], "TestTable");
            tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Custom;

            tbl.ShowFirstColumn = true;
            tbl.ShowTotal = true;
            tbl.ShowHeader = true;
            tbl.ShowLastColumn = true;

            ws.Cells["K2"].Value = 5;
            ws.Cells["J3"].Value = 4;

            tbl.Columns[8].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
            tbl.Columns[9].TotalsRowFormula = string.Format("SUM([{0}])",tbl.Columns[9].Name);

            ws.Cells["B2"].Value = 1;
            ws.Cells["B3"].Value = 2;
            ws.Cells["B4"].Value = 3;
            ws.Cells["B5"].Value = 4;
            ws.Cells["B6"].Value = 5;
            ws.Cells["B7"].Value = 6;
            ws.Cells["B8"].Value = 7;
            ws.Cells["B9"].Value = 8;
            ws.Cells["B10"].Value = 9;
            ws.Cells["B11"].Value = 10;
            ws.Cells["B12"].Value = 11;


            ws.Cells["C7"].Value = "Table test";
            ws.Cells["C8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C8"].Style.Fill.BackgroundColor.SetColor(Color.Red);

            tbl=ws.Tables.Add(ws.Cells["a12:a13"], "");

            tbl = ws.Tables.Add(ws.Cells["C16:Y35"], "");
            tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Medium14;
            tbl.ShowFirstColumn = true;
            tbl.ShowLastColumn = true;
            tbl.ShowColumnStripes = true;
            tbl.Columns[2].Name = "Test Column Name";        
        }

        [TestMethod]
        public void Stylebug()
        {
            ExcelPackage p = new ExcelPackage(new FileInfo(@"c:\temp\FullProjecte.xlsx"));

            var ws = p.Workbook.Worksheets.First();
            ws.Cells[12, 1].Value = 0;
            ws.Cells[12, 2].Value = new DateTime(2010,9,14);
            ws.Cells[12, 3].Value = "Federico Lois";
            ws.Cells[12, 4].Value = "Nakami";
            ws.Cells[12, 5].Value = "Hores";
            ws.Cells[12, 7].Value = 120;
            ws.Cells[12, 8].Value="A definir";
            ws.Cells[12, 9].Value = new DateTime(2010,9,14);
            ws.Cells[12, 10].Value = new DateTime(2010,9,14);
            ws.Cells[12, 11].Value = "Transferència";

            ws.InsertRow(13, 1, 12);
            ws.Cells[13, 1].Value = 0;
            ws.Cells[13, 2].Value = new DateTime(2010, 9, 14);
            ws.Cells[13, 3].Value = "Federico Lois";
            ws.Cells[13, 4].Value = "Nakami";
            ws.Cells[13, 5].Value = "Hores";
            ws.Cells[13, 7].Value = 120;
            ws.Cells[13, 8].Value = "A definir";
            ws.Cells[13, 9].Value = new DateTime(2010, 9, 14);
            ws.Cells[13, 10].Value = new DateTime(2010, 9, 14);
            ws.Cells[13, 11].Value = "Transferència";

            ws.InsertRow(14, 1, 13);

            ws.InsertRow(19, 1, 19);
            ws.InsertRow(26, 1, 26);
            ws.InsertRow(33, 1, 33);
            p.SaveAs(new FileInfo(@"c:\temp\FullProjecte_new.xlsx"));
        }
        [TestMethod]
        public void FormulaOverwrite()
        {
            var ws = _pck.Workbook.Worksheets.Add("FormulaOverwrite");
            //Inside
            ws.Cells["A1:G12"].Formula = "B1+C1";
            ws.Cells["B2:C3"].Formula = "G2+E1";


            //Top bottom overwrite
            ws.Cells["A14:G26"].Formula = "B1+C1+D1";
            ws.Cells["B13:C28"].Formula = "G2+E1";

            //Top bottom overwrite
            ws.Cells["B30:E42"].Formula = "B1+C1+$D$1";
            ws.Cells["A32:H33"].Formula = "G2+E1";

            ws.Cells["A50:A59"].CreateArrayFormula("C50+D50");

            ws.Cells["A15"].Value = "Värde";
            ws.Cells["C12"].AddComment("Test", "JJOD");
            ws.Cells["D12:I12"].Merge = true;
            ws.Cells["D12:I12"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["D12:I12"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
            ws.Cells["D12:I12"].Style.WrapText = true;
        }
        [TestMethod]
        public void DefinedName()
        {
            var ws = _pck.Workbook.Worksheets.Add("Names");
            ws.Names.Add("RefError", ws.Cells["#REF!"]);

            ws.Cells["A1"].Value = "Test";
            ws.Cells["A1"].Style.Font.Size = 8.5F;
        }
        [TestMethod]
        public void LoadDataTable()
        {
            var ws = _pck.Workbook.Worksheets.Add("Loaded DataTable");

            var dt = new DataTable();
            dt.Columns.Add("String", typeof(string));
            dt.Columns.Add("Int", typeof(int));
            dt.Columns.Add("Bool", typeof(bool));
            dt.Columns.Add("Double", typeof(double));


            var dr=dt.NewRow();
            dr[0] = "Row1";
            dr[1] = 1;
            dr[2] = true;
            dr[3] = 1.5;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = "Row2";
            dr[1] = 2;
            dr[2] = false;
            dr[3] = 2.25;
            dt.Rows.Add(dr);

            dr = dt.NewRow();
            dr[0] = "Row3";
            dr[1] = 3;
            dr[2] = true;
            dr[3] = 3.125;
            dt.Rows.Add(dr);

            ws.Cells["A1"].LoadFromDataTable(dt,true,OfficeOpenXml.Table.TableStyles.Medium5);
        }
        [TestMethod]
        public void LoadText()
        {
            var ws = _pck.Workbook.Worksheets.Add("Loaded Text");

            ws.Cells["A1"].LoadFromText("1.2");
            ws.Cells["A2"].LoadFromText("1,\"Test av data\",\"12,2\",\"\"Test\"\"");
            ws.Cells["A3"].LoadFromText("\"1,3\",\"Test av \"\"data\",\"12,2\",\"Test\"\"\"", new ExcelTextFormat() { TextQualifier = '"' });

            ws = _pck.Workbook.Worksheets.Add("File1");
            ws.Cells["A1"].LoadFromText(new FileInfo(@"c:\temp\csv\et1c1004.csv"), new ExcelTextFormat() {SkipLinesBeginning=3,SkipLinesEnd=1, EOL="\n"});

            ws = _pck.Workbook.Worksheets.Add("File2");
            ws.Cells["A1"].LoadFromText(new FileInfo(@"c:\temp\csv\etiv2812.csv"), new ExcelTextFormat() { SkipLinesBeginning = 3, SkipLinesEnd = 1, EOL = "\n" });

            //ws = _pck.Workbook.Worksheets.Add("File3");
            //ws.Cells["A1"].LoadFromText(new FileInfo(@"c:\temp\csv\last_gics.txt"), new ExcelTextFormat() { SkipLinesBeginning = 1, Delimiter='|'});

            ws = _pck.Workbook.Worksheets.Add("File4");
            ws.Cells["A1"].LoadFromText(new FileInfo(@"c:\temp\csv\20060927.custom_open_positions.cdf.SPP"), new ExcelTextFormat() { SkipLinesBeginning = 2, SkipLinesEnd=2, TextQualifier='"', DataTypes=new ExcelTextFormat.eDataTypes[] {ExcelTextFormat.eDataTypes.Number,ExcelTextFormat.eDataTypes.String, ExcelTextFormat.eDataTypes.Number, ExcelTextFormat.eDataTypes.Number, ExcelTextFormat.eDataTypes.Number, ExcelTextFormat.eDataTypes.String, ExcelTextFormat.eDataTypes.Number, ExcelTextFormat.eDataTypes.Number, ExcelTextFormat.eDataTypes.String, ExcelTextFormat.eDataTypes.String, ExcelTextFormat.eDataTypes.Number, ExcelTextFormat.eDataTypes.Number, ExcelTextFormat.eDataTypes.Number}},
                OfficeOpenXml.Table.TableStyles.Medium27, true);            
        }
        [TestMethod]
        public void LoadArray()
        {
            var ws = _pck.Workbook.Worksheets.Add("Loaded Array");
            List<object[]> testArray = new List<object[]>() { new object[] { 3, 4, 5, 6 }, new string[] { "Test1", "test", "5", "6" } };
            ws.Cells["A1"].LoadFromArrays(testArray);
        }
    }
}
