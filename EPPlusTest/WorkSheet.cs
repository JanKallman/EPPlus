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
using OfficeOpenXml.Table.PivotTable;
using System.Reflection;

namespace EPPlusTest
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
            ws.Cells["C30"].Value = "Text orientation 180";
            ws.Cells["C30"].Style.TextRotation = 180;
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

            ws.Cells["e24"].Value = "ReadingOrder LeftToRight";
            ws.Cells["e24"].Style.ReadingOrder = ExcelReadingOrder.LeftToRight;
            ws.Cells["e25"].Value = "ReadingOrder RightToLeft";
            ws.Cells["e25"].Style.ReadingOrder = ExcelReadingOrder.RightToLeft;
            ws.Cells["e26"].Value = "ReadingOrder Context";
            ws.Cells["e26"].Style.ReadingOrder = ExcelReadingOrder.ContextDependent;
            ws.Cells["e27"].Value = "Default Readingorder";

            //Underline

            ws.Cells["F1:F7"].Value = "Underlined";
            ws.Cells["F1"].Style.Font.UnderLineType = ExcelUnderLineType.Single;
            ws.Cells["F2"].Style.Font.UnderLineType = ExcelUnderLineType.Double;
            ws.Cells["F3"].Style.Font.UnderLineType = ExcelUnderLineType.SingleAccounting;
            ws.Cells["F4"].Style.Font.UnderLineType = ExcelUnderLineType.DoubleAccounting;
            ws.Cells["F5"].Style.Font.UnderLineType = ExcelUnderLineType.None;
            ws.Cells["F6:F7"].Style.Font.UnderLine = true;
            ws.Cells["F7"].Style.Font.UnderLine = false;



            ws.Names.Add("SheetName", ws.Cells["A1:A2"]);
            ws.View.FreezePanes(3, 5);

            foreach (ExcelRangeBase cell in ws.Cells["A1"])
            {
                Assert.Fail("A1 is not set");
            }

            foreach (ExcelRangeBase cell in ws.Cells[ws.Dimension.Address])
            {
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
                ws.Cells[i, 3].Value=DateTime.Now;
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
            
            ws.Cells["A1:C1"].Merge = true;
            ws.Cells["A2:A5"].Merge = true;
            ws.DeleteRow(1, 1);
            ws.InsertRow(1, 1);
            ws.InsertRow(3, 1);
            
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
            //ws.Cells.Value = 0;
            ws.Cells["A1:C5"].Value = 1;
            Assert.AreEqual(((object[,])ws.Cells["A1:C5"].Value)[1, 1], 1);
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

            ws.Row(200).Height = 50;
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


            rs = ws.Cells["C8"].RichText;
            r1 = rs.Add("Blue ");
            r1.Color = Color.Blue;

            r2 = rs.Add("Red");
            r2.Color = Color.Red;

            ws.Cells["G1"].RichText.Add("Room 02 & 03");
            ws.Cells["G2"].RichText.Text = "Room 02 & 03";

            ws = ws = _pck.Workbook.Worksheets.Add("RichText2");
            ws.Cells["A1"].RichText.Text = "Room 02 & 03";
            ws.TabColor = Color.PowderBlue;

            r1=ws.Cells["G3"].RichText.Add("Test");
            r1.Bold = true;
            ws.Cells["G3"].RichText.Add(" a new t");
            ws.Cells["G3"].RichText[1].Bold = false; ;
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
            var comment = ws.Comments.Add(ws.Cells["C3"], "Jan Källman\r\nAuthor\r\n", "JK");            
            comment.RichText[0].Bold = true;
            comment.RichText[0].PreserveSpace = true;
            var rt = comment.RichText.Add("Test comment");
            comment.VerticalAlignment = eTextAlignVerticalVml.Center;
            
            comment = ws.Comments.Add(ws.Cells["A2"], "Jan Källman\r\nAuthor\r\n1", "JK");            
            comment = ws.Comments.Add(ws.Cells["A1"], "Jan Källman\r\nAuthor\r\n2", "JK");            
            comment = ws.Comments.Add(ws.Cells["C2"], "Jan Källman\r\nAuthor\r\n3", "JK");            
            comment = ws.Comments.Add(ws.Cells["C1"], "Jan Källman\r\nAuthor\r\n5", "JK");
            comment = ws.Comments.Add(ws.Cells["B1"], "Jan Källman\r\nAuthor\r\n7", "JK");

            ws.Comments.Remove(ws.Cells["A2"].Comment);
            //comment.HorizontalAlignment = eTextAlignHorizontalVml.Center;
            //comment.Visible = true;
            //comment.BackgroundColor = Color.Green;
            //comment.To.Row += 4;
            //comment.To.Column += 2;
            //comment.LineStyle = eLineStyleVml.LongDash;
            //comment.LineColor = Color.Red;
            //comment.LineWidth = (Single)2.5;
            //rt.Color = Color.Red;

            var rt2=ws.Cells["B2"].AddComment("Range Added Comment test test test test test test test test test test testtesttesttesttesttesttesttesttesttesttest", "Jan Källman");
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

            var ws2 = _pck.Workbook.Worksheets.Add("Copy Cells");
            ws.Cells["1:4"].Copy(ws2.Cells["1:1"]);

            ws.Cells["H1:J5"].Merge = true;
            ws.Cells["2:3"].Copy(ws.Cells["50:51"]);

            ws.Cells["A40"].Value = new string(new char[] {(char)8, (char)9});

            ExcelRange styleRng = ws.Cells["A1"];
            ExcelStyle tempStyle = styleRng.Style;
            var namedStyle = _pck.Workbook.Styles.CreateNamedStyle("HyperLink", tempStyle);  
        }
        [TestMethod]
        public void Encoding()
        {
            var ws = _pck.Workbook.Worksheets.Add("Encoding");
            ws.Cells["A1"].Value = "_x0099_";
            ws.Cells["A2"].Value = " Test \b" + (char)1 + " end\"";
            ws.Cells["A3"].Value = "_x0097_ test_x001D_1234";
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
            Assert.AreEqual(6, wsCopy.Comments.Count);
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
            w.DeleteRow(1, 2);
           
            pack.Save();
        }
        [TestMethod]
        public void LoadFromCollectionTest()
        {                        
            var ws = _pck.Workbook.Worksheets.Add("LoadFromCollection");
            List<TestDTO> list = new List<TestDTO>();
            list.Add(new TestDTO() { Id = 1, Name = "Item1", Boolean = false, Date = new DateTime(2011, 1, 1), dto = null, NameVar = "Field 1" });
            list.Add(new TestDTO() { Id = 2, Name = "Item2", Boolean = true, Date = new DateTime(2011, 1, 15), dto = new TestDTO(), NameVar = "Field 2" });
            list.Add(new TestDTO() { Id = 3, Name = "Item3", Boolean = false, Date = new DateTime(2011, 2, 1), dto = null, NameVar = "Field 3" });
            list.Add(new TestDTO() { Id = 4, Name = "Item4", Boolean = true, Date = new DateTime(2011, 4, 19), dto = list[1], NameVar = "Field 4" });
            list.Add(new TestDTO() { Id = 5, Name = "Item5", Boolean = false, Date = new DateTime(2011, 5, 8), dto = null, NameVar = "Field 5" });
            list.Add(new TestDTO() { Id = 6, Name = "Item6", Boolean = true, Date = new DateTime(2010, 3, 27), dto = null, NameVar = "Field 6" });
            list.Add(new TestDTO() { Id = 7, Name = "Item7", Boolean = false, Date = new DateTime(2009, 1, 5), dto = list[3], NameVar = "Field 7" });
            list.Add(new TestDTO() { Id = 8, Name = "Item8", Boolean = true, Date = new DateTime(2018, 12, 31), dto = null, NameVar = "Field 8" });
            list.Add(new TestDTO() { Id = 9, Name = "Item9", Boolean = false, Date = new DateTime(2010, 2, 1), dto = null, NameVar = "Field 9" });

            ws.Cells["A1"].LoadFromCollection(list, true);
            ws.Cells["A30"].LoadFromCollection(list, true, OfficeOpenXml.Table.TableStyles.Medium9, BindingFlags.Instance | BindingFlags.Instance, typeof(TestDTO).GetFields());

            ws.Cells["A45"].LoadFromCollection(list, true, OfficeOpenXml.Table.TableStyles.Light1, BindingFlags.Instance | BindingFlags.Instance, new MemberInfo[] { typeof(TestDTO).GetMethod("GetNameID"), typeof(TestDTO).GetField("NameVar") });
            ws.Cells["J1"].LoadFromCollection(from l in list where l.Boolean orderby l.Date select new { Name = l.Name, Id = l.Id, Date = l.Date, NameVariable = l.NameVar }, true, OfficeOpenXml.Table.TableStyles.Dark4);

            var ints = new int[] {1,3,4,76,2,5};
            ws.Cells["A15"].Value = ints;
            ws.Column(1).Hidden = true;
            ws.Cells.AutoFitColumns(0);
        }        
        static void Create(string file)
        {
            ExcelPackage pack = new ExcelPackage(new FileInfo(file));
            ExcelWorksheet w = pack.Workbook.Worksheets.Add("delete");
            w.Cells[1, 1].Value = "test";
            w.Cells[1, 2].Value = "test";
            w.Cells[2, 1].Value = "to delete";
            w.Cells[2, 2].Value = "to delete";
            w.Cells[3, 1].Value = "3Left";
            w.Cells[3, 2].Value = "3Left";
            w.Cells[4, 1].Formula = "B3+C3";
            w.Cells[4, 2].Value = "C3+D3";
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
        public void Hyperlink()
        {
            var ws = _pck.Workbook.Worksheets.Add("HyperLinks");
            var hl = new ExcelHyperLink("G1", "Till G1");
            hl.ToolTip = "Link to cell G1";
            ws.Cells["A1"].Hyperlink = hl;
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
            var r = ws.Cells["A26"];
            r.Value = "X";
            r.Worksheet.Row(26).PageBreak = true;
            ws.PrinterSettings.PrintArea=ws.Cells["A1:B2"];
            ws.PrinterSettings.HorizontalCentered = true;
            ws.PrinterSettings.VerticalCentered = true;

            ws.Select(new ExcelAddress("3:4,E5:F6"));

            ws = _pck.Workbook.Worksheets["RichText"];
            ws.PrinterSettings.RepeatColumns = ws.Cells["A:B"];
            ws.PrinterSettings.RepeatRows = ws.Cells["1:11"];
            ws.PrinterSettings.TopMargin = 1M;
            ws.PrinterSettings.LeftMargin = 1M;
            ws.PrinterSettings.BottomMargin = 1M;
            ws.PrinterSettings.RightMargin = 1M;
            ws.PrinterSettings.Orientation = eOrientation.Landscape;
            ws.PrinterSettings.PaperSize = ePaperSize.A4;
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
            ws.DefaultRowHeight = 35;
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

            ws.Cells["D5"].Formula = "COUNTIF(A1:A100,\"Miss\")";
            ws.Cells["A1:K3"].Formula = "A3+A4";
            ws.Cells["A4"].FormulaR1C1 = "+ROUNDUP(RC[1]/10,0)*10";

            //ws.Cells["B2:I2"].Formula = "";   //Error
        }
        [TestMethod]
        public void TableTest()
        {            
            var ws = _pck.Workbook.Worksheets.Add("Table");
            ws.Cells["B1"].Value = 123;
            var tbl = ws.Tables.Add(ws.Cells["B1:P12"], "TestTable");
            tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Custom;

            tbl.ShowFirstColumn = true;
            tbl.ShowTotal = true;
            tbl.ShowHeader = true;
            tbl.ShowLastColumn = true;
            tbl.ShowFilter = false;
            Assert.AreEqual(tbl.ShowFilter, false);
            ws.Cells["K2"].Value = 5;
            ws.Cells["J3"].Value = 4;

            tbl.Columns[8].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
            tbl.Columns[9].TotalsRowFormula = string.Format("SUM([{0}])",tbl.Columns[9].Name);
            tbl.Columns[14].CalculatedColumnFormula = "TestTable[[#This Row],[123]]+TestTable[[#This Row],[Column2]]";                                                       
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
            Assert.AreEqual(tbl.ShowFilter, true);
            tbl.Columns[2].Name = "Test Column Name";

            ws.Cells["G50"].Value = "Timespan";
            ws.Cells["G51"].Value = new DateTime(new TimeSpan(1, 1, 10).Ticks); //new DateTime(1899, 12, 30, 1, 1, 10);
            ws.Cells["G52"].Value = new DateTime(1899, 12, 30, 2, 3, 10);
            ws.Cells["G53"].Value = new DateTime(1899, 12, 30, 3, 4, 10);
            ws.Cells["G54"].Value = new DateTime(1899, 12, 30, 4, 5, 10);
            
            ws.Cells["G51:G55"].Style.Numberformat.Format = "HH:MM:SS";
            tbl = ws.Tables.Add(ws.Cells["G50:G54"], "");
            tbl.ShowTotal = true;
            tbl.ShowFilter = false;
            tbl.Columns[0].TotalsRowFunction = OfficeOpenXml.Table.RowFunctions.Sum;
        }
        [TestMethod]
        public void CopyTable()
        {
            _pck.Workbook.Worksheets.Copy("File4", "Copied table");
        }
        [TestMethod]
        public void CopyPivotTable()
        {
            _pck.Workbook.Worksheets.Copy("Pivot-Group Date", "Copied Pivottable 1");
            _pck.Workbook.Worksheets.Copy("Pivot-Group Number", "Copied Pivottable 2");
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
        public void ReadBug()
        {
            using (var package = new ExcelPackage(new FileInfo(@"c:\temp\error.xlsx")))
            {
                var fulla = package.Workbook.Worksheets.FirstOrDefault();
                var r= fulla == null ? null : fulla.Cells["a:a"]
                .Where(t => !string.IsNullOrWhiteSpace(t.Text)).Select(cell => cell.Value.ToString())
                .ToList();
            }
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

            ws.Names.Add("Address", ws.Cells["A2:A3"]);
            ws.Cells["Address"].Value = 1;
            ws.Names.AddValue("Value", 5);            
            //ws.Names["Value"].Style.Border.Bottom.Color.SetColor(Color.Black);
            ws.Names.AddFormula("Formula", "Names!A2+Names!A3+Names!Value");
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
            ws.Cells["A1"].LoadFromText(new FileInfo(@"c:\temp\csv\20060927.custom_open_positions.cdf.SPP"), new ExcelTextFormat() { SkipLinesBeginning = 2, SkipLinesEnd=2, TextQualifier='"', DataTypes=new eDataTypes[] {eDataTypes.Number,eDataTypes.String, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number, eDataTypes.String, eDataTypes.Number, eDataTypes.Number, eDataTypes.String, eDataTypes.String, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number}},
                OfficeOpenXml.Table.TableStyles.Medium27, true);
            
            var style = _pck.Workbook.Styles.CreateNamedStyle("RedStyle");
            style.Style.Fill.PatternType=ExcelFillStyle.Solid;
            style.Style.Fill.BackgroundColor.SetColor(Color.Red);
            
            var tbl = ws.Tables[ws.Tables.Count - 1];
            tbl.ShowTotal = true;
            tbl.TotalsRowCellStyle = "RedStyle";
            tbl.HeaderRowCellStyle = "RedStyle";
        }
        [TestMethod]
        public void Merge()
        {
            var ws = _pck.Workbook.Worksheets.Add("Merge");
            ws.Cells["A1:A4"].Merge=true;
            ws.Cells["C1:C4,C8:C12"].Merge=true;
            ws.Cells["D13:E18,G5,U32:U45"].Merge = true;
            ws.Cells["D13:E18,G5,U32:U45"].Style.WrapText = true;
            ws.SetValue(13, 4, "Merged\r\nnew row");
        }
        [TestMethod]
        public void DefaultColWidth()
        {
            var ws = _pck.Workbook.Worksheets.Add("DefColWidth");
            ws.DefaultColWidth = 45;
        }
        [TestMethod]
        public void LoadArray()
        {
            var ws = _pck.Workbook.Worksheets.Add("Loaded Array");
            List<object[]> testArray = new List<object[]>() { new object[] { 3, 4, 5, 6 }, new string[] { "Test1", "test", "5", "6" } };
            ws.Cells["A1"].LoadFromArrays(testArray);
        }
        [TestMethod]
        public void DefColWidthBug()
        {
            ExcelWorkbook book = _pck.Workbook;
            ExcelWorksheet sheet = book.Worksheets.Add("Gebruikers");

            sheet.DefaultColWidth = 25d;
            //sheet.defaultRowHeight = 15d; // needed to make sure the resulting file is valid!

            // Create the header row
            sheet.Cells[1, 1].Value = "Afdeling code";
            sheet.Cells[1, 2].Value = "Afdeling naam";
            sheet.Cells[1, 3].Value = "Voornaam";
            sheet.Cells[1, 4].Value = "Tussenvoegsel";
            sheet.Cells[1, 5].Value = "Achternaam";
            sheet.Cells[1, 6].Value = "Gebruikersnaam";
            sheet.Cells[1, 7].Value = "E-mail adres";
            ExcelRange headerRow = sheet.Cells[1, 1, 1, 7];
            headerRow.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            headerRow.Style.Font.Size = 12;
            headerRow.Style.Font.Bold = true;

            //// Create a context for retrieving the users
            //using (PalauDataContext context = new PalauDataContext())
            //{
            //    int currentRow = 2;

            //    // iterate through all users in the export and add their info
            //    // to the worksheet.
            //    foreach (vw_ExportUser user in
            //      context.vw_ExportUsers
            //      .OrderBy(u => u.DepartmentCode)
            //      .ThenBy(u => u.AspNetUserName))
            //    {
            //        sheet.Cells[currentRow, 1].Value = user.DepartmentCode;
            //        sheet.Cells[currentRow, 2].Value = user.DepartmentName;
            //        sheet.Cells[currentRow, 3].Value = user.UserFirstName;
            //        sheet.Cells[currentRow, 4].Value = user.UserInfix;
            //        sheet.Cells[currentRow, 5].Value = user.UserSurname;
            //        sheet.Cells[currentRow, 6].Value = user.AspNetUserName;
            //        sheet.Cells[currentRow, 7].Value = user.AspNetEmail;

            //        currentRow++;
            //    }
            //}

            // return the filled Excel workbook
          //  return pkg

        }
        [TestMethod]
        public void CloseProblem()
        {
            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Manual Receipts");

            ws.Cells["A1"].Value = " SpaceNeedle Manual Receipt Form";

            using (ExcelRange r = ws.Cells["A1:F1"])
            {
                r.Merge = true;
                r.Style.Font.SetFromFont(new Font("Arial", 18, FontStyle.Italic));
                r.Style.Font.Color.SetColor(Color.DarkRed);
                r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                //r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
            }
            //			ws.Column(1).BestFit = true;
            ws.Column(1).Width = 17;
            ws.Column(5).Width = 20;


            ws.Cells["A2"].Value = "Date Produced";

            ws.Cells["A2"].Style.Font.Bold = true;
            ws.Cells["B2"].Value = DateTime.Now.ToShortDateString();
            ws.Cells["D2"].Value = "Quantity";
            ws.Cells["D2"].Style.Font.Bold = true;
            ws.Cells["E2"].Value = "txt";

            ws.Cells["C4"].Value = "Receipt Number";
            ws.Cells["C4"].Style.WrapText = true;
            ws.Cells["C4"].Style.Font.Bold = true;

            int rowNbr = 5;
            for (int entryNbr = 1; entryNbr <= 1; entryNbr += 1)
            {
                ws.Cells["B" + rowNbr].Value = entryNbr;
                ws.Cells["C" + rowNbr].Value = 1 + entryNbr - 1;
                rowNbr += 1;
            }
            pck.SaveAs(new FileInfo(".\\test.xlsx"));
        }
        [TestMethod]
        public void OpenXlsm()
        {
            ExcelPackage p=new ExcelPackage(new FileInfo("c:\\temp\\cs1.xlsx"));
            int c = p.Workbook.Worksheets.Count;
            p.Save();
        }
        [TestMethod]
        public void Mergebug()
        {
            var xlPackage = new ExcelPackage();
            var xlWorkSheet = xlPackage.Workbook.Worksheets.Add("Test Sheet");
            var Cells = xlWorkSheet.Cells;
            var TitleCell = Cells[1, 1, 1, 3];

            TitleCell.Merge = true;
            TitleCell.Value = "Test Spreadsheet";
            Cells[2, 1].Value = "Test Sub Heading\r\ntest"+(char)22;
            for (int i = 0; i < 256; i++)
            {
                Cells[3, i + 1].Value = (char)i;
            }
            Cells[2, 1].Style.WrapText = true;
            xlWorkSheet.Row(1).Height=50;
            xlPackage.SaveAs(new FileInfo("c:\\temp\\Mergebug.xlsx"));
        }
        [TestMethod]
        public void OpenProblem()
        {
            var xlPackage = new ExcelPackage();
            var ws = xlPackage.Workbook.Worksheets.Add("W1");
            xlPackage.Workbook.Worksheets.Add("W2");

            ws.Cells["A1:A10"].Formula = "W2!A1+C1";
            ws.Cells["B1:B10"].FormulaR1C1 = "W2!R1C1+C1";
            xlPackage.SaveAs(new FileInfo("c:\\temp\\Mergebug.xlsx"));
        }
        [TestMethod]
        public void ProtectionProblem()
        {
            var xlPackage = new ExcelPackage(new FileInfo("c:\\temp\\CovenantsCheckReportTemplate.xlsx"));
            var ws = xlPackage.Workbook.Worksheets.First();
            ws.Protection.SetPassword("Test");
            xlPackage.SaveAs(new FileInfo("c:\\temp\\Mergebug.xlsx"));
        }
        [TestMethod]
        public void Nametest()
        {
            var pck = new ExcelPackage(new FileInfo("c:\\temp\\names.xlsx"));
            var ws = pck.Workbook.Worksheets.First();
            ws.Cells["H37"].Formula = "\"Test\"";
            pck.SaveAs(new FileInfo(@"c:\\temp\\nametest_new.xlsx"));
        }
        [TestMethod]
        public void CreatePivotTable()
        {
            var wsPivot1 = _pck.Workbook.Worksheets.Add("Rows-Data on columns");
            var wsPivot2 = _pck.Workbook.Worksheets.Add("Rows-Data on rows");
            var wsPivot3 = _pck.Workbook.Worksheets.Add("Columns-Data on columns");
            var wsPivot4 = _pck.Workbook.Worksheets.Add("Columns-Data on rows");
            var wsPivot5 = _pck.Workbook.Worksheets.Add("Columns/Rows-Data on columns");
            var wsPivot6 = _pck.Workbook.Worksheets.Add("Columns/Rows-Data on rows");
            var wsPivot7 = _pck.Workbook.Worksheets.Add("Rows/Page-Data on Columns");
            var wsPivot8 = _pck.Workbook.Worksheets.Add("Pivot-Group Date");
            var wsPivot9 = _pck.Workbook.Worksheets.Add("Pivot-Group Number");

            var ws = _pck.Workbook.Worksheets.Add("Data");
            ws.Cells["K1"].Value = "Item";
            ws.Cells["L1"].Value = "Category";
            ws.Cells["M1"].Value = "Stock";
            ws.Cells["N1"].Value = "Price";
            ws.Cells["O1"].Value = "Date for grouping";

            ws.Cells["K2"].Value = "Crowbar";
            ws.Cells["L2"].Value = "Hardware";
            ws.Cells["M2"].Value = 12;
            ws.Cells["N2"].Value = 85.2;
            ws.Cells["O2"].Value = new DateTime(2010, 1, 31);

            ws.Cells["K3"].Value = "Crowbar";
            ws.Cells["L3"].Value = "Hardware";
            ws.Cells["M3"].Value = 15;
            ws.Cells["N3"].Value = 12.2;
            ws.Cells["O3"].Value = new DateTime(2010, 2, 28);

            ws.Cells["K4"].Value = "Hammer";
            ws.Cells["L4"].Value = "Hardware";
            ws.Cells["M4"].Value = 550;
            ws.Cells["N4"].Value = 72.7;
            ws.Cells["O4"].Value = new DateTime(2010, 3, 31);

            ws.Cells["K5"].Value = "Hammer";
            ws.Cells["L5"].Value = "Hardware";
            ws.Cells["M5"].Value = 120;
            ws.Cells["N5"].Value = 11.3;
            ws.Cells["O5"].Value = new DateTime(2010, 4, 30);

            ws.Cells["K6"].Value = "Crowbar";
            ws.Cells["L6"].Value = "Hardware";
            ws.Cells["M6"].Value = 120;
            ws.Cells["N6"].Value = 173.2;
            ws.Cells["O6"].Value = new DateTime(2010, 5, 31);

            ws.Cells["K7"].Value = "Hammer";
            ws.Cells["L7"].Value = "Hardware";
            ws.Cells["M7"].Value = 1;
            ws.Cells["N7"].Value = 4.2;
            ws.Cells["O7"].Value = new DateTime(2010, 6, 30);

            ws.Cells["K8"].Value = "Saw";
            ws.Cells["L8"].Value = "Hardware";
            ws.Cells["M8"].Value = 4;
            ws.Cells["N8"].Value = 33.12;
            ws.Cells["O8"].Value = new DateTime(2010, 6, 28);

            ws.Cells["K9"].Value = "Screwdriver";
            ws.Cells["L9"].Value = "Hardware";
            ws.Cells["M9"].Value = 1200;
            ws.Cells["N9"].Value = 45.2;
            ws.Cells["O9"].Value = new DateTime(2010, 8, 31);

            ws.Cells["K10"].Value = "Apple";
            ws.Cells["L10"].Value = "Groceries";
            ws.Cells["M10"].Value = 807;
            ws.Cells["N10"].Value = 1.2;
            ws.Cells["O10"].Value = new DateTime(2010, 9, 30);

            ws.Cells["K11"].Value = "Butter";
            ws.Cells["L11"].Value = "Groceries";
            ws.Cells["M11"].Value = 52;
            ws.Cells["N11"].Value = 7.2;
            ws.Cells["O11"].Value = new DateTime(2010, 10, 31);
            ws.Cells["O2:O11"].Style.Numberformat.Format = "yyyy-MM-dd";

            var pt = wsPivot1.PivotTables.Add(wsPivot1.Cells["A1"], ws.Cells["K1:N11"], "Pivottable1");
            pt.GrandTotalCaption = "Total amount";
            pt.RowFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataFields[0].Function = DataFieldFunctions.Product;
            pt.DataOnRows = false;

            pt = wsPivot2.PivotTables.Add(wsPivot2.Cells["A1"], ws.Cells["K1:N11"], "Pivottable2");
            pt.RowFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataFields[0].Function = DataFieldFunctions.Average;
            pt.DataOnRows = true;

            pt = wsPivot3.PivotTables.Add(wsPivot3.Cells["A1"], ws.Cells["K1:N11"], "Pivottable3");
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.ColumnFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = false;

            pt = wsPivot4.PivotTables.Add(wsPivot4.Cells["A1"], ws.Cells["K1:N11"], "Pivottable4");
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.ColumnFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = true;

            pt = wsPivot5.PivotTables.Add(wsPivot5.Cells["A1"], ws.Cells["K1:N11"], "Pivottable5");
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = false;

            pt = wsPivot6.PivotTables.Add(wsPivot6.Cells["A1"], ws.Cells["K1:N11"], "Pivottable6");
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = true;
            wsPivot6.Drawings.AddChart("Pivotchart6",OfficeOpenXml.Drawing.Chart.eChartType.BarStacked3D, pt);

            pt = wsPivot7.PivotTables.Add(wsPivot7.Cells["A3"], ws.Cells["K1:N11"], "Pivottable7");
            pt.PageFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[0]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = false;
            
            pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Sum | eSubTotalFunctions.Max;
            Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.Sum | eSubTotalFunctions.Max);

            pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Sum | eSubTotalFunctions.Product | eSubTotalFunctions.StdDevP;
            Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.Sum | eSubTotalFunctions.Product | eSubTotalFunctions.StdDevP);

            pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.None;
            Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.None);

            pt.Fields[0].SubTotalFunctions = eSubTotalFunctions.Default;
            Assert.AreEqual(pt.Fields[0].SubTotalFunctions, eSubTotalFunctions.Default);

            pt.Fields[0].Sort = eSortType.Descending;
            pt.TableStyle = OfficeOpenXml.Table.TableStyles.Medium14;

            pt = wsPivot8.PivotTables.Add(wsPivot8.Cells["A3"], ws.Cells["K1:O11"], "Pivottable8");
            pt.RowFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[4]);
            pt.Fields[4].AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months | eDateGroupBy.Days | eDateGroupBy.Quarters, new DateTime(2010, 01, 31), new DateTime(2010, 11, 30));
            pt.RowHeaderCaption = "År";
            pt.Fields[4].Name = "Dag";
            pt.Fields[5].Name = "Månad";
            pt.Fields[6].Name = "Kvartal";
            pt.GrandTotalCaption = "Totalt";
           
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = true;

            pt = wsPivot9.PivotTables.Add(wsPivot9.Cells["A3"], ws.Cells["K1:N11"], "Pivottable9");
            pt.PageFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[3]);            
            pt.RowFields[0].AddNumericGrouping(-3.3, 5.5, 4.0);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = false;
            pt.TableStyle = OfficeOpenXml.Table.TableStyles.Medium14;

            pt = wsPivot8.PivotTables.Add(wsPivot8.Cells["H3"], ws.Cells["K1:O11"], "Pivottable10");
            pt.RowFields.Add(pt.Fields[1]);
            pt.RowFields.Add(pt.Fields[4]);
            pt.Fields[4].AddDateGrouping(7, new DateTime(2010, 01, 31), new DateTime(2010, 11, 30));
            pt.RowHeaderCaption = "Veckor";
            pt.GrandTotalCaption = "Totalt";

            pt = wsPivot8.PivotTables.Add(wsPivot8.Cells["A60"], ws.Cells["K1:O11"], "Pivottable11");
            pt.RowFields.Add(pt.Fields["Category"]);
            pt.RowFields.Add(pt.Fields["Item"]);
            pt.RowFields.Add(pt.Fields["Date for grouping"]);

            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.DataOnRows = true;
        }
        [TestMethod]
        public void ReadPivotTable()
        {
            ExcelPackage pck = new ExcelPackage(new FileInfo(@"c:\temp\pivot\pivotforread.xlsx"));

            var pivot1 = pck.Workbook.Worksheets[2].PivotTables[0];

            Assert.AreEqual(pivot1.Fields.Count, 24);
            Assert.AreEqual(pivot1.RowFields.Count, 3);
            Assert.AreEqual(pivot1.DataFields.Count, 7);
            Assert.AreEqual(pivot1.ColumnFields.Count, 0);

            Assert.AreEqual(pivot1.DataFields[1].Name, "Sum of n3");
            Assert.AreEqual(pivot1.Fields[2].Sort, eSortType.Ascending);

            Assert.AreEqual(pivot1.DataOnRows, false);

            var pivot2 = pck.Workbook.Worksheets[2].PivotTables[0];
            var pivot3 = pck.Workbook.Worksheets[3].PivotTables[0];

            var pivot4 = pck.Workbook.Worksheets[4].PivotTables[0];
            var pivot5 = pck.Workbook.Worksheets[5].PivotTables[0];
            pivot5.CacheDefinition.SourceRange = pck.Workbook.Worksheets[1].Cells["Q1:X300"];
            
            var pivot6 = pck.Workbook.Worksheets[6].PivotTables[0];
            
            pck.Workbook.Worksheets[6].Drawings.AddChart("chart1", OfficeOpenXml.Drawing.Chart.eChartType.ColumnStacked3D, pivot6);

            pck.SaveAs(new FileInfo(@"c:\temp\pivot\pivotforread_new.xlsx"));
        }
        [TestMethod]
        public void SetBackground()
        {
            var ws = _pck.Workbook.Worksheets.Add("backimg");

            ws.BackgroundImage.Image = Properties.Resources.Test1;
            ws = _pck.Workbook.Worksheets.Add("backimg2");
            ws.BackgroundImage.SetFromFile(new FileInfo(@"C:\Program Files (x86)\Microsoft Office\CLIPART\PUB60COR\WHIRL1.WMF"));
        }
        [TestMethod]
        public void SetHeaderFooterImage()
        {
            var ws = _pck.Workbook.Worksheets.Add("HeaderImage");
            ws.HeaderFooter.OddHeader.CenteredText = "Before ";
            var img=ws.HeaderFooter.OddHeader.InsertPicture(Properties.Resources.Test1, PictureAlignment.Centered);
            img.Title = "Renamed Image";
            img.GrayScale = true;
            img.BiLevel = true;
            img.Gain = .5;
            img.Gamma = .35;

            Assert.AreEqual(img.Width, 426);
            img.Width /= 4;
            Assert.AreEqual(img.Height, 49.5);
            img.Height /= 4;
            Assert.AreEqual(img.Left, 0);
            Assert.AreEqual(img.Top, 0);
            ws.HeaderFooter.OddHeader.CenteredText += " After";


            img = ws.HeaderFooter.EvenFooter.InsertPicture(new FileInfo(@"C:\Program Files (x86)\Microsoft Office\CLIPART\PUB60COR\WHIRL1.WMF"), PictureAlignment.Left);
            img.Title = "DiskFile";

            img = ws.HeaderFooter.FirstHeader.InsertPicture(new FileInfo(@"C:\Program Files (x86)\Microsoft Office\CLIPART\PUB60COR\WING1.WMF"), PictureAlignment.Right);
            img.Title = "DiskFile2";
            ws.Cells["A1:A400"].Value = 1;

            _pck.Workbook.Worksheets.Copy(ws.Name, "Copied HeaderImage");
        }
        [TestMethod]
        public void NamedStyles()
        {
            var wsSheet = _pck.Workbook.Worksheets.Add("NamedStyles");

            var firstNamedStyle =
				_pck.Workbook.Styles.CreateNamedStyle("templateFirst");
				
            var s=firstNamedStyle.Style;

            s.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
            s.VerticalAlignment = ExcelVerticalAlignment.Center;

            var secondNamedStyle = _pck.Workbook.Styles.CreateNamedStyle("first", firstNamedStyle.Style).Style;
            secondNamedStyle.Font.Bold = true;
            secondNamedStyle.Font.SetFromFont(new Font("Arial Black", 8));
            secondNamedStyle.Border.Bottom.Style = ExcelBorderStyle.Medium;
            secondNamedStyle.Border.Left.Style = ExcelBorderStyle.Medium;

            wsSheet.Cells["B2"].Value = "Text Center";
            wsSheet.Cells["B2"].StyleName = "first";
        }
        [TestMethod]
        public void StyleFill()
        {
            var ws = _pck.Workbook.Worksheets.Add("Fills");
            ws.Cells["A1:C3"].Style.Fill.Gradient.Type = ExcelFillGradientType.Linear;
            ws.Cells["A1:C3"].Style.Fill.Gradient.Color1.SetColor(Color.Red);
            ws.Cells["A1:C3"].Style.Fill.Gradient.Color2.SetColor(Color.Blue);

            ws.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.MediumGray;
            ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.ForestGreen);
            var r=ws.Cells["A2:A3"];
            r.Style.Fill.Gradient.Type = ExcelFillGradientType.Path;
            r.Style.Fill.Gradient.Left = 0.7;
            r.Style.Fill.Gradient.Right = 0.7;
            r.Style.Fill.Gradient.Top = 0.7;
            r.Style.Fill.Gradient.Bottom = 0.7;
            
            ws.Cells[4,1,4,360].Style.Fill.Gradient.Type = ExcelFillGradientType.Path;

            for (double col = 1; col < 360; col++)
            {                
                r = ws.Cells[4, Convert.ToInt32(col)];
                r.Style.Fill.Gradient.Degree = col;
                r.Style.Fill.Gradient.Left = col / 360;
                r.Style.Fill.Gradient.Right = col / 360;
                r.Style.Fill.Gradient.Top = col / 360;
                r.Style.Fill.Gradient.Bottom = col / 360;
            }
            r = ws.Cells["A5"];
            r.Style.Fill.Gradient.Left = .50;

            ws = _pck.Workbook.Worksheets.Add("FullFills");
            ws.Cells.Style.Fill.Gradient.Left = 0.25;
            ws.Cells["A1"].Value = "test";
            ws.Cells["A1"].RichText.Add("Test rt");
            ws.Cells.AutoFilter=true;
            Assert.AreNotEqual(ws.Cells["A1:D5"].Value, null);
        }
        [TestMethod]
        public void LoadFromCollection_Strings()
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Sheet 1");
                IEnumerable<String> list = new List<String>() { "a", "b", "c", "d", "e", "f", "g" };

                ws.Cells["A1"].LoadFromCollection(list, false);

                Assert.AreEqual("b", ws.Cells["A2"].Value);
            }
        }

        [TestMethod]
        public void LoadFromCollection_Objects()
        {
            using (ExcelPackage pck = new ExcelPackage())
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Sheet 1");
                IEnumerable<Obj> list = new[] { new Obj("a"), new Obj("b"), new Obj("c"), new Obj("d"), new Obj("e"), new Obj("f"), new Obj("g") };

                ws.Cells["A1"].LoadFromCollection(list, false);

                Assert.AreEqual("b", ws.Cells["A2"].Value);
            }
        }

        class Obj
        {
            public Obj(String name)
            {
                Name = name;
            }

            public String Name { get; set; }
        }
    }
}
