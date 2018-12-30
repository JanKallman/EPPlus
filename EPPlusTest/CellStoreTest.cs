using System;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace EPPlusTest
{
    [TestClass]
    public class CellStoreTest : TestBase
    {
        
        [TestMethod]
        public void Insert1()
        {
            var ws=_pck.Workbook.Worksheets.Add("Insert1");
            LoadData(ws);

            ws.InsertRow(2, 1000);
            Assert.AreEqual(ws.GetValue(1002,1),"1,0");
            ws.InsertRow(1003, 1000);
            Assert.AreEqual(ws.GetValue(2003, 1), "2,0");
            ws.InsertRow(2004, 1000);
            Assert.AreEqual(ws.GetValue(3004, 1), "3,0");
            ws.InsertRow(2006, 1000);
            Assert.AreEqual(ws.GetValue(4005, 1), "4,0");
            ws.InsertRow(4500, 500);
            Assert.AreEqual(ws.GetValue(5000, 1), "499,0");

            ws.InsertRow(1, 1);
            Assert.AreEqual(ws.GetValue(1003, 1), "1,0");
            Assert.AreEqual(ws.GetValue(5001, 1), "499,0");

            ws.InsertRow(1, 15);
            Assert.AreEqual(ws.GetValue(4020, 1), "3,0");
            Assert.AreEqual(ws.GetValue(5016, 1), "499,0");
   
        }
        [TestMethod]
        public void Insert2()
        {
            var ws = _pck.Workbook.Worksheets.Add("Insert2-1");
            LoadData(ws);

            for (int i = 0; i < 32; i++)
            {
                ws.InsertRow(1, 1);
            }
            Assert.AreEqual(ws.GetValue(33,1),"0,0");

            ws = _pck.Workbook.Worksheets.Add("Insert2-2");
            LoadData(ws);

            for (int i = 0; i < 32; i++)
            {
                ws.InsertRow(15, 1);
            }
            Assert.AreEqual(ws.GetValue(1, 1), "0,0");
            Assert.AreEqual(ws.GetValue(47, 1), "14,0");
        }
        [TestMethod]
        public void Insert3()
        {
            var ws = _pck.Workbook.Worksheets.Add("Insert3");
            LoadData(ws);

            for (int i = 0; i < 500; i+=4)
            {
                ws.InsertRow(i+1, 2);
            }
        }

        [TestMethod]
        public void InsertRandomTest()
        {
            var ws = _pck.Workbook.Worksheets.Add("Insert4-1");
            
            LoadData(ws, 5000);

            for (int i = 5000; i > 0; i-=2)
            {
                ws.InsertRow(i, 1);
            }
        }
        [TestMethod]
        public void EnumCellstore()
        {
            var ws = _pck.Workbook.Worksheets.Add("enum");

            LoadData(ws, 5000);

            var o = new CellsStoreEnumerator<ExcelCoreValue>(ws._values, 2, 1, 5, 3);
            foreach (var i in o)
            {
                Console.WriteLine(i);
            }
        }
        [TestMethod]
        public void DeleteCells()
        {
            var ws = _pck.Workbook.Worksheets.Add("Delete");
            LoadData(ws, 5000);

            ws.DeleteRow(2, 2);
            Assert.AreEqual("3,0",ws.GetValue(2,1));
            ws.DeleteRow(10, 10);
            Assert.AreEqual("21,0", ws.GetValue(10, 1));
            ws.DeleteRow(50, 40);
            Assert.AreEqual("101,0", ws.GetValue(50, 1));
            ws.DeleteRow(100, 100);
            Assert.AreEqual("251,0", ws.GetValue(100, 1));
            ws.DeleteRow(1, 31);
            Assert.AreEqual("43,0", ws.GetValue(1, 1));
        }
        [TestMethod]
        public void DeleteCellsFirst()
        {
            var ws = _pck.Workbook.Worksheets.Add("DeleteFirst");
            LoadData(ws, 5000);

            ws.DeleteRow(32, 30);
            for (int i = 1; i < 50; i++)
            {
                ws.DeleteRow(1,1);
            }
        }
        [TestMethod]
        public void DeleteInsert()
        {
            var ws = _pck.Workbook.Worksheets.Add("DeleteInsert");
            LoadData(ws, 5000);

            ws.DeleteRow(2, 33);
            ws.InsertRow(2, 38);

            for (int i = 0; i < 33; i++)
            {
                ws.SetValue(i + 2,1, i + 2);
            }
        }
        private void LoadData(ExcelWorksheet ws)
        {
            LoadData(ws, 1000);
        }
        private void LoadData(ExcelWorksheet ws, int rows, int cols=1, bool isNumeric = false)
        {
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    if (isNumeric)
                        ws.SetValue(r + 1, c + 1, r + c);
                    else
                        ws.SetValue(r+1, c+1, r.ToString()+","+c.ToString());
                }
            }
        }
        [TestMethod]
        public void FillInsertTest()
        {
            var ws = _pck.Workbook.Worksheets.Add("FillInsert");

            LoadData(ws, 500);

            var r=1;
            for(int i=1;i<=500;i++)
            {
                ws.InsertRow(r,i);
                Assert.AreEqual((i-1).ToString()+",0", ws.GetValue(r+i,1).ToString());
                r+=i+1;
            }
        }
        [TestMethod]
        public void CopyCellsTest()
        {
            var ws = _pck.Workbook.Worksheets.Add("CopyCells");

            LoadData(ws, 100, isNumeric: true);
            ws.Cells["B1"].Formula = "SUM(A1:A500)";
            ws.Calculate();
            ws.Cells["B1"].Copy(ws.Cells["C1"]);
            ws.Cells["B1"].Copy(ws.Cells["D1"], ExcelRangeCopyOptionFlags.ExcludeFormulas);

            Assert.AreEqual(ws.Cells["B1"].Value, ws.Cells["C1"].Value);
            Assert.AreEqual("SUM(B1:B500)", ws.Cells["C1"].Formula);

            Assert.AreEqual(ws.Cells["B1"].Value, ws.Cells["D1"].Value);
            Assert.AreNotEqual(ws.Cells["B1"].Formula, ws.Cells["D1"].Formula);
        }
        [TestMethod]
        public void Issues351()
        {
            using (var package = new ExcelPackage())
            {
                // Arrange
                var worksheet = package.Workbook.Worksheets.Add("Test");
                worksheet.Cells[1, 1].Value = "A";                      // If you remove this "anchor", the problem doesn't happen.
                worksheet.Cells[1026, 1].Value = "B";
                worksheet.Cells[1026, 2].Value = "B";
                var range = worksheet.Row(1026);
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 0));

                // Act - This should shift the whole row 1026 down 1
                worksheet.InsertRow(1024, 1);

                // Assert - This value should be null, instead it's "B"
                Assert.IsNull(worksheet.Cells[1025, 1].Value);
            }
        }
    }
}
