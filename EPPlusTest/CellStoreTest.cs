using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace EPPlusTest
{
    [TestClass]
    public class CellStoreTest
    {
        ExcelPackage _package;
        [TestInitialize]
        public void Init()
        {
            _package=new ExcelPackage();
        }
        [TestCleanup]
        public void CleanUp()
        {
            if(!Directory.Exists("Test"))
            {
                Directory.CreateDirectory("test");
            }
            _package.SaveAs(new FileInfo("test\\Insert.xlsx"));
        }
        [TestMethod]
        public void Insert1()
        {
            var ws=_package.Workbook.Worksheets.Add("Insert1");
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
            var ws = _package.Workbook.Worksheets.Add("Insert2-1");
            LoadData(ws);

            for (int i = 0; i < 32; i++)
            {
                ws.InsertRow(1, 1);
            }
            Assert.AreEqual(ws.GetValue(33,1),"0,0");
            
            ws = _package.Workbook.Worksheets.Add("Insert2-2");
            LoadData(ws);

            for (int i = 0; i < 32; i++)
            {
                ws.InsertRow(15, 1);
            }
            Assert.AreEqual(ws.GetValue(1, 1), "0,0");
            Assert.AreEqual(ws.GetValue(34, 1), "1,0");
        }
        [TestMethod]
        public void Insert3()
        {
            var ws = _package.Workbook.Worksheets.Add("Insert3");
            LoadData(ws);

            for (int i = 0; i < 500; i+=4)
            {
                ws.InsertRow(i+1, 2);
            }
        }

        [TestMethod]
        public void InsertRandomTest()
        {
            var ws = _package.Workbook.Worksheets.Add("Insert4-1");
            
            LoadData(ws, 5000);

            for (int i = 5000; i > 0; i-=2)
            {
                ws.InsertRow(i, 1);
            }
        }
        
        private void LoadData(ExcelWorksheet ws)
        {
            LoadData(ws, 1000);
        }
        private void LoadData(ExcelWorksheet ws, int rows)
        {
            for (int r = 0; r < rows; r++)
            {
                for(int c=0;c<1;c++)
                {
                    ws.SetValue(r+1, c+1, r.ToString()+","+c.ToString());
                }
            }
        }
    }
}
