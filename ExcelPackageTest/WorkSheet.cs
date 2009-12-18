
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Drawing;
using System.Drawing;

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
            _pck.Save();
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

            ws.Cells["X19"].Value = 210;
            ws.Cells["X20"].Value = 212;
            ws.Cells["X21"].Value = 221;
            ws.Cells["X22"].Value = 123;
            ws.Cells["X23"].Value = 135;
            ws.Cells["X24"].Value = 134;

            ExcelPicture pic = ws.Drawings.AddPicture("Pic1", Image.FromFile(@"C:\dev\AIM\applications\trunk\PropMIS\VB\User\IMAGES\Alecta_rect.jpg"));
            pic.SetPositionPixels(150, 140);
            //pic.From.Column = 15;
            //pic.From.Row = 32;
            //pic.To.Column = 18;
            //pic.To.Row = 49;
        }
    }
}
