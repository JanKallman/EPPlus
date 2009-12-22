using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.IO;
using System.Drawing;

namespace ExcelPackageTest
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class DrawingTest
    {
        public DrawingTest()
        {
        }
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

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
         [ClassInitialize()]
         public static void MyClassInitialize(TestContext testContext) 
         {
             if (Directory.Exists("Test"))
             {
                 Directory.Delete("Test", true);
             }
             Directory.CreateDirectory(string.Format("Test"));
             _pck = new ExcelPackage(new FileInfo("Test\\Drawing.xlsx"));         
         }
        
        // Use ClassCleanup to run code after all tests in a class have run
         [ClassCleanup()]
         public static void MyClassCleanup() 
         {
             _pck.Save();
             _pck = null;         
         }
        
        // //Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() 
        // {
        // }
        
        //// Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() 
        // {
        // }
        
        #endregion
         [TestMethod]
         public void Picture()
         {
             var ws = _pck.Workbook.Worksheets.Add("Picture");
             ExcelPicture pic = ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);

             pic = ws.Drawings.AddPicture("Pic2", Properties.Resources.Test1);
             pic.SetPosition(150, 200);


             pic = ws.Drawings.AddPicture("Pic2", Properties.Resources.Test1);
             pic.SetPosition(400, 200);
             pic.SetSize(150);
         }

        [TestMethod]
        public void BarChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("BarChart");            
            var chrt = ws.Drawings.AddChart("barChart", eChartType.xlBarClustered) as ExcelBarChart;
            chrt.SetPosition(50, 50);
            chrt.SetSize(800, 300);
            AddTestSerie(ws, chrt);

            Assert.IsTrue(chrt.ChartType == eChartType.xlBarClustered, "Invalid Charttype");
            Assert.IsTrue(chrt.Direction == eDirection.Bar, "Invalid Bardirection");
            Assert.IsTrue(chrt.Grouping == eGrouping.Clustered, "Invalid Grouping");
            Assert.IsTrue(chrt.Shape == eShape.Box, "Invalid Shape");
        }

        private static void AddTestSerie(ExcelWorksheet ws, ExcelChart chrt)
        {
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
            
            chrt.Series.Add("U19:U24", "V19:V24");
        }
        [TestMethod]
        public void PieChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("PieChart");
            var chrt = ws.Drawings.AddChart("pieChart", eChartType.xlPie) as ExcelPieChart;
            
            AddTestSerie(ws, chrt);

            chrt.To.Row = 25;
            chrt.To.Column = 12;

            chrt.Series[0].DataLabel.Position = eLabelPosition.Center;
            Assert.IsTrue(chrt.ChartType == eChartType.xlPie, "Invalid Charttype");
            Assert.IsTrue(chrt.VaryColors);

        }
        [TestMethod]
        public void PieChart3D()
        {
            var ws = _pck.Workbook.Worksheets.Add("PieChart3d");
            var chrt = ws.Drawings.AddChart("pieChart3d", eChartType.xl3DPie) as ExcelPieChart;
            AddTestSerie(ws, chrt);

            chrt.To.Row = 25;
            chrt.To.Column = 12;

            chrt.Series[0].DataLabel.Position = eLabelPosition.Center;
            Assert.IsTrue(chrt.ChartType == eChartType.xl3DPie, "Invalid Charttype");
            Assert.IsTrue(chrt.VaryColors);

        }
        [TestMethod]
        public void Drawings()
        {
            var ws = _pck.Workbook.Worksheets.Add("Shapes");
            int y=100, i=1;
            foreach(eShapeStyle style in Enum.GetValues(typeof(eShapeStyle)))
            {
                var shape = ws.Drawings.AddShape("shape"+i.ToString(), style);
                shape.SetPosition(y, 100);
                shape.SetSize(300, 300);
                y += 400;
                shape.Text = style.ToString();
                i++;
            }

            (ws.Drawings["shape1"] as ExcelShape).TextAnchoring = eTextAnchoringType.Top;
            (ws.Drawings["shape2"] as ExcelShape).TextVertical = eTextVerticalType.Vertical;

            (ws.Drawings["shape3"] as ExcelShape).TextAnchoring=eTextAnchoringType.Bottom;
            (ws.Drawings["shape3"] as ExcelShape).TextAnchoringControl=true ;

            (ws.Drawings["shape4"] as ExcelShape).TextVertical = eTextVerticalType.Vertical270;
            (ws.Drawings["shape4"] as ExcelShape).TextAnchoring = eTextAnchoringType.Top;

            (ws.Drawings["shape5"] as ExcelShape).Fill.Style=eFillStyle.SolidFill;
            (ws.Drawings["shape5"] as ExcelShape).Fill.Color=Color.Red;
            (ws.Drawings["shape5"] as ExcelShape).Fill.Transparancy = 50;

            (ws.Drawings["shape6"] as ExcelShape).Fill.Style = eFillStyle.NoFill;

            (ws.Drawings["shape7"] as ExcelShape).Fill.Style = eFillStyle.SolidFill;
            (ws.Drawings["shape7"] as ExcelShape).Fill.Color=Color.Gray;
            (ws.Drawings["shape7"] as ExcelShape).Line.Fill.Style=eFillStyle.SolidFill;
            (ws.Drawings["shape7"] as ExcelShape).Line.Fill.Color=Color.Black;
            (ws.Drawings["shape7"] as ExcelShape).Line.Fill.Transparancy=43;
        }
    }
}
