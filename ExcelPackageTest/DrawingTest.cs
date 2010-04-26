using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.IO;
using System.Drawing;
using OfficeOpenXml.Style;
namespace ExcelPackageTest
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class DrawingTest
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
             pic.Border.LineStyle = eLineStyle.Solid;
             pic.Border.Fill.Color = Color.DarkCyan;
             pic.Fill.Style=eFillStyle.SolidFill;
             pic.Fill.Color = Color.White;
             pic.Fill.Transparancy = 50;

             pic = ws.Drawings.AddPicture("Pic3", Properties.Resources.Test1);
             pic.SetPosition(400, 200);
             pic.SetSize(150);

             pic = ws.Drawings.AddPicture("Pic4", new FileInfo(@"C:\Program Files (x86)\Microsoft Office\CLIPART\PUB60COR\WHIRL1.WMF"));
             pic = ws.Drawings.AddPicture("Pic5", new FileInfo(@"C:\Program Files (x86)\Microsoft Office\CLIPART\PUB60COR\AG00004_.GIF"));
             pic.SetPosition(400, 200);
             pic.SetSize(150);


             pic = ws.Drawings.AddPicture("Pic6", new FileInfo(@"C:\Program Files (x86)\Microsoft Office\CLIPART\PUB60COR\AG00004_.GIF"));
             pic.SetPosition(400, 400);
             pic.SetSize(100);
         }

        [TestMethod]
        public void BarChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("BarChart");            
            var chrt = ws.Drawings.AddChart("barChart", eChartType.BarClustered) as ExcelBarChart;
            chrt.SetPosition(50, 50);
            chrt.SetSize(800, 300);
            AddTestSerie(ws, chrt);

            Assert.IsTrue(chrt.ChartType == eChartType.BarClustered, "Invalid Charttype");
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

            chrt.Series.Add("V19:V24", "U19:U24");
        }
        [TestMethod]
        public void PieChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("PieChart");
            var chrt = ws.Drawings.AddChart("pieChart", eChartType.Pie) as ExcelPieChart;
            
            AddTestSerie(ws, chrt);

            chrt.To.Row = 25;
            chrt.To.Column = 12;

            chrt.DataLabel.ShowPercent = true;
            chrt.Legend.Font.Color = Color.SteelBlue;
            chrt.Title.Border.Fill.Style = eFillStyle.SolidFill;
            Assert.IsTrue(chrt.ChartType == eChartType.Pie, "Invalid Charttype");
            Assert.IsTrue(chrt.VaryColors);

        }
        [TestMethod]
        public void PieChart3D()
        {
            var ws = _pck.Workbook.Worksheets.Add("PieChart3d");
            var chrt = ws.Drawings.AddChart("pieChart3d", eChartType.Pie3D) as ExcelPieChart;
            AddTestSerie(ws, chrt);

            chrt.To.Row = 25;
            chrt.To.Column = 12;

            chrt.DataLabel.ShowValue = true;
            Assert.IsTrue(chrt.ChartType == eChartType.Pie3D, "Invalid Charttype");
            Assert.IsTrue(chrt.VaryColors);

        }
        [TestMethod]
        public void Scatter()
        {
            var ws = _pck.Workbook.Worksheets.Add("Scatter");
            var chrt = ws.Drawings.AddChart("ScatterChart1", eChartType.XYScatterSmoothNoMarkers) as ExcelScatterChart;
            AddTestSerie(ws, chrt);
           // chrt.Series[0].Marker = eMarkerStyle.Diamond;
            chrt.To.Row = 23;
            chrt.To.Column = 12;
            chrt.Title.Text = "Header Text";
            chrt.Title.Fill.Style = eFillStyle.SolidFill;
            chrt.Title.Fill.Color = Color.LightBlue;
            chrt.Title.Fill.Transparancy = 50;
            ExcelScatterChartSerie ser = chrt.Series[0] as ExcelScatterChartSerie;
            ser.DataLabel.Position = eLabelPosition.Center;
            ser.DataLabel.ShowValue = true;
            ser.DataLabel.ShowCategory = true;
            ser.DataLabel.Fill.Color = Color.BlueViolet;
            ser.DataLabel.Font.Color = Color.White;
            ser.DataLabel.Font.Italic = true;
            ser.DataLabel.Font.SetFromFont(new Font("bookman old style", 8));
            Assert.IsTrue(chrt.ChartType == eChartType.XYScatterSmoothNoMarkers, "Invalid Charttype");
            chrt.Series[0].Header = "Test serie";
            chrt = ws.Drawings.AddChart("ScatterChart2", eChartType.XYScatterSmooth) as ExcelScatterChart;
            chrt.Series.Add("U19:U24", "V19:V24");

            chrt.From.Column = 0;
            chrt.From.Row=25;
            chrt.To.Row = 53;
            chrt.To.Column = 12;
            
            ////chrt.Series[0].DataLabel.Position = eLabelPosition.Center;
            //Assert.IsTrue(chrt.ChartType == eChartType.XYScatter, "Invalid Charttype");

        }
        [TestMethod]
        public void Pyramid()
        {
            var ws = _pck.Workbook.Worksheets.Add("Pyramid");
            var chrt = ws.Drawings.AddChart("Pyramid1", eChartType.PyramidCol) as ExcelBarChart;
            AddTestSerie(ws, chrt);
            // chrt.Series[0].Marker = eMarkerStyle.Diamond;
            chrt.To.Row = 23;
            chrt.To.Column = 12;
            chrt.Title.Text = "Header Text";
            chrt.Title.Fill.Style= eFillStyle.SolidFill;
            chrt.Title.Fill.Color = Color.DarkBlue;
            chrt.DataLabel.ShowValue = true;
            //chrt.DataLabel.ShowSeriesName = true;
            //chrt.DataLabel.Separator = ",";
            chrt.Border.LineCap = eLineCap.Round;            
            chrt.Border.LineStyle = eLineStyle.LongDashDotDot;
            chrt.Border.Fill.Style = eFillStyle.SolidFill;
            chrt.Border.Fill.Color = Color.Blue;

            chrt.Fill.Color = Color.LightCyan;
            chrt.PlotArea.Fill.Color = Color.White;
            chrt.PlotArea.Border.Fill.Style = eFillStyle.SolidFill;
            chrt.PlotArea.Border.Fill.Color = Color.Beige;
            chrt.PlotArea.Border.LineStyle = eLineStyle.LongDash;

            chrt.Legend.Fill.Color = Color.Aquamarine;
            chrt.Axis[0].Fill.Style = eFillStyle.SolidFill;
            chrt.Axis[0].Fill.Color = Color.Black;
            chrt.Axis[0].Font.Color = Color.White;

            chrt.Axis[1].Fill.Style = eFillStyle.SolidFill;
            chrt.Axis[1].Fill.Color = Color.LightSlateGray;
            chrt.Axis[1].Font.Color = Color.DarkRed;

            chrt.DataLabel.Font.Bold = true;
            chrt.DataLabel.Fill.Color = Color.LightBlue;
            chrt.DataLabel.Border.Fill.Style=eFillStyle.SolidFill;
            chrt.DataLabel.Border.Fill.Color=Color.Black;
            chrt.DataLabel.Border.LineStyle = eLineStyle.Solid;
        }
        [TestMethod]
        public void Cone()
        {
            var ws = _pck.Workbook.Worksheets.Add("Cone");
            var chrt = ws.Drawings.AddChart("Cone1", eChartType.ConeBarClustered) as ExcelBarChart;
            AddTestSerie(ws, chrt);
            chrt.SetSize(200);
            chrt.Title.Text = "Cone bar";
            chrt.Series[0].Header = "Serie 1";
        }
        [TestMethod]
        public void Column()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column");
            var chrt = ws.Drawings.AddChart("Column1", eChartType.ColumnClustered) as ExcelBarChart;
            AddTestSerie(ws, chrt);
            chrt.SetSize(200);
            chrt.Title.Text = "Column";
            chrt.Series[0].Header = "Serie 1";
        }
        [TestMethod]
        public void Dougnut()
        {
            var ws = _pck.Workbook.Worksheets.Add("Dougnut");
            var chrt = ws.Drawings.AddChart("Dougnut1", eChartType.DoughnutExploded) as ExcelDoughnutChart;
            AddTestSerie(ws, chrt);
            chrt.SetSize(200);
            chrt.Title.Text = "Doughnut Exploded";
            chrt.Series[0].Header = "Serie 1";
        }
        [TestMethod]
        public void Line()
        {
            var ws = _pck.Workbook.Worksheets.Add("Line");
            var chrt = ws.Drawings.AddChart("Line1", eChartType.Line3D);
            AddTestSerie(ws, chrt);
            chrt.SetSize(150);
            chrt.Title.Text = "Line 3D";
            chrt.Series[0].Header = "Line serie 1";
            chrt.Fill.Color = Color.LightSteelBlue;
            chrt.Border.LineStyle = eLineStyle.Dot;
            chrt.Border.Fill.Color=Color.Black;

            chrt.Legend.Font.Color = Color.Red;
            chrt.Legend.Font.Strike = eStrikeType.Double;
            chrt.Title.Font.Color = Color.DarkGoldenrod;
            chrt.Title.Font.LatinFont = "Arial";
            chrt.Title.Font.Bold = true;
            chrt.Title.Fill.Color = Color.White;
            chrt.Title.Border.Fill.Style = eFillStyle.SolidFill;
            chrt.Title.Border.LineStyle = eLineStyle.LongDashDotDot;
            chrt.Title.Border.Fill.Color = Color.Tomato;
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
            var rt = (ws.Drawings["shape1"] as ExcelShape).RichText.Add("Added formated richtext");
            rt.Bold = true;
            rt.Color = Color.Aquamarine;
            rt.Italic = true;
            rt.Size = 17;
            (ws.Drawings["shape2"] as ExcelShape).TextVertical = eTextVerticalType.Vertical;
            rt = (ws.Drawings["shape2"] as ExcelShape).RichText.Add("\r\nAdded formated richtext");
            rt.Bold = true;
            rt.Color = Color.DarkGoldenrod ;
            rt.SetFromFont(new Font("Times new roman", 18, FontStyle.Underline));
            rt.UnderLineColor = Color.Green;


            (ws.Drawings["shape3"] as ExcelShape).TextAnchoring=eTextAnchoringType.Bottom;
            (ws.Drawings["shape3"] as ExcelShape).TextAnchoringControl=true ;

            (ws.Drawings["shape4"] as ExcelShape).TextVertical = eTextVerticalType.Vertical270;
            (ws.Drawings["shape4"] as ExcelShape).TextAnchoring = eTextAnchoringType.Top;

            (ws.Drawings["shape5"] as ExcelShape).Fill.Style=eFillStyle.SolidFill;
            (ws.Drawings["shape5"] as ExcelShape).Fill.Color=Color.Red;
            (ws.Drawings["shape5"] as ExcelShape).Fill.Transparancy = 50;

            (ws.Drawings["shape6"] as ExcelShape).Fill.Style = eFillStyle.NoFill;
            (ws.Drawings["shape6"] as ExcelShape).Font.Color = Color.Black;
            (ws.Drawings["shape6"] as ExcelShape).Border.Fill.Color = Color.Black;

            (ws.Drawings["shape7"] as ExcelShape).Fill.Style = eFillStyle.SolidFill;
            (ws.Drawings["shape7"] as ExcelShape).Fill.Color=Color.Gray;
            (ws.Drawings["shape7"] as ExcelShape).Border.Fill.Style=eFillStyle.SolidFill;
            (ws.Drawings["shape7"] as ExcelShape).Border.Fill.Color = Color.Black;
            (ws.Drawings["shape7"] as ExcelShape).Border.Fill.Transparancy=43;
            (ws.Drawings["shape7"] as ExcelShape).Border.LineCap=eLineCap.Round;
            (ws.Drawings["shape7"] as ExcelShape).Border.LineStyle = eLineStyle.LongDash;
            (ws.Drawings["shape7"] as ExcelShape).Font.UnderLineColor = Color.Blue;
            (ws.Drawings["shape7"] as ExcelShape).Font.Color = Color.Black;
            (ws.Drawings["shape7"] as ExcelShape).Font.Bold = true;
            (ws.Drawings["shape7"] as ExcelShape).Font.LatinFont = "Arial";
            (ws.Drawings["shape7"] as ExcelShape).Font.ComplexFont = "Arial";
            (ws.Drawings["shape7"] as ExcelShape).Font.Italic = true;
            (ws.Drawings["shape7"] as ExcelShape).Font.UnderLine = eUnderLineType.Dotted;

            (ws.Drawings["shape8"] as ExcelShape).Fill.Style = eFillStyle.SolidFill;
            (ws.Drawings["shape8"] as ExcelShape).Font.LatinFont = "Miriam";
            (ws.Drawings["shape8"] as ExcelShape).Font.UnderLineColor = Color.CadetBlue;
            (ws.Drawings["shape8"] as ExcelShape).Font.UnderLine = eUnderLineType.Single;
        }
        [TestMethod]
        public void SaveDrawing()
        {
            _pck.Save();
        }
    }
}
