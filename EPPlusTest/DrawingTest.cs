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
using System.Xml;
namespace EPPlusTest
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

             ws.Column(1).Width = 53;
             ws.Column(4).Width = 58;

             pic = ws.Drawings.AddPicture("Pic6öäå", new FileInfo(@"C:\Program Files (x86)\Microsoft Office\CLIPART\PUB60COR\AG00004_.GIF"));
             pic.SetPosition(400, 400);
             pic.SetSize(100);
         }
         [TestMethod]
         public void DrawingSizingAndPositioning()
         {
             var ws = _pck.Workbook.Worksheets.Add("DrawingPosSize");

             var pic = ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);
             pic.SetPosition(1, 0, 1, 0);

             pic = ws.Drawings.AddPicture("Pic2", Properties.Resources.Test1);
             pic.EditAs = eEditAs.Absolute;
             pic.SetPosition(10, 5, 1, 4);

             pic = ws.Drawings.AddPicture("Pic3", Properties.Resources.Test1);
             pic.EditAs = eEditAs.TwoCell;
             pic.SetPosition(20, 5, 2, 4);

             
             ws.Column(1).Width = 100;
             ws.Column(3).Width = 100;
         }

        [TestMethod]
        public void BarChart()
        {
            var ws = _pck.Workbook.Worksheets.Add("BarChart");            
            var chrt = ws.Drawings.AddChart("barChart", eChartType.BarClustered) as ExcelBarChart;
            chrt.SetPosition(50, 50);
            chrt.SetSize(800, 300);
            AddTestSerie(ws, chrt);
            chrt.VaryColors = true;
            chrt.XAxis.Orientation = eAxisOrientation.MaxMin;
            chrt.XAxis.MajorTickMark = eAxisTickMark.In;
            chrt.YAxis.Orientation = eAxisOrientation.MaxMin;
            chrt.YAxis.MinorTickMark = eAxisTickMark.Out;
            chrt.ShowHiddenData = true;
            chrt.DisplayBlanksAs = eDisplayBlanksAs.Zero;
            chrt.Title.RichText.Text = "Barchart Test";
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

            ws.Cells["X19"].Value = "öäå";
            ws.Cells["X20"].Value = "ÖÄÅ";
            ws.Cells["X21"].Value = "üÛ";
            ws.Cells["X22"].Value = "&%#¤";
            ws.Cells["X23"].Value = "ÿ";
            ws.Cells["X24"].Value = "û";

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
            chrt.Legend.Position = eLegendPosition.TopRight;
            Assert.IsTrue(chrt.ChartType == eChartType.Pie, "Invalid Charttype");
            Assert.IsTrue(chrt.VaryColors);
            chrt.Title.Text = "Piechart";
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
            chrt.Legend.Position = eLegendPosition.Left;
            chrt.ShowHiddenData = false;
            chrt.DisplayBlanksAs = eDisplayBlanksAs.Gap;
            chrt.Title.RichText.Add("Pie RT Title add");
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
            //chrt.Title.Text = "Header Text";
            var r1=chrt.Title.RichText.Add("Header");
            r1.Bold = true;
            var r2=chrt.Title.RichText.Add("  Text");
            r2.UnderLine = eUnderLineType.WavyHeavy;

            chrt.Title.Fill.Style = eFillStyle.SolidFill;
            chrt.Title.Fill.Color = Color.LightBlue;
            chrt.Title.Fill.Transparancy = 50;
            chrt.VaryColors = true;
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
            chrt.Legend.Position = eLegendPosition.Bottom;
            
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
            chrt.VaryColors = true;
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
            chrt.Legend.Position = eLegendPosition.Top;
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
            chrt.VaryColors = true;
            chrt.SetSize(200);
            chrt.Title.Text = "Cone bar";
            chrt.Series[0].Header = "Serie 1";
            chrt.Legend.Position = eLegendPosition.Right;
        }
        [TestMethod]
        public void Column()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column");
            var chrt = ws.Drawings.AddChart("Column1", eChartType.ColumnClustered3D) as ExcelBarChart;
            AddTestSerie(ws, chrt);
            chrt.VaryColors = true;
            chrt.View3D.RightAngleAxes = true;
            chrt.View3D.DepthPercent = 99;
            //chrt.View3D.HeightPercent = 99;
            chrt.View3D.RightAngleAxes = true;
            chrt.SetSize(200);
            chrt.Title.Text = "Column";
            chrt.Series[0].Header = "Serie 1";
            chrt.Locked = false;
            chrt.Print = false;
            chrt.EditAs = eEditAs.TwoCell;
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
            chrt.EditAs = eEditAs.Absolute;
        }
        [TestMethod]
        public void Line()
        {
            var ws = _pck.Workbook.Worksheets.Add("Line");
            var chrt = ws.Drawings.AddChart("Line1", eChartType.Line) as ExcelLineChart;
            AddTestSerie(ws, chrt);
            chrt.SetSize(150);
            chrt.VaryColors = true;
            chrt.Smooth = false;
            chrt.Title.Text = "Line 3D";
            chrt.Series[0].Header = "Line serie 1";
            var tl = chrt.Series[0].TrendLines.Add(eTrendLine.Polynomial);
            tl.Name = "Test";
            tl.DisplayRSquaredValue = true;
            tl.DisplayEquation = true;
            tl.Forward = 15;
            tl.Backward = 1;
            tl.Intercept = 6;
            //tl.Period = 12;
            tl.Order = 5;

            tl = chrt.Series[0].TrendLines.Add(eTrendLine.MovingAvgerage);
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
            chrt.DataLabel.ShowSeriesName = true;
            chrt.DataLabel.ShowLeaderLines=true;
            chrt.EditAs = eEditAs.OneCell;
            chrt.DisplayBlanksAs = eDisplayBlanksAs.Span;
            chrt.Axis[0].Title.Text = "Axis 0";
            chrt.Axis[0].Title.Rotation = 90;
            chrt.Axis[0].Title.Overlay = true;
            chrt.Axis[1].Title.Text = "Axis 1";
            chrt.Axis[1].Title.AnchorCtr = true;
            chrt.Axis[1].Title.TextVertical = eTextVerticalType.Vertical270;
            chrt.Axis[1].Title.Border.LineStyle=eLineStyle.LongDashDotDot;

        }
        [TestMethod]
        public void LineMarker()
        {
            var ws = _pck.Workbook.Worksheets.Add("LineMarker1");
            var chrt = ws.Drawings.AddChart("Line1", eChartType.LineMarkers) as ExcelLineChart;
            AddTestSerie(ws, chrt);
            chrt.SetSize(150);
            chrt.Title.Text = "Line Markers";
            chrt.Series[0].Header = "Line serie 1";
            ((ExcelLineChartSerie)chrt.Series[0]).Marker = eMarkerStyle.Plus;

            var chrt2 = ws.Drawings.AddChart("Line2", eChartType.LineMarkers) as ExcelLineChart;
            AddTestSerie(ws, chrt2);
            chrt2.SetPosition(500,0);
            chrt2.SetSize(150);
            chrt2.Title.Text = "Line Markers";
            var serie = (ExcelLineChartSerie)chrt2.Series[0];
            serie.Marker = eMarkerStyle.X;

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
            (ws.Drawings["shape1"] as ExcelShape).LockText = false;
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

            (ws.Drawings["shape9"] as ExcelShape).TextAlignment = eTextAlignment.Right;

        }
        [TestMethod]
        public void DrawingWorksheetCopy()
        {
            var wsShapes = _pck.Workbook.Worksheets.Add("Copy Shapes", _pck.Workbook.Worksheets["Shapes"]);
            var wsScatterChart = _pck.Workbook.Worksheets.Add("Copy Scatter", _pck.Workbook.Worksheets["Scatter"]);
            var wsPicture = _pck.Workbook.Worksheets.Add("Copy Picture", _pck.Workbook.Worksheets["Picture"]);
        }    
        [TestMethod]
        public void Line2Test()
        {
           ExcelWorksheet worksheet = _pck.Workbook.Worksheets.Add("LineIssue");

           ExcelChart chart = worksheet.Drawings.AddChart("LineChart", eChartType.Line);
           
           worksheet.Cells["A1"].Value=1;
           worksheet.Cells["A2"].Value=2;
           worksheet.Cells["A3"].Value=3;
           worksheet.Cells["A4"].Value=4;
           worksheet.Cells["A5"].Value=5;
           worksheet.Cells["A6"].Value=6;

           worksheet.Cells["B1"].Value=10000;
           worksheet.Cells["B2"].Value=10100;
           worksheet.Cells["B3"].Value=10200;
           worksheet.Cells["B4"].Value=10150;
           worksheet.Cells["B5"].Value=10250;
           worksheet.Cells["B6"].Value=10200;

           chart.Series.Add(ExcelRange.GetAddress(1, 2, worksheet.Dimension.End.Row, 2),
                            ExcelRange.GetAddress(1, 1, worksheet.Dimension.End.Row, 1));

           var Series = chart.Series[0];
           
           chart.Series[0].Header = "Blah";
        }
        [TestMethod]
        public void MultiChartSeries()
        {
            ExcelWorksheet worksheet = _pck.Workbook.Worksheets.Add("MultiChartTypes");

            ExcelChart chart = worksheet.Drawings.AddChart("chtPie", eChartType.LineMarkers);
            chart.SetPosition(100, 100);
            chart.SetSize(800,600);
            AddTestSerie(worksheet, chart);
            chart.Series[0].Header = "Serie5";
            chart.Style = eChartStyle.Style27;
            worksheet.Cells["W19"].Value = 120;
            worksheet.Cells["W20"].Value = 122;
            worksheet.Cells["W21"].Value = 121;
            worksheet.Cells["W22"].Value = 123;
            worksheet.Cells["W23"].Value = 125;
            worksheet.Cells["W24"].Value = 124;

            worksheet.Cells["X19"].Value = 90;
            worksheet.Cells["X20"].Value = 52;
            worksheet.Cells["X21"].Value = 88;
            worksheet.Cells["X22"].Value = 75;
            worksheet.Cells["X23"].Value = 77;
            worksheet.Cells["X24"].Value = 99;
            
            var cs2 = chart.PlotArea.ChartTypes.Add(eChartType.ColumnClustered);
            var s = cs2.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["U19:U24"]);
            s.Header = "Serie4";
            cs2.YAxis.MaxValue = 300;
            cs2.YAxis.MinValue = -5.5;
            var cs3 = chart.PlotArea.ChartTypes.Add(eChartType.Line);
            s=cs3.Series.Add(worksheet.Cells["X19:X24"], worksheet.Cells["U19:U24"]);
            s.Header = "Serie1";
            cs3.UseSecondaryAxis = true;
                        
            cs3.XAxis.Deleted = false;
            cs3.XAxis.MajorUnit = 20;
            cs3.XAxis.MinorUnit = 3;

            cs3.XAxis.TickLabelPosition = eTickLabelPosition.High;
            cs3.YAxis.LogBase = 10.2;

            var chart2 = worksheet.Drawings.AddChart("scatter1", eChartType.XYScatterSmooth);
            s=chart2.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["U19:U24"]);
            s.Header = "Serie2";

            var c2ct2 = chart2.PlotArea.ChartTypes.Add(eChartType.XYScatterSmooth);
            s=c2ct2.Series.Add(worksheet.Cells["X19:X24"], worksheet.Cells["V19:V24"]);
            s.Header="Serie3";
            s=c2ct2.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["V19:V24"]);
            s.Header = "Serie4";

            c2ct2.UseSecondaryAxis = true;
            c2ct2.XAxis.Deleted = false;
            c2ct2.XAxis.TickLabelPosition = eTickLabelPosition.High;

            ExcelChart chart3 = worksheet.Drawings.AddChart("chart", eChartType.LineMarkers);
            chart3.SetPosition(300, 1000);
            var s31=chart3.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["U19:U24"]);
            s31.Header = "Serie1";

            var c3ct2 = chart3.PlotArea.ChartTypes.Add(eChartType.LineMarkers);
            var c32 = c3ct2.Series.Add(worksheet.Cells["X19:X24"], worksheet.Cells["V19:V24"]);
            c3ct2.UseSecondaryAxis = true;
            c32.Header = "Serie2";
            
            XmlNamespaceManager ns=new XmlNamespaceManager(new NameTable());
            ns.AddNamespace("c","http://schemas.openxmlformats.org/drawingml/2006/chart");
            var element = chart.ChartXml.SelectSingleNode("//c:plotVisOnly", ns);
            if (element!=null) element.ParentNode.RemoveChild(element);
        }
        [TestMethod]
        public void DeleteDrawing()
        {
            var ws=_pck.Workbook.Worksheets.Add("DeleteDrawing1");
            var chart1 = ws.Drawings.AddChart("Chart1", eChartType.Line);
            var chart2 = ws.Drawings.AddChart("Chart2", eChartType.Line);
            var shape1 = ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonBackPrevious);
            var pic1 = ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);
            ws.Drawings.Remove(2);
            ws.Drawings.Remove(chart2);
            ws.Drawings.Remove("Pic1");

            ws = _pck.Workbook.Worksheets.Add("DeleteDrawing2");
            chart1 = ws.Drawings.AddChart("Chart1", eChartType.Line);
            chart2 = ws.Drawings.AddChart("Chart2", eChartType.Line);
            shape1 = ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonBackPrevious);
            pic1 = ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);

            ws.Drawings.Remove("chart1");

            ws = _pck.Workbook.Worksheets.Add("ClearDrawing2");
            chart1 = ws.Drawings.AddChart("Chart1", eChartType.Line);
            chart2 = ws.Drawings.AddChart("Chart2", eChartType.Line);
            shape1 = ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonBackPrevious);
            pic1 = ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);
            ws.Drawings.Clear();
        }
        [TestMethod]
        public void ReadDocument()
        {
            ExcelPackage pck = new ExcelPackage(new FileInfo("Test\\Drawing.xlsx"),true);

            foreach(var ws in pck.Workbook.Worksheets)
            {
                foreach(ExcelDrawing d in pck.Workbook.Worksheets[1].Drawings)
                {
                    if (d is ExcelChart)
                    {
                        TestContext.WriteLine(((ExcelChart)d).ChartType.ToString());
                    }
                }
            }
        }
        [TestMethod]
        public void SaveDrawing()
        {
            _pck.Save();
        }   
        [TestMethod]
        public void ReadMultiChartSeries()
        {
            ExcelPackage pck = new ExcelPackage(new FileInfo("c:\\temp\\chartseries.xlsx"), true);

            var ws = pck.Workbook.Worksheets[1];
            ExcelChart c = ws.Drawings[0] as ExcelChart;

            var p = c.PlotArea;
            p.ChartTypes[1].Series[0].Series = "S7:S15";

            var c2=ws.Drawings.AddChart("NewChart", eChartType.ColumnClustered);
            var serie1 = c2.Series.Add("R7:R15", "Q7:Q15");
            c2.SetSize(800, 800);
            serie1.Header = "Column Clustered";

            var subChart = c2.PlotArea.ChartTypes.Add(eChartType.LineMarkers);
            var serie2 = subChart.Series.Add("S7:S15", "Q7:Q15");
            serie2.Header = "Line";

            //var subChart2 = c2.PlotArea.ChartTypes.Add(eChartType.DoughnutExploded);
            //var serie3 = subChart2.Series.Add("S7:S15", "Q7:Q15");
            //serie3.Header = "Doughnut";

            var subChart3 = c2.PlotArea.ChartTypes.Add(eChartType.Area);
            var serie4 = subChart3.Series.Add("R7:R15", "Q7:Q15");
            serie4.Header = "Area";
            subChart3.UseSecondaryAxis = true;

            var serie5 = subChart.Series.Add("R7:R15","Q7:Q15");
            serie5.Header = "Line 2";

            pck.SaveAs(new FileInfo("c:\\temp\\chartseriesnew.xlsx"));
        }
    }
}
