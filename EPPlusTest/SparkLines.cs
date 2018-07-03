using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Sparkline;

namespace EPPlusTest
{
    [TestClass]
    public class SparkLines : TestBase
    {
        string _pckfile;
        public SparkLines()
        {
            InitBase();
            _pckfile = "Sparklines.xlsx";
        }
        [TestMethod]
        public void StartTest()
        {
            WriteSparklines();
            ReadSparklines();
        }
        public void ReadSparklines()
        {
            _pck = new ExcelPackage();
            OpenPackage(_pckfile);
            var ws = _pck.Workbook.Worksheets[_pck.Compatibility.IsWorksheets1Based?1:0];
            Assert.AreEqual(4, ws.SparklineGroups.Count);
            var sg1 = ws.SparklineGroups[0];
            Assert.AreEqual("A1:A4",sg1.LocationRange.Address);
            Assert.AreEqual("B1:C4", sg1.DataRange.Address);
            Assert.AreEqual(null, sg1.DateAxisRange);

            var sg2 = ws.SparklineGroups[1];
            Assert.AreEqual("D1:D2", sg2.LocationRange.Address);
            Assert.AreEqual("B1:C4", sg2.DataRange.Address);

            var sg3 = ws.SparklineGroups[2];
            Assert.AreEqual("A10:B10", sg3.LocationRange.Address);
            Assert.AreEqual("B1:C4", sg3.DataRange.Address);

            var sg4 = ws.SparklineGroups[3];
            Assert.AreEqual("D10:G10", sg4.LocationRange.Address);
            Assert.AreEqual("B1:C4", sg4.DataRange.Address);
            Assert.AreEqual("'Sparklines'!A20:A23", sg4.DateAxisRange.Address);

            var c1 = sg1.ColorMarkers;
            Assert.AreEqual(c1.Rgb, "FFD00000");
            var ec = sg1.DisplayEmptyCellsAs;
            Assert.AreEqual(eDispBlanksAs.Gap, ec);
            var t = sg1.Type;
        }
        public void WriteSparklines()
        {            
            var ws = _pck.Workbook.Worksheets.Add("Sparklines");
            ws.Cells["B1"].Value = 15;
            ws.Cells["B2"].Value = 30;
            ws.Cells["B3"].Value = 35;
            ws.Cells["B4"].Value = 28;
            ws.Cells["C1"].Value = 7;
            ws.Cells["C2"].Value = 33;
            ws.Cells["C3"].Value = 12;
            ws.Cells["C4"].Value = -1;

            //Column<->Row
            var sg1 = ws.SparklineGroups.Add(eSparklineType.Line, ws.Cells["A1:A4"], ws.Cells["B1:C4"]);
            sg1.DisplayEmptyCellsAs = eDispBlanksAs.Gap;
            sg1.Type = eSparklineType.Line;

            //Column<->Column
            var sg2 = ws.SparklineGroups.Add(eSparklineType.Column, ws.Cells["D1:D2"], ws.Cells["B1:C4"]);

            //Row<->Column
            var sg3 = ws.SparklineGroups.Add(eSparklineType.Stacked, ws.Cells["A10:B10"], ws.Cells["B1:C4"]);
            sg3.RightToLeft=true;
            //Row<->Row
            var sg4 = ws.SparklineGroups.Add(eSparklineType.Line, ws.Cells["D10:G10"], ws.Cells["B1:C4"]);
            ws.Cells["A20"].Value = new DateTime(2016, 12, 30);
            ws.Cells["A21"].Value = new DateTime(2017, 1, 31);
            ws.Cells["A22"].Value = new DateTime(2017, 2, 28);
            ws.Cells["A23"].Value = new DateTime(2017, 3, 31);

            sg4.DateAxisRange = ws.Cells["A20:A23"];

            sg4.ManualMax = 5;
            sg4.ManualMin = 3;

            SaveWorksheet(_pckfile);
        }
    }
}
