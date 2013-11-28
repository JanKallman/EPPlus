using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System.Globalization;
using OfficeOpenXml;
using System.IO;
using System.Diagnostics;

namespace ExcelPackageTest
{
    [TestClass]
    public class StyleTest
    {
        [TestMethod]
        public void GetStyles()
        {
            ExcelPackage pck = new ExcelPackage(new FileInfo("c:\\temp\\formats.xlsx"));

            foreach (var cell in pck.Workbook.Worksheets[1].Cells["A:X"])
            {
                Debug.WriteLine(cell.Text);
            }
        }
        [TestMethod]
        public void TestFraction()
        {
            //string f = ExcelNumberFormatXml.FormatFraction(0.25, 2);
            
            //f = ExcelNumberFormatXml.FormatFraction(0.333, 2);
            //f = ExcelNumberFormatXml.FormatFraction(-0.888, 2);
            //f = ExcelNumberFormatXml.FormatFraction(-0.272471, 3);
            //f = ExcelNumberFormatXml.FormatFraction(0.666666, 4);
            //f = ExcelNumberFormatXml.FormatFraction(0.21, 4);
            //f = ExcelNumberFormatXml.FormatFraction(0.09999, 3);
            //f = ExcelNumberFormatXml.FormatFraction(0, 3);

        }
        
    }
}
