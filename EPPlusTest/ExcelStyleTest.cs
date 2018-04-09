﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Xml;
using OfficeOpenXml.Style.XmlAccess;

namespace EPPlusTest
{
    [TestClass]
    public class ExcelStyleTest
    {
        [TestMethod]
        public void QuotePrefixStyle()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("QuotePrefixTest");
                var cell = ws.Cells["B2"];
                cell.Style.QuotePrefix = true;
                Assert.IsTrue(cell.Style.QuotePrefix);

                p.Workbook.Styles.UpdateXml();                
                var nodes = p.Workbook.StylesXml.SelectNodes("//d:cellXfs/d:xf", p.Workbook.NameSpaceManager);
                // Since the quotePrefix attribute is not part of the default style,
                // a new one should be created and referenced.
                Assert.AreNotEqual(0, cell.StyleID);
                Assert.IsNull(nodes[0].Attributes["quotePrefix"]);
                Assert.AreEqual("1", nodes[cell.StyleID].Attributes["quotePrefix"].Value);
            }
        }

        [TestMethod]
        public void GetFontHeightBug()
        {
            var r = ExcelFontXml.GetFontHeight("Calibri Light", 4);
            Assert.AreEqual(r, 20);
        }
    }
}
