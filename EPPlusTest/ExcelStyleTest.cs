using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

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
                ws.Cells["B2"].Style.QuotePrefix = true;
                Assert.IsTrue(ws.Cells["B2"].Style.QuotePrefix);

                p.Workbook.Styles.UpdateXml();
                var node = p.Workbook.StylesXml.SelectSingleNode("//d:cellXfs/d:xf", p.Workbook.NameSpaceManager);
                Assert.AreEqual("1", node.Attributes["quotePrefix"].Value);
            }
        }
    }
}
