using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest
{
    [TestClass]
    public class ExcelRangeBase
    {
        [TestMethod]
        public void ClearMethodShouldNotClearSurroundingCells()
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                var wks = excel.Workbook.Worksheets.Add("test");
                wks.Cells[2, 2].Value = "something";
                wks.Cells[2, 3].Value = "something";
                

                wks.Cells[2, 3].Clear();

                Assert.IsNotNull(wks.Cells[2, 2].Value);
            }
        }
    }
}
