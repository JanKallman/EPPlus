using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusTest
{
    [TestClass]
    public class DTS_FailingTests
    {

        [TestMethod]
        public void DeleteWorksheetWithReferencedImage()
        {
            var ms = new MemoryStream();
            using (var pck = new ExcelPackage(ms))
            {
                var ws = pck.Workbook.Worksheets.Add("original");
                ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);
                pck.Workbook.Worksheets.Copy("original", "copy");
                pck.SaveAs(ms);
            }
            ms.Position = 0;

            using (var pck = new ExcelPackage(ms))
            {
                var ws = pck.Workbook.Worksheets["original"];
                pck.Workbook.Worksheets.Delete(ws);
                pck.Save();
            }
        }

        [TestMethod]
        public void CopyAndDeleteWorksheetWithImage()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var ws = pck.Workbook.Worksheets.Add("original");
                ws.Drawings.AddPicture("Pic1", Properties.Resources.Test1);
                pck.Workbook.Worksheets.Copy("original", "copy");
                pck.Workbook.Worksheets.Delete(ws);
                pck.Save();
            }
        }
    }
}
