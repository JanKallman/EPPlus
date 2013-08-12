using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Calculation;
using System.IO;

namespace EPPlusTest
{
    [TestClass]
    public class Calculation
    {
        [TestMethod]
        public void Calulation()
        {
            var pck = new ExcelPackage(new FileInfo("c:\\temp\\chain.xlsx"));
            pck.Workbook.Calculate();
            Assert.AreEqual(50D, pck.Workbook.Worksheets[1].Cells["C1"].Value);
        }
        [TestMethod]
        public void Calulation2()
        {
            var pck = new ExcelPackage(new FileInfo("c:\\temp\\chainTest.xlsx"));
            pck.Workbook.Calculate();
            Assert.AreEqual(1124999960382D, pck.Workbook.Worksheets[1].Cells["C1"].Value);
        }
    }
}
