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
            var pck = new ExcelPackage(new FileInfo("c:\\temp\\chaintest.xlsx"));
            pck.Workbook.Calculate();
        }
    }
}
