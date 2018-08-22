using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace EPPlusTest
{
    [TestClass]
    public class ExcelPackageTests
    {
        [TestMethod]
        public void SaveAsyncTest()
        {
            using (var ms = new MemoryStream())
            {
                using (var package = new ExcelPackage(ms))
                {
                    var sheet = package.Workbook.Worksheets.Add("test");
                    var task = package.SaveAsync();
                    Assert.AreEqual(0, ms.Length);
                    task.Wait();
                    Assert.IsTrue(ms.Length > 0);

                }
            }
        }
    }
}
