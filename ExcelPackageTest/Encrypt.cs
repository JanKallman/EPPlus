using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace ExcelPackageTest
{
    [TestClass]
    public class Encrypt
    {
        [TestMethod]
        public void ReadEncrypt()
        {
            MemoryStream ms=new MemoryStream();
            using (ExcelPackage pck = new ExcelPackage(new FileInfo(@"c:\temp\SampleApp\sample7.xlsx"), true, "test"))
            {
                
            }

        }        
    }
}
