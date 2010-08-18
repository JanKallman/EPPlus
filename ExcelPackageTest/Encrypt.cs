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
            using (ExcelPackage pck = new ExcelPackage(new FileInfo(@"c:\temp\SampleApp\sample7.xlsx"), true))
            //using (ExcelPackage pck = new ExcelPackage(new FileInfo(@"c:\temp\sample7Encr_Test.xlsx"), true, "EPPlus"))
    
            {
                pck.Encryption.Password = "EPPlus";
                pck.Encryption.Algorithm = EncryptionAlgorithm.AES192;
                pck.Workbook.Protection.SetPassword("test");
                pck.Workbook.Protection.LockStructure = true;
                //pck.Workbook.Protection.LockWindows = true;

                pck.SaveAs(new FileInfo(@"c:\temp\sample7Encr_Test.xlsx"));
            }

            using (ExcelPackage pck = new ExcelPackage(new FileInfo(@"c:\temp\sample7Encr_Test.xlsx"), true, "EPPlus"))            
            {
                pck.Encryption.IsEncrypted = false;
                pck.SaveAs(new FileInfo(@"c:\temp\sample7NotEncr.xlsx"));
            }
            

        }
        [TestMethod]
        public void WriteEncrypt()
        {
            ExcelPackage package = new ExcelPackage();
            //Load the sheet with one string column, one date column and a few random numbers.
            var ws = package.Workbook.Worksheets.Add("First line test");

            ws.Cells[1, 1].Value = "1; 1";
            ws.Cells[2, 1].Value = "2; 1";
            ws.Cells[1, 2].Value = "1; 2";
            ws.Cells[2, 2].Value = "2; 2";

            ws.Row(1).Style.Font.Bold = true;
            ws.Column(1).Style.Font.Bold = true;

            package.Encryption.Algorithm = EncryptionAlgorithm.AES256;
            package.SaveAs(new FileInfo(@"c:\temp\encrTest.xlsx"), "test");
        }
    }
}
