using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Utils;
using OfficeOpenXml;
using System.Security.Cryptography.X509Certificates;
using System.IO;
using System.Security.Cryptography;
using OfficeOpenXml.VBA;

namespace EPPlusTest
{
    [TestClass]
    public class VBA
    {
        [Ignore]
        [TestMethod]
        public void Compression()
        {
            //Compression/Decompression
            string value = "#aaabcdefaaaaghijaaaaaklaaamnopqaaaaaaaaaaaarstuvwxyzaaa";

            byte[] compValue = CompoundDocument.CompressPart(Encoding.GetEncoding(1252).GetBytes(value));
            string decompValue = Encoding.GetEncoding(1252).GetString(CompoundDocument.DecompressPart(compValue));
            Assert.AreEqual(value, decompValue);

            value = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";

            compValue = CompoundDocument.CompressPart(Encoding.GetEncoding(1252).GetBytes(value));
            decompValue = Encoding.GetEncoding(1252).GetString(CompoundDocument.DecompressPart(compValue));
            Assert.AreEqual(value, decompValue);
        }
        [Ignore]
        [TestMethod]
        public void ReadVBA()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\report.xlsm"));
            File.WriteAllText(@"c:\temp\vba\modules\dir.txt", package.Workbook.VbaProject.CodePage + "," + package.Workbook.VbaProject.Constants + "," + package.Workbook.VbaProject.Description + "," + package.Workbook.VbaProject.HelpContextID.ToString() + "," + package.Workbook.VbaProject.HelpFile1 + "," + package.Workbook.VbaProject.HelpFile2 + "," + package.Workbook.VbaProject.Lcid.ToString() + "," + package.Workbook.VbaProject.LcidInvoke.ToString() + "," + package.Workbook.VbaProject.LibFlags.ToString() + "," + package.Workbook.VbaProject.MajorVersion.ToString() + "," + package.Workbook.VbaProject.MinorVersion.ToString() + "," + package.Workbook.VbaProject.Name + "," + package.Workbook.VbaProject.ProjectID + "," + package.Workbook.VbaProject.SystemKind.ToString() + "," + package.Workbook.VbaProject.Protection.HostProtected.ToString() + "," + package.Workbook.VbaProject.Protection.UserProtected.ToString() + "," + package.Workbook.VbaProject.Protection.VbeProtected.ToString() + "," + package.Workbook.VbaProject.Protection.VisibilityState.ToString());
            foreach (var module in package.Workbook.VbaProject.Modules)
            {
                File.WriteAllText(string.Format(@"c:\temp\vba\modules\{0}.txt", module.Name), module.Code);
            }
            foreach (var r in package.Workbook.VbaProject.References)
            {
                File.WriteAllText(string.Format(@"c:\temp\vba\modules\{0}.txt", r.Name), r.Libid + " " + r.ReferenceRecordID.ToString());
            }

            List<X509Certificate2> ret = new List<X509Certificate2>();
            X509Store store = new X509Store(StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            package.Workbook.VbaProject.Signature.Certificate = store.Certificates[19];
            //package.Workbook.VbaProject.Protection.SetPassword("");
            package.SaveAs(new FileInfo(@"c:\temp\vbaSaved.xlsm"));
        }
        [Ignore]
        [TestMethod]
        public void WriteVBA()
        {
            var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("Sheet1");
            package.Workbook.CreateVBAProject();
            package.Workbook.VbaProject.Modules["Sheet1"].Code += "\r\nPrivate Sub Worksheet_SelectionChange(ByVal Target As Range)\r\nMsgBox(\"Test of the VBA Feature!\")\r\nEnd Sub\r\n";
            package.Workbook.VbaProject.Modules["Sheet1"].Name = "Blad1";
            package.Workbook.CodeModule.Name = "DenHärArbetsboken";
            package.Workbook.Worksheets[1].Name = "FirstSheet";
            package.Workbook.CodeModule.Code += "\r\nPrivate Sub Workbook_Open()\r\nBlad1.Cells(1,1).Value = \"VBA test\"\r\nMsgBox \"VBA is running!\"\r\nEnd Sub";
            //X509Store store = new X509Store(StoreLocation.CurrentUser);
            //store.Open(OpenFlags.ReadOnly);
            //package.Workbook.VbaProject.Signature.Certificate = store.Certificates[11];

            var m = package.Workbook.VbaProject.Modules.AddModule("Module1");
            m.Code += "Public Sub Test(param1 as string)\r\n\r\nEnd sub\r\nPublic Function functest() As String\r\n\r\nEnd Function\r\n";
            var c = package.Workbook.VbaProject.Modules.AddClass("Class1", false);
            c.Code += "Private Sub Class_Initialize()\r\n\r\nEnd Sub\r\nPrivate Sub Class_Terminate()\r\n\r\nEnd Sub";
            var c2 = package.Workbook.VbaProject.Modules.AddClass("Class2", true);
            c2.Code += "Private Sub Class_Initialize()\r\n\r\nEnd Sub\r\nPrivate Sub Class_Terminate()\r\n\r\nEnd Sub";

            package.Workbook.VbaProject.Protection.SetPassword("EPPlus");
            package.SaveAs(new FileInfo(@"c:\temp\vbaWrite.xlsm"));

        }
        [Ignore]
        [TestMethod]
        public void Resign()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\vbaWrite.xlsm"));
            //package.Workbook.VbaProject.Signature.Certificate = store.Certificates[11];
            package.SaveAs(new FileInfo(@"c:\temp\vbaWrite2.xlsm"));
        }
        [Ignore]
        [TestMethod]
        public void WriteLongVBAModule()
        {
            var package = new ExcelPackage();
            package.Workbook.Worksheets.Add("VBASetData");
            package.Workbook.CreateVBAProject();
            package.Workbook.CodeModule.Code = "Private Sub Workbook_Open()\r\nCreateData\r\nEnd Sub";
            var module = package.Workbook.VbaProject.Modules.AddModule("Code");

            StringBuilder code = new StringBuilder("Public Sub CreateData()\r\n");
            for (int row = 1; row < 30; row++)
            {
                for (int col = 1; col < 30; col++)
                {
                    code.AppendLine(string.Format("VBASetData.Cells({0},{1}).Value=\"Cell {2}\"", row, col, new ExcelAddressBase(row, col, row, col).Address));
                }
            }
            code.AppendLine("End Sub");
            module.Code = code.ToString();

            //X509Store store = new X509Store(StoreLocation.CurrentUser);
            //store.Open(OpenFlags.ReadOnly);
            //package.Workbook.VbaProject.Signature.Certificate = store.Certificates[19];

            package.SaveAs(new FileInfo(@"c:\temp\vbaLong.xlsm"));
        }
        [TestMethod]
        public void VbaError()
        {
            DirectoryInfo workingDir = new DirectoryInfo(@"C:\epplusExample\folder");
            if (!workingDir.Exists) workingDir.Create();
            FileInfo f = new FileInfo(workingDir.FullName + "//" + "temp.xlsx");
            if (f.Exists) f.Delete();
            ExcelPackage myPackage = new ExcelPackage(f);
            myPackage.Workbook.CreateVBAProject();
            ExcelWorksheet excelWorksheet = myPackage.Workbook.Worksheets.Add("Sheet1");
            ExcelWorksheet excelWorksheet2 = myPackage.Workbook.Worksheets.Add("Sheet2");
            ExcelWorksheet excelWorksheet3 = myPackage.Workbook.Worksheets.Add("Sheet3");
            FileInfo f2 = new FileInfo(workingDir.FullName + "//" + "newfile.xlsm");
            ExcelVBAModule excelVbaModule = myPackage.Workbook.VbaProject.Modules.AddModule("Module1");
            StringBuilder mybuilder = new StringBuilder(); mybuilder.AppendLine("Sub Jiminy()");
            mybuilder.AppendLine("Range(\"D6\").Select");
            mybuilder.AppendLine("ActiveCell.FormulaR1C1 = \"Jiminy\"");
            mybuilder.AppendLine("End Sub");
            excelVbaModule.Code = mybuilder.ToString();
            myPackage.SaveAs(f2);
            myPackage.Dispose();
        }
        [Ignore]
        [TestMethod]
        public void ReadVBAUnicodeWsName()
        {
            var package = new ExcelPackage(new FileInfo(@"c:\temp\bug\VbaUnicodeWS.xlsm"));
            File.WriteAllText(@"c:\temp\vba\modules\dir.txt", package.Workbook.VbaProject.CodePage + "," + package.Workbook.VbaProject.Constants + "," + package.Workbook.VbaProject.Description + "," + package.Workbook.VbaProject.HelpContextID.ToString() + "," + package.Workbook.VbaProject.HelpFile1 + "," + package.Workbook.VbaProject.HelpFile2 + "," + package.Workbook.VbaProject.Lcid.ToString() + "," + package.Workbook.VbaProject.LcidInvoke.ToString() + "," + package.Workbook.VbaProject.LibFlags.ToString() + "," + package.Workbook.VbaProject.MajorVersion.ToString() + "," + package.Workbook.VbaProject.MinorVersion.ToString() + "," + package.Workbook.VbaProject.Name + "," + package.Workbook.VbaProject.ProjectID + "," + package.Workbook.VbaProject.SystemKind.ToString() + "," + package.Workbook.VbaProject.Protection.HostProtected.ToString() + "," + package.Workbook.VbaProject.Protection.UserProtected.ToString() + "," + package.Workbook.VbaProject.Protection.VbeProtected.ToString() + "," + package.Workbook.VbaProject.Protection.VisibilityState.ToString());
            foreach (var module in package.Workbook.VbaProject.Modules)
            {
                File.WriteAllText(string.Format(@"c:\temp\vba\modules\{0}.txt", module.Name), module.Code);
            }
            foreach (var r in package.Workbook.VbaProject.References)
            {
                File.WriteAllText(string.Format(@"c:\temp\vba\modules\{0}.txt", r.Name), r.Libid + " " + r.ReferenceRecordID.ToString());
            }

            List<X509Certificate2> ret = new List<X509Certificate2>();
            X509Store store = new X509Store(StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            package.Workbook.VbaProject.Signature.Certificate = store.Certificates[19];
            //package.Workbook.VbaProject.Protection.SetPassword("");
            package.SaveAs(new FileInfo(@"c:\temp\vbaSaved.xlsm"));
        }
        [TestMethod]
        public void CreateUnicodeWsName()
        {
            using (var package = new ExcelPackage())
            {
                //ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Test");
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("测试");

                package.Workbook.CreateVBAProject();
                var sb = new StringBuilder();
                sb.AppendLine("Sub GetData()");
                sb.AppendLine("MsgBox (\"Hello,World\")");
                sb.AppendLine("End Sub");
                
                ExcelWorksheet worksheet2 = package.Workbook.Worksheets.Add("Sheet1");
                var stringBuilder = new StringBuilder();
                stringBuilder.AppendLine("Private Sub Worksheet_Change(ByVal Target As Range)");
                stringBuilder.AppendLine("GetData");
                stringBuilder.AppendLine("End Sub");
                worksheet.CodeModule.Code = stringBuilder.ToString();

                package.SaveAs(new FileInfo(@"c:\temp\invvba.xlsm"));
            }
        }
    }
}