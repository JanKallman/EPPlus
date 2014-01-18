using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace EPPlusTest
{
    [TestClass]
    public class TestBase
    {
        protected ExcelPackage _pck;        
        protected string _worksheetPath="";
        protected string _worksheetName = "";
        public TestContext TestContext { get; set; }
        protected void InitBase()
        {

            _worksheetPath = AppDomain.CurrentDomain.BaseDirectory + @"\..\..\worksheets";
            if (!Directory.Exists(_worksheetPath))
            {
                Directory.CreateDirectory(_worksheetPath);
            }
            var di=new DirectoryInfo(_worksheetPath);            
            _worksheetPath = di.FullName + "\\";

            _pck = new ExcelPackage();
        }        
        protected void OpenPackage(string name)
        {
            var fi = new FileInfo(_worksheetPath + name);
            _pck = new ExcelPackage(fi);
        }
        protected void SaveWorksheet(string name)
        {
            if (_pck.Workbook.Worksheets.Count == 0) return;
            var fi = new FileInfo(_worksheetPath + name);
            if (fi.Exists)
            {
                fi.Delete();
            }
            _pck.SaveAs(fi);
        }
    }
}
