using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using System.Reflection;

namespace EPPlusTest
{
    [TestClass]
    public abstract class TestBase
    {
        protected ExcelPackage _pck;
        protected string _clipartPath="";
        protected string _worksheetPath="";
        public TestContext TestContext { get; set; }
        
        [TestInitialize]
        public void InitBase()
        {

            _clipartPath = Path.Combine(Path.GetTempPath(), @"EPPlus clipart");
            if (!Directory.Exists(_clipartPath))
            {
                Directory.CreateDirectory(_clipartPath);
            }
            var asm = Assembly.GetExecutingAssembly();
            var validExtensions = new[]
                {
                    ".gif", ".wmf"
                };
            foreach (var name in asm.GetManifestResourceNames())
            {
                foreach (var ext in validExtensions)
                {
                    if (name.EndsWith(ext, StringComparison.InvariantCultureIgnoreCase))
                    {
                        string fileName = name.Replace("EPPlusTest.Resources.", "");
                        using (var stream = asm.GetManifestResourceStream(name))
                        using (var file = File.Create(Path.Combine(_clipartPath, fileName)))
                        {
                            stream.CopyTo(file);
                        }
                        break;
                    }
                }
            }
            _worksheetPath = Path.Combine(Path.GetTempPath(), @"EPPlus worksheets");
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
