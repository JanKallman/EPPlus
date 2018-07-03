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
        protected string _worksheetPath= @"c:\epplusTest\Testoutput\";
        protected string _testInputPath = @"c:\epplusTest\workbooks\";
        public TestContext TestContext { get; set; }
        
        [TestInitialize]
        public void InitBase()
        {
            _clipartPath = Path.Combine(Path.GetTempPath(), @"EPPlus clipart");
            if (!Directory.Exists(_clipartPath))
            {
                Directory.CreateDirectory(_clipartPath);
            }
            if(Environment.GetEnvironmentVariable("EPPlusTestInputPath")!=null)
            {
                _testInputPath = Environment.GetEnvironmentVariable("EPPlusTestInputPath");
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
                    if (name.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
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
            
            //_worksheetPath = Path.Combine(Path.GetTempPath(), @"EPPlus worksheets");
            //if (!Directory.Exists(_worksheetPath))
            //{
            //    Directory.CreateDirectory(_worksheetPath);
            //}
            var di=new DirectoryInfo(_worksheetPath);            
            _worksheetPath = di.FullName + "\\";

            _pck = new ExcelPackage();
        }

        protected ExcelPackage OpenPackage(string name, bool delete=false)
        {
            var fi = new FileInfo(_worksheetPath + name);
            if(delete && fi.Exists)
            {
                fi.Delete();
            }
            _pck = new ExcelPackage(fi);
            return _pck;
        }
        protected ExcelPackage OpenTemplatePackage(string name)
        {
            var t = new FileInfo(_testInputPath + name);
            if (t.Exists)
            {
                var fi = new FileInfo(_worksheetPath + name);
                _pck = new ExcelPackage(fi, t);
            }
            else
            {
                Assert.Inconclusive($"Template {name} does not exist in path {_testInputPath}");
            }
            return _pck;
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
