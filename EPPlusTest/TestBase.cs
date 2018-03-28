using System.Diagnostics.CodeAnalysis;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;

namespace EPPlusTest
{
    [TestClass]
    public abstract class TestBase
    {
        protected ExcelPackage _pck;

        [SuppressMessage("ReSharper", "UnusedAutoPropertyAccessor.Global", Justification = "MSStest needs this setter to be public")]
        [SuppressMessage("ReSharper", "MemberCanBeProtected.Global", Justification = "MSStest needs this setter to be public")]
        public TestContext TestContext { get; set; }

        [TestInitialize]
        public void InitBase()
        {
            _pck = new ExcelPackage();
        }
        
        protected void SaveWorksheet(string name)
        {
            if (_pck.Workbook.Worksheets.Count == 0) return;
            var fi = new FileInfo(Path.Combine(Scaffolding.WorksheetPath, name));
            if (fi.Exists) {
                fi.Delete();
            }
            _pck.SaveAs(fi);
        }
    }
}

