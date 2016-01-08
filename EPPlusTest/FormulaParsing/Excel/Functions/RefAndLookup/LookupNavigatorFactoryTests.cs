using System;
using System.IO;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class LookupNavigatorFactoryTests
    {
        private ExcelPackage _excelPackage;
        private ParsingContext _context;

        [TestInitialize]
        public void Initialize()
        {
            _excelPackage = new ExcelPackage(new MemoryStream());
            _excelPackage.Workbook.Worksheets.Add("Test");
            _context = ParsingContext.Create();
            _context.ExcelDataProvider = new EpplusExcelDataProvider(_excelPackage);
            _context.Scopes.NewScope(RangeAddress.Empty);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _excelPackage.Dispose();
        }

        [TestMethod]
        public void Should_Return_ExcelLookupNavigator_When_Range_Is_Set()
        {
            var args = new LookupArguments(FunctionsHelper.CreateArgs(8, "A:B", 1));
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Horizontal, args, _context);
            Assert.IsInstanceOfType(navigator, typeof(ExcelLookupNavigator));
        }

        [TestMethod]
        public void Should_Return_ArrayLookupNavigator_When_Array_Is_Supplied()
        {
            var args = new LookupArguments(FunctionsHelper.CreateArgs(8, FunctionsHelper.CreateArgs(1,2), 1));
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Horizontal, args, _context);
            Assert.IsInstanceOfType(navigator, typeof(ArrayLookupNavigator));
        }
    }
}
