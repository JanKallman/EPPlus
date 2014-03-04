using System;
using System.IO;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class IndexTests
    {
        private ParsingContext _parsingContext;
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
            _package = new ExcelPackage(new MemoryStream());
            _worksheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }
        
        [TestMethod]
        public void Index_Should_Return_Value_By_Index()
        {
            var func = new Index();
            var result = func.Execute(
                FunctionsHelper.CreateArgs(
                    FunctionsHelper.CreateArgs(1, 2, 5),
                    3
                    ),_parsingContext);
            Assert.AreEqual(5, result.Result);
        }

        [TestMethod, ExpectedException(typeof(ExcelErrorValueException))]
        public void Index_Should_Throw_Exception_When_NonNumeric_Result()
        {
            var func = new Index();
            var result = func.Execute(
                FunctionsHelper.CreateArgs(
                    FunctionsHelper.CreateArgs(1, 2, "a"),
                    3
                    ), _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void Index_Should_Handle_SingleRange()
        {
            _worksheet.Cells["A1"].Value = 1d;
            _worksheet.Cells["A2"].Value = 3d;
            _worksheet.Cells["A3"].Value = 5d;

            _worksheet.Cells["A4"].Formula = "INDEX(A1:A3;3)";

            _worksheet.Calculate();

            Assert.AreEqual(5d, _worksheet.Cells["A4"].Value);
        }
    }
}
