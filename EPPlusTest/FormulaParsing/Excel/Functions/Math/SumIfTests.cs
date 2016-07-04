using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestClass]
    public class SumIfTests
    {
        private ExcelPackage _package;
        private EpplusExcelDataProvider _provider;
        private ParsingContext _parsingContext;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _provider = new EpplusExcelDataProvider(_package);
            _parsingContext = ParsingContext.Create();
            _parsingContext.Scopes.NewScope(RangeAddress.Empty);
            _worksheet = _package.Workbook.Worksheets.Add("testsheet");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [TestMethod]
        public void SumIfNumeric()
        {
            _worksheet.Cells["A1"].Value = 1d;
            _worksheet.Cells["A2"].Value = 2d;
            _worksheet.Cells["A3"].Value = 3d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">1", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(8d, result.Result);
        }

        [TestMethod]
        public void SumIfNonNumeric()
        {
            _worksheet.Cells["A1"].Value = "Monday";
            _worksheet.Cells["A2"].Value = "Tuesday";
            _worksheet.Cells["A3"].Value = "Thursday";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "T*day", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(8d, result.Result);
        }

        [TestMethod]
        public void SumIfNumericExpression()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = 1d;
            _worksheet.Cells["A3"].Value = "Not Empty";
            var func = new SumIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, 1d);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void SumIfEqualToEmptyString()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(1d, result.Result);
        }

        [TestMethod]
        public void SumIfNotEqualToNull()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<>", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(8d, result.Result);
        }

        [TestMethod]
        public void SumIfEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 0d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfNotEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 0d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<>0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(4d, result.Result);
        }

        [TestMethod]
        public void SumIfGreaterThanZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 1d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfGreaterThanOrEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 1d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">=0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfLessThanZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = -1d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfLessThanOrEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = -1d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<=0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfLessThanCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void SumIfLessThanOrEqualToCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<=a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void SumIfGreaterThanCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfGreaterThanOrEqualToCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">=a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfHandleDates()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new SumIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">=a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(5d, result.Result);
        }

        [TestMethod]
        public void SumIfShouldHandleBooleanArg()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = true;
                sheet.Cells["B1"].Value = 1;
                sheet.Cells["A2"].Value = false;
                sheet.Cells["B2"].Value = 1;
                sheet.Cells["C1"].Formula = "SUMIF(A1:A2,TRUE,B1:B2)";
                sheet.Calculate();
                Assert.AreEqual(1d, sheet.Cells["C1"].Value);
            }
        }
    }
}
