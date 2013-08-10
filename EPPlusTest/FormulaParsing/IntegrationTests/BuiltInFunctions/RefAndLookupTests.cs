using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class RefAndLookupTests : FormulaParserTestBase
    {
        private ExcelDataProvider _excelDataProvider;
        const string WorksheetName = "";

        [TestInitialize]
        public void Setup()
        {
            _excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(_excelDataProvider);
        }

        [TestMethod]
        public void VLookupShouldReturnCorrespondingValue()
        {
            var lookupAddress = "A1:B2";
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,0, 0)).Return(new ExcelCell(3, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,0, 1)).Return(new ExcelCell(1, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 0)).Return(new ExcelCell(2, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 1)).Return(new ExcelCell(5, null, 0, 0));
            var result = _parser.Parse("VLOOKUP(2, " + lookupAddress + ", 2)");
            Assert.AreEqual(5, result);
        }

        [TestMethod]
        public void VLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
        {
            var lookupAddress = "A1:B2";
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,0, 0)).Return(new ExcelCell(3, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,0, 1)).Return(new ExcelCell(1, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 0)).Return(new ExcelCell(5, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 1)).Return(new ExcelCell(5, null, 0, 0));
            var result = _parser.Parse("VLOOKUP(4, " + lookupAddress + ", 2, true)");
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void HLookupShouldReturnCorrespondingValue()
        {
            var lookupAddress = "A1:B2";
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,0, 0)).Return(new ExcelCell(3, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,0, 1)).Return(new ExcelCell(1, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 0)).Return(new ExcelCell(2, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 1)).Return(new ExcelCell(5, null, 0, 0));
            var result = _parser.Parse("HLOOKUP(1, " + lookupAddress + ", 2)");
            Assert.AreEqual(5, result);
        }

        [TestMethod]
        public void HLookupShouldReturnClosestValueBelowIfLastArgIsTrue()
        {
            var lookupAddress = "A1:B2";
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,0, 0)).Return(new ExcelCell(3, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,0, 1)).Return(new ExcelCell(5, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 0)).Return(new ExcelCell(1, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 1)).Return(new ExcelCell(2, null, 0, 0));
            var result = _parser.Parse("HLOOKUP(4, " + lookupAddress + ", 2, true)");
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void LookupShouldReturnMatchingValue()
        {
            var lookupAddress = "A1:B2";
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,0, 0)).Return(new ExcelCell(3, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,0, 1)).Return(new ExcelCell(5, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 0)).Return(new ExcelCell(4, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,1, 1)).Return(new ExcelCell(1, null, 0, 0));
            var result = _parser.Parse("LOOKUP(4, " + lookupAddress + ")");
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void MatchShouldReturnIndexOfMatchingValue()
        {
            var lookupAddress = "A1:A2";
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,0, 0)).Return(new ExcelCell(3, null, 0, 0));
            _excelDataProvider.Stub(x => x.GetCellValue(WorksheetName,0, 1)).Return(new ExcelCell(5, null, 0, 0));
            var result = _parser.Parse("MATCH(3, " + lookupAddress + ")");
            Assert.AreEqual(1, result);
        }

        [TestMethod]
        public void RowShouldReturnRowNumber()
        {
            _excelDataProvider.Stub(x => x.GetRangeValues("A4")).Return(new List<ExcelCell> { new ExcelCell(null, "Row()", 0, 0) });
            var result = _parser.ParseAt("A4");
            Assert.AreEqual(4, result);
        }

        [TestMethod]
        public void ColumnShouldReturnRowNumber()
        {
            _excelDataProvider.Stub(x => x.GetRangeValues("B4")).Return(new List<ExcelCell> { new ExcelCell(null, "Column()", 0, 0) });
            var result = _parser.ParseAt("B4");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void RowsShouldReturnNbrOfRows()
        {
            _excelDataProvider.Stub(x => x.GetRangeValues("A4")).Return(new List<ExcelCell> { new ExcelCell(null, "Rows(A5:B7)", 0, 0) });
            var result = _parser.ParseAt("A4");
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void ColumnsShouldReturnNbrOfCols()
        {
            _excelDataProvider.Stub(x => x.GetRangeValues("A4")).Return(new List<ExcelCell> { new ExcelCell(null, "Columns(A5:B7)", 0, 0) });
            var result = _parser.ParseAt("A4");
            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void ChooseShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Choose(1, 'A', 'B')");
            Assert.AreEqual("A", result);
        }

        [TestMethod]
        public void AddressShouldReturnCorrectResult()
        {
            _excelDataProvider.Stub(x => x.ExcelMaxRows).Return(12345);
            var result = _parser.Parse("Address(1, 1)");
            Assert.AreEqual("$A$1", result);
        }
    }
}
