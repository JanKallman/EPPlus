using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using FakeItEasy;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using AddressFunction = OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Address;
using static OfficeOpenXml.FormulaParsing.EpplusExcelDataProvider;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class RefAndLookupTests
    {
        const string WorksheetName = null;
        [TestMethod]
        public void LookupArgumentsShouldSetSearchedValue()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.AreEqual(1, lookupArgs.SearchedValue);
        }

        [TestMethod]
        public void LookupArgumentsShouldSetRangeAddress()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.AreEqual("A:B", lookupArgs.RangeAddress);
        }

        [TestMethod]
        public void LookupArgumentsShouldSetColIndex()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.AreEqual(2, lookupArgs.LookupIndex);
        }

        [TestMethod]
        public void LookupArgumentsShouldSetRangeLookupToTrueAsDefaultValue()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.IsTrue(lookupArgs.RangeLookup);
        }

        [TestMethod]
        public void LookupArgumentsShouldSetRangeLookupToTrueWhenTrueIsSupplied()
        {
            var args = FunctionsHelper.CreateArgs(1, "A:B", 2, true);
            var lookupArgs = new LookupArguments(args, ParsingContext.Create());
            Assert.IsTrue(lookupArgs.RangeLookup);
        }

        [TestMethod]
        public void VLookupShouldReturnResultFromMatchingRow()
        {
            var func = new VLookup();
            var args = FunctionsHelper.CreateArgs(2, "A1:B2", 2);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);
            
            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns(2);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 2)).Returns(5);
            A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(100,10));

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual(5, result.Result);
        }

        [TestMethod]
        public void VLookupShouldReturnClosestValueBelowWhenRangeLookupIsTrue()
        {
            var func = new VLookup();
            var args = FunctionsHelper.CreateArgs(4, "A1:B2", 2, true);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns(5);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 2)).Returns(4);

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual(1, result.Result);
        }

        [TestMethod]
        public void VLookupShouldReturnClosestStringValueBelowWhenRangeLookupIsTrue()
        {
            var func = new VLookup();
            var args = FunctionsHelper.CreateArgs("B", "A1:B2", 2, true);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();;

            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns("A");
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns("C");
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(4);

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual(1, result.Result);
        }

        [TestMethod]
        public void HLookupShouldReturnResultFromMatchingRow()
        {
            var func = new HLookup();
            var args = FunctionsHelper.CreateArgs(2, "A1:B2", 2);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();

            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(2);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(5);

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual(5, result.Result);
        }

        [TestMethod]
        public void HLookupShouldReturnNaErrorIfNoMatchingRecordIsFoundWhenRangeLookupIsFalse()
        {
            var func = new HLookup();
            var args = FunctionsHelper.CreateArgs(2, "A1:B2", 2, false);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();

            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(2);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(5);

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            var expectedResult = ExcelErrorValue.Create(eErrorType.NA);
            Assert.AreEqual(expectedResult, result.Result);
        }

        [TestMethod]
        public void HLookupShouldReturnErrorIfNoMatchingRecordIsFoundWhenRangeLookupIsTrue()
        {
            var func = new HLookup();
            var args = FunctionsHelper.CreateArgs(1, "A1:B2", 2, true);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();

            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 1)).Returns(2);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 1, 2)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName, 2, 2)).Returns(5);

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual(result.DataType, DataType.ExcelError);
        }

        [TestMethod]
        public void LookupShouldReturnResultFromMatchingRowArrayVertical()
        {
            var func = new Lookup();
            var args = FunctionsHelper.CreateArgs(4, "A1:B3", 2);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns("A");
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 2)).Returns("B");
            A.CallTo(() => provider.GetCellValue(WorksheetName,3, 1)).Returns(5);
            A.CallTo(() => provider.GetCellValue(WorksheetName,3, 2)).Returns("C");
            A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(100, 10));

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual("B", result.Result);
        }

        [TestMethod]
        public void LookupShouldReturnResultFromMatchingRowArrayHorizontal()
        {
            var func = new Lookup();
            var args = FunctionsHelper.CreateArgs(4, "A1:C2", 2);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 3)).Returns(5);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns("A");
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 2)).Returns("B");
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 3)).Returns("C");

            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual("B", result.Result);
        }

        [TestMethod]
        public void LookupShouldReturnResultFromMatchingSecondArrayHorizontal()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 3;
                sheet.Cells["C1"].Value = 5;
                sheet.Cells["A3"].Value = "A";
                sheet.Cells["B3"].Value = "B";
                sheet.Cells["C3"].Value = "C";

                sheet.Cells["D1"].Formula = "LOOKUP(4, A1:C1, A3:C3)";
                sheet.Calculate();
                var result = sheet.Cells["D1"].Value;
                Assert.AreEqual("B", result);

            }
        }

        [TestMethod]
        public void LookupShouldReturnResultFromMatchingSecondArrayHorizontalWithOffset()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B1"].Value = 3;
                sheet.Cells["C1"].Value = 5;
                sheet.Cells["B3"].Value = "A";
                sheet.Cells["C3"].Value = "B";
                sheet.Cells["D3"].Value = "C";

                sheet.Cells["D1"].Formula = "LOOKUP(4, A1:C1, B3:D3)";
                sheet.Calculate();
                var result = sheet.Cells["D1"].Value;
                Assert.AreEqual("B", result);

            } 
        }

        [TestMethod]
        public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeExact()
        {
            var func = new Match();
            var args = FunctionsHelper.CreateArgs(3, "A1:C1", 0);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 3)).Returns(5);
            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void MatchShouldReturnIndexOfMatchingValVertical_MatchTypeExact()
        {
            var func = new Match();
            var args = FunctionsHelper.CreateArgs(3, "A1:A3", 0);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,2, 1)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,3, 1)).Returns(5);
            A.CallTo(() => provider.GetDimensionEnd(A<string>.Ignored)).Returns(new ExcelCellAddress(100, 10));
            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestBelow()
        {
            var func = new Match();
            var args = FunctionsHelper.CreateArgs(4, "A1:C1", 1);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(1);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(3);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 3)).Returns(5);
            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void MatchShouldReturnIndexOfMatchingValHorizontal_MatchTypeClosestAbove()
        {
            var func = new Match();
            var args = FunctionsHelper.CreateArgs(6, "A1:C1", -1);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(10);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(8);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 3)).Returns(5);
            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void MatchShouldReturnFirstItemWhenExactMatch_MatchTypeClosestAbove()
        {
            var func = new Match();
            var args = FunctionsHelper.CreateArgs(10, "A1:C1", -1);
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);

            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 1)).Returns(10);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 2)).Returns(8);
            A.CallTo(() => provider.GetCellValue(WorksheetName,1, 3)).Returns(5);
            parsingContext.ExcelDataProvider = provider;
            var result = func.Execute(args, parsingContext);
            Assert.AreEqual(1, result.Result);
        }

        [TestMethod]
        public void MatchShouldHandleAddressOnOtherSheet()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
                sheet1.Cells["A1"].Formula = "Match(10, Sheet2!A1:Sheet2!A3, 0)";
                sheet2.Cells["A1"].Value = 9;
                sheet2.Cells["A2"].Value = 10;
                sheet2.Cells["A3"].Value = 11;
                sheet1.Calculate();
                Assert.AreEqual(2, sheet1.Cells["A1"].Value);
            }    
        }

        [TestMethod]
        public void RowShouldReturnRowFromCurrentScopeIfNoAddressIsSupplied()
        {
            var func = new Row();
            var parsingContext = ParsingContext.Create();
            var rangeAddressFactory = new RangeAddressFactory(A.Fake<ExcelDataProvider>());
            parsingContext.Scopes.NewScope(rangeAddressFactory.Create("A2"));
            var result = func.Execute(Enumerable.Empty<FunctionArgument>(), parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void RowShouldReturnRowSuppliedAddress()
        {
            var func = new Row();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("A3"), parsingContext);
            Assert.AreEqual(3, result.Result);
        }

        [TestMethod]
        public void ColumnShouldReturnRowFromCurrentScopeIfNoAddressIsSupplied()
        {
            var func = new Column();
            var parsingContext = ParsingContext.Create();
            var rangeAddressFactory = new RangeAddressFactory(A.Fake<ExcelDataProvider>());
            parsingContext.Scopes.NewScope(rangeAddressFactory.Create("B2"));
            var result = func.Execute(Enumerable.Empty<FunctionArgument>(), parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void ColumnShouldReturnRowSuppliedAddress()
        {
            var func = new Column();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("E3"), parsingContext);
            Assert.AreEqual(5, result.Result);
        }

        [TestMethod]
        public void RowsShouldReturnNbrOfRowsSuppliedRange()
        {
            var func = new Rows();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("A1:B3"), parsingContext);
            Assert.AreEqual(3, result.Result);
        }

        [TestMethod]
        public void RowsShouldReturnNbrOfRowsForEntireColumn()
        {
            var func = new Rows();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("A:B"), parsingContext);
            Assert.AreEqual(1048576, result.Result);
        }

        [TestMethod]
        public void ColumnssShouldReturnNbrOfRowsSuppliedRange()
        {
            var func = new Columns();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            var result = func.Execute(FunctionsHelper.CreateArgs("A1:E3"), parsingContext);
            Assert.AreEqual(5, result.Result);
        }

        [TestMethod]
        public void ChooseShouldReturnItemByIndex()
        {
            var func = new Choose();
            var parsingContext = ParsingContext.Create();
            var result = func.Execute(FunctionsHelper.CreateArgs(1, "A", "B"), parsingContext);
            Assert.AreEqual("A", result.Result);
        }

        [TestMethod]
        public void AddressShouldReturnAddressByIndexWithDefaultRefType()
        {
            var func = new AddressFunction();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
            var result = func.Execute(FunctionsHelper.CreateArgs(1, 2), parsingContext);
            Assert.AreEqual("$B$1", result.Result);
        }

        [TestMethod]
        public void AddressShouldReturnAddressByIndexWithRelativeType()
        {
            var func = new AddressFunction();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
            var result = func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn), parsingContext);
            Assert.AreEqual("B1", result.Result);
        }

        [TestMethod]
        public void AddressShouldReturnAddressByWithSpecifiedWorksheet()
        {
            var func = new AddressFunction();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
            var result = func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn, true, "Worksheet1"), parsingContext);
            Assert.AreEqual("Worksheet1!B1", result.Result);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void AddressShouldThrowIfR1C1FormatIsSpecified()
        {
            var func = new AddressFunction();
            var parsingContext = ParsingContext.Create();
            parsingContext.ExcelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => parsingContext.ExcelDataProvider.ExcelMaxRows).Returns(10);
            var result = func.Execute(FunctionsHelper.CreateArgs(1, 2, (int)ExcelReferenceType.RelativeRowAndColumn, false), parsingContext);
        }
    }
}
