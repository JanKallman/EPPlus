using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class ExcelAddressExpressionTests
    {
        private ParsingContext _parsingContext;
        private ParsingScope _scope;

        private ExcelCell CreateItem(object val)
        {
            return new ExcelCell(val, null, 0, 0);
        }

        [TestInitialize]
        public void Setup()
        {
            _parsingContext = ParsingContext.Create();
            _scope = _parsingContext.Scopes.NewScope(RangeAddress.Empty);
        }

        [TestCleanup]
        public void Cleanup()
        {
            _scope.Dispose();
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ConstructorShouldThrowIfExcelDataProviderIsNull()
        {
            new ExcelAddressExpression("A1", null, _parsingContext);
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ConstructorShouldThrowIfParsingContextIsNull()
        {
            new ExcelAddressExpression("A1", A.Fake<ExcelDataProvider>(), null);
        }

        //TODO:Fix Test /Janne
        //[TestMethod]
        //public void ShouldCallReturnResultFromProvider()
        //{
        //    var expectedAddress = "A1";
        //    var provider = MockRepository.GenerateStub<ExcelDataProvider>();
        //    provider
        //        .Stub(x => x.GetRangeValues(string.Empty, expectedAddress))
        //        .Return(new object[]{ 1 });

        //    var expression = new ExcelAddressExpression(expectedAddress, provider, _parsingContext);
        //    var result = expression.Compile();
        //    Assert.AreEqual(1, result.Result);
        //}

        //TODO:Fix Test /Janne
        //[TestMethod]
        //public void CompileShouldReturnAddress()
        //{
        //    var expectedAddress = "A1";
        //    var provider = MockRepository.GenerateStub<ExcelDataProvider>();
        //    provider
        //        .Stub(x => x.GetRangeValues(expectedAddress))
        //        .Return(new ExcelCell[] { CreateItem(1) });

        //    var expression = new ExcelAddressExpression(expectedAddress, provider, _parsingContext);
        //    expression.ParentIsLookupFunction = true;
        //    var result = expression.Compile();
        //    Assert.AreEqual(expectedAddress, result.Result);

        //}

        #region Compile Tests
        [TestMethod]
        public void CompileSingleCellReference()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("A1", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    Assert.IsNull(result.Result);
                }
            }
        }

        [TestMethod]
        public void CompileSingleCellReferenceWithValue()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                sheet.Cells[1, 1].Value = "Value";
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("A1", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    Assert.AreEqual("Value", result.Result);
                }
            }
        }

        [TestMethod]
        public void CompileSingleCellReferenceResolveToRange()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("A1", excelDataProvider, parsingContext);
                    expression.ResolveAsRange = true;
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.IsNotNull(rangeInfo);
                    Assert.AreEqual("A1", rangeInfo.Address.Address);
                }
            }
        }

        [TestMethod]
        public void CompileSingleCellReferenceResolveToRangeColumnAbsolute()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("$A1", excelDataProvider, parsingContext);
                    expression.ResolveAsRange = true;
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.IsNotNull(rangeInfo);
                    Assert.AreEqual("$A1", rangeInfo.Address.Address);
                }
            }
        }

        [TestMethod]
        public void CompileSingleCellReferenceResolveToRangeRowAbsolute()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("$A1", excelDataProvider, parsingContext);
                    expression.ResolveAsRange = true;
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.IsNotNull(rangeInfo);
                    Assert.AreEqual("$A1", rangeInfo.Address.Address);
                }
            }
        }

        [TestMethod]
        public void CompileSingleCellReferenceResolveToRangeAbsolute()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("$A$1", excelDataProvider, parsingContext);
                    expression.ResolveAsRange = true;
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.IsNotNull(rangeInfo);
                    Assert.AreEqual("$A$1", rangeInfo.Address.Address);
                }
            }
        }

        [TestMethod]
        public void CompileMultiCellReference()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("A1:A5", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.IsNotNull(rangeInfo);
                    Assert.AreEqual("A1:A5", rangeInfo.Address.Address);
                    // Enumerating the range still yields no results.
                    Assert.AreEqual(0, rangeInfo.Count());
                }
            }
        }

        [TestMethod]
        public void CompileMultiCellReferenceWithValues()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                sheet.Cells[1, 1].Value = "Value1";
                sheet.Cells[2, 1].Value = "Value2";
                sheet.Cells[3, 1].Value = "Value3";
                sheet.Cells[4, 1].Value = "Value4";
                sheet.Cells[5, 1].Value = "Value5";
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("A1:A5", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.IsNotNull(rangeInfo);
                    Assert.AreEqual("A1:A5", rangeInfo.Address.Address);
                    Assert.AreEqual(5, rangeInfo.Count());
                    for (int i = 1; i <= 5; i++)
                    {
                        var rangeItem = rangeInfo.ElementAt(i - 1);
                        Assert.AreEqual("Value" + i, rangeItem.Value);
                    }
                }
            }
        }

        [TestMethod]
        public void CompileMultiCellReferenceColumnAbsolute()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("$A1:$A5", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.IsNotNull(rangeInfo);
                    Assert.AreEqual("$A1:$A5", rangeInfo.Address.Address);
                    // Enumerating the range still yields no results.
                    Assert.AreEqual(0, rangeInfo.Count());
                }
            }
        }

        [TestMethod]
        public void CompileMultiCellReferenceRowAbsolute()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("A$1:A$5", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.IsNotNull(rangeInfo);
                    Assert.AreEqual("A$1:A$5", rangeInfo.Address.Address);
                    // Enumerating the range still yields no results.
                    Assert.AreEqual(0, rangeInfo.Count());
                }
            }
        }

        [TestMethod]
        public void CompileMultiCellReferenceAbsolute()
        {
            var parsingContext = ParsingContext.Create();
            var file = new FileInfo("filename.xlsx");
            using (var package = new ExcelPackage(file))
            using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
            using (var excelDataProvider = new EpplusExcelDataProvider(package))
            {
                var rangeAddressFactory = new RangeAddressFactory(excelDataProvider);
                using (parsingContext.Scopes.NewScope(rangeAddressFactory.Create("NewSheet", 3, 3)))
                {
                    var expression = new ExcelAddressExpression("$A$1:$A$5", excelDataProvider, parsingContext);
                    var result = expression.Compile();
                    var rangeInfo = result.Result as ExcelDataProvider.IRangeInfo;
                    Assert.IsNotNull(rangeInfo);
                    Assert.AreEqual("$A$1:$A$5", rangeInfo.Address.Address);
                    // Enumerating the range still yields no results.
                    Assert.AreEqual(0, rangeInfo.Count());
                }
            }
        }
        #endregion
    }
}
