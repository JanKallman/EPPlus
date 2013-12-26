using System;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class IndexTests
    {
        private ParsingContext _parsingContext;

        [TestInitialize]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
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
            Assert.AreEqual(5d, result.Result);
        }
    }
}
