using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using EPPlusTest.FormulaParsing.TestHelpers;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class NumberFunctionsTests
    {
        private ParsingContext _parsingContext = ParsingContext.Create();

        [TestMethod]
        public void CIntShouldConvertTextToInteger()
        {
            var func = new CInt();
            var args = FunctionsHelper.CreateArgs("2");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void IntShouldConvertDecimalToInteger()
        {
            var func = new CInt();
            var args = FunctionsHelper.CreateArgs(2.88m);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(2, result.Result);
        }

        [TestMethod]
        public void IntShouldConvertNegativeDecimalToInteger()
        {
            var func = new CInt();
            var args = FunctionsHelper.CreateArgs(-2.88m);
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(-3, result.Result);
        }

        [TestMethod]
        public void IntShouldConvertStringToInteger()
        {
            var func = new CInt();
            var args = FunctionsHelper.CreateArgs("-2.88");
            var result = func.Execute(args, _parsingContext);
            Assert.AreEqual(-3, result.Result);
        }
    }
}
