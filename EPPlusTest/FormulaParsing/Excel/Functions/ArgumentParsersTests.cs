using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class ArgumentParsersTests
    {
        [TestMethod]
        public void ShouldReturnSameInstanceOfIntParserWhenCalledTwice()
        {
            var parsers = new ArgumentParsers();
            var parser1 = parsers.GetParser(DataType.Integer);
            var parser2 = parsers.GetParser(DataType.Integer);
            Assert.AreEqual(parser1, parser2);
        }
    }
}
