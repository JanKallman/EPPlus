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
    public class ArgumentParserFactoryTests
    {
        private ArgumentParserFactory _parserFactory;

        [TestInitialize]
        public void Setup()
        {
            _parserFactory = new ArgumentParserFactory();
        }

        [TestMethod]
        public void ShouldReturnIntArgumentParserWhenDataTypeIsInteger()
        {
            var parser = _parserFactory.CreateArgumentParser(DataType.Integer);
            Assert.IsInstanceOfType(parser, typeof(IntArgumentParser));
        }

        [TestMethod]
        public void ShouldReturnBoolArgumentParserWhenDataTypeIsBoolean()
        {
            var parser = _parserFactory.CreateArgumentParser(DataType.Boolean);
            Assert.IsInstanceOfType(parser, typeof(BoolArgumentParser));
        }

        [TestMethod]
        public void ShouldReturnDoubleArgumentParserWhenDataTypeIsDecial()
        {
            var parser = _parserFactory.CreateArgumentParser(DataType.Decimal);
            Assert.IsInstanceOfType(parser, typeof(DoubleArgumentParser));
        }
    }
}
