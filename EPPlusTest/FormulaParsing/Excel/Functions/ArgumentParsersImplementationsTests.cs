using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace EPPlusTest.Excel.Functions
{
    [TestClass]
    public class ArgumentParsersImplementationsTests
    {
        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void IntParserShouldThrowIfArgumentIsNull()
        {
            var parser = new IntArgumentParser();
            parser.Parse(null);
        }

        [TestMethod]
        public void IntParserShouldConvertToAnInteger()
        {
            var parser = new IntArgumentParser();
            var result = parser.Parse(3);
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void IntParserShouldConvertADoubleToAnInteger()
        {
            var parser = new IntArgumentParser();
            var result = parser.Parse(3d);
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void IntParserShouldConvertAStringValueToAnInteger()
        {
            var parser = new IntArgumentParser();
            var result = parser.Parse("3");
            Assert.AreEqual(3, result);
        }

        [TestMethod]
        public void BoolParserShouldConvertNullToFalse()
        {
            var parser = new BoolArgumentParser();
            var result = (bool)parser.Parse(null);
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void BoolParserShouldConvertStringValueTrueToBoolValueTrue()
        {
            var parser = new BoolArgumentParser();
            var result = (bool)parser.Parse("true");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void BoolParserShouldConvert0ToFalse()
        {
            var parser = new BoolArgumentParser();
            var result = (bool)parser.Parse(0);
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void BoolParserShouldConvert1ToTrue()
        {
            var parser = new BoolArgumentParser();
            var result = (bool)parser.Parse(0);
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void DoubleParserShouldConvertDoubleToDouble()
        {
            var parser = new DoubleArgumentParser();
            var result = parser.Parse(3d);
            Assert.AreEqual(3d, result);
        }

        [TestMethod]
        public void DoubleParserShouldConvertIntToDouble()
        {
            var parser = new DoubleArgumentParser();
            var result = parser.Parse(3);
            Assert.AreEqual(3d, result);
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void DoubleParserShouldThrowIfArgumentIsNull()
        {
            var parser = new DoubleArgumentParser();
            parser.Parse(null);
        }

        [TestMethod]
        public void DoubleParserConvertStringToDouble()
        {
            var parser = new DoubleArgumentParser();
            var result = parser.Parse("3.3");
            Assert.AreEqual(3.3d, result);
        }

    }
}
