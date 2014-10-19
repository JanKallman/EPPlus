using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using Rhino.Mocks;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestClass]
    public class StringFunctionsTests : FormulaParserTestBase
    {
        private ExcelDataProvider _provider;
        [TestInitialize]
        public void Setup()
        {
            _provider = MockRepository.GenerateStub<ExcelDataProvider>();
            _parser = new FormulaParser(_provider);
        }

        [TestMethod]
        public void TextShouldConcatenateWithNextExpression()
        {
            _provider.Stub(x => x.GetFormat(23.5, "$0.00")).Return("$23.50");
            var result = _parser.Parse("TEXT(23.5,\"$0.00\") & \" per hour\"");
            Assert.AreEqual("$23.50 per hour", result);
        }

        [TestMethod]
        public void LenShouldAddLengthUsingSuppliedOperator()
        {
            var result = _parser.Parse("Len('abc') + 2");
            Assert.AreEqual(5d, result);
        }

        [TestMethod]
        public void LowerShouldReturnALowerCaseString()
        {
            var result = _parser.Parse("Lower('ABC')");
            Assert.AreEqual("abc", result);
        }

        [TestMethod]
        public void UpperShouldReturnAnUpperCaseString()
        {
            var result = _parser.Parse("Upper('abc')");
            Assert.AreEqual("ABC", result);
        }

        [TestMethod]
        public void LeftShouldReturnSubstringFromLeft()
        {
            var result = _parser.Parse("Left('abacd', 2)");
            Assert.AreEqual("ab", result);
        }

        [TestMethod]
        public void RightShouldReturnSubstringFromRight()
        {
            var result = _parser.Parse("RIGHT('abacd', 2)");
            Assert.AreEqual("cd", result);
        }

        [TestMethod]
        public void MidShouldReturnSubstringAccordingToParams()
        {
            var result = _parser.Parse("Mid('abacd', 2, 2)");
            Assert.AreEqual("ba", result);
        }

        [TestMethod]
        public void ReplaceShouldReturnSubstringAccordingToParams()
        {
            var result = _parser.Parse("Replace('testar', 3, 3, 'hej')");
            Assert.AreEqual("tehejr", result);
        }

        [TestMethod]
        public void SubstituteShouldReturnSubstringAccordingToParams()
        {
            var result = _parser.Parse("Substitute('testar testar', 'es', 'xx')");
            Assert.AreEqual("txxtar txxtar", result);
        }

        [TestMethod]
        public void ConcatenateShouldReturnAccordingToParams()
        {
            var result = _parser.Parse("CONCATENATE('One', 'Two', 'Three')");
            Assert.AreEqual("OneTwoThree", result);
        }

        [TestMethod]
        public void TShouldReturnText()
        {
            var result = _parser.Parse("T('One')");
            Assert.AreEqual("One", result);
        }
    }
}
