using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.Excel.Functions.Text
{
    [TestClass]
    public class TextFunctionsTests
    {
        private ParsingContext _parsingContext = ParsingContext.Create();

        [TestMethod]
        public void CStrShouldConvertNumberToString()
        {
            var func = new CStr();
            var result = func.Execute(FunctionsHelper.CreateArgs(1), _parsingContext);
            Assert.AreEqual(DataType.String, result.DataType);
            Assert.AreEqual("1", result.Result);
        }

        [TestMethod]
        public void LenShouldReturnStringsLength()
        {
            var func = new Len();
            var result = func.Execute(FunctionsHelper.CreateArgs("abc"), _parsingContext);
            Assert.AreEqual(3d, result.Result);
        }

        [TestMethod]
        public void LowerShouldReturnLowerCaseString()
        {
            var func = new Lower();
            var result = func.Execute(FunctionsHelper.CreateArgs("ABC"), _parsingContext);
            Assert.AreEqual("abc", result.Result);
        }

        [TestMethod]
        public void UpperShouldReturnUpperCaseString()
        {
            var func = new Upper();
            var result = func.Execute(FunctionsHelper.CreateArgs("abc"), _parsingContext);
            Assert.AreEqual("ABC", result.Result);
        }

        [TestMethod]
        public void LeftShouldReturnSubstringFromLeft()
        {
            var func = new Left();
            var result = func.Execute(FunctionsHelper.CreateArgs("abcd", 2), _parsingContext);
            Assert.AreEqual("ab", result.Result);
        }

        [TestMethod]
        public void RightShouldReturnSubstringFromRight()
        {
            var func = new Right();
            var result = func.Execute(FunctionsHelper.CreateArgs("abcd", 2), _parsingContext);
            Assert.AreEqual("cd", result.Result);
        }

        [TestMethod]
        public void MidShouldReturnSubstringAccordingToParams()
        {
            var func = new Mid();
            var result = func.Execute(FunctionsHelper.CreateArgs("abcd", 1, 2), _parsingContext);
            Assert.AreEqual("bc", result.Result);
        }

        [TestMethod]
        public void ReplaceShouldReturnAReplacedStringAccordingToParamsWhenStartIxIs1()
        {
            var func = new Replace();
            var result = func.Execute(FunctionsHelper.CreateArgs("testar", 1, 2, "hej"), _parsingContext);
            Assert.AreEqual("hejstar", result.Result);
        }

        [TestMethod]
        public void ReplaceShouldReturnAReplacedStringAccordingToParamsWhenStartIxIs3()
        {
            var func = new Replace();
            var result = func.Execute(FunctionsHelper.CreateArgs("testar", 3, 3, "hej"), _parsingContext);
            Assert.AreEqual("tehejr", result.Result);
        }

        [TestMethod]
        public void SubstituteShouldReturnAReplacedStringAccordingToParamsWhen()
        {
            var func = new Substitute();
            var result = func.Execute(FunctionsHelper.CreateArgs("testar testar", "es", "xx"), _parsingContext);
            Assert.AreEqual("txxtar txxtar", result.Result);
        }

        [TestMethod]
        public void ConcatenateShouldConcatenateThreeStrings()
        {
            var func = new Concatenate();
            var result = func.Execute(FunctionsHelper.CreateArgs("One", "Two", "Three"), _parsingContext);
            Assert.AreEqual("OneTwoThree", result.Result);
        }
    }
}
