using System;
using System.Text;
using System.Collections.Generic;
//using System.Diagnostics.Design;
using System.Globalization;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest
{
    [TestClass]
    public class ExpressionEvaluatorTests
    {
        private ExpressionEvaluator _evaluator;

        [TestInitialize]
        public void Setup()
        {
            _evaluator = new ExpressionEvaluator();
        }

        #region Numeric Expression Tests
        [TestMethod]
        public void EvaluateShouldReturnTrueIfOperandsAreEqual()
        {
            var result = _evaluator.Evaluate("1", "1");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateShouldReturnTrueIfOperandsAreMatchingButDifferentTypes()
        {
            var result = _evaluator.Evaluate(1d, "1");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateShouldEvaluateOperator()
        {
            var result = _evaluator.Evaluate(1d, "<2");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateShouldEvaluateNumericString()
        {
            var result = _evaluator.Evaluate("1", ">0");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateShouldHandleBooleanArg()
        {
            var result = _evaluator.Evaluate(true, "TRUE");
            Assert.IsTrue(result);
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void EvaluateShouldThrowIfOperatorIsNotBoolean()
        {
            var result = _evaluator.Evaluate(1d, "+1");
        }
        #endregion

        #region Date tests
        [TestMethod]
        public void EvaluateShouldHandleDateArg()
        {
            #if (!Core)
                Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            #endif
            var result = _evaluator.Evaluate(new DateTime(2016,6,28), "2016-06-28");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateShouldHandleDateArgWithOperator()
        {
#if (!Core)
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
#endif
            var result = _evaluator.Evaluate(new DateTime(2016, 6, 28), ">2016-06-27");
            Assert.IsTrue(result);
        }
#endregion

#region Blank Expression Tests
        [TestMethod]
        public void EvaluateBlankExpressionEqualsNull()
        {
            var result = _evaluator.Evaluate(null, "");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateBlankExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, "");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateBlankExpressionEqualsZero()
        {
            var result = _evaluator.Evaluate(0d, "");
            Assert.IsFalse(result);
        }
#endregion

#region Quotes Expression Tests
        [TestMethod]
        public void EvaluateQuotesExpressionEqualsNull()
        {
            var result = _evaluator.Evaluate(null, "\"\"");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateQuotesExpressionEqualsZero()
        {
            var result = _evaluator.Evaluate(0d, "\"\"");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateQuotesExpressionEqualsCharacter()
        {
            var result = _evaluator.Evaluate("a", "\"\"");
            Assert.IsFalse(result);
        }
#endregion

#region NotEqualToZero Expression Tests
        [TestMethod]
        public void EvaluateNotEqualToZeroExpressionEqualsNull()
        {
            var result = _evaluator.Evaluate(null, "<>0");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToZeroExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, "<>0");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToZeroExpressionEqualsCharacter()
        {
            var result = _evaluator.Evaluate("a", "<>0");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToZeroExpressionEqualsNonZero()
        {
            var result = _evaluator.Evaluate(1d, "<>0");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToZeroExpressionEqualsZero()
        {
            var result = _evaluator.Evaluate(0d, "<>0");
            Assert.IsFalse(result);
        }
#endregion

#region NotEqualToBlank Expression Tests
        [TestMethod]
        public void EvaluateNotEqualToBlankExpressionEqualsNull()
        {
            var result = _evaluator.Evaluate(null, "<>");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToBlankExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, "<>");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToBlankExpressionEqualsCharacter()
        {
            var result = _evaluator.Evaluate("a", "<>");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToBlankExpressionEqualsNonZero()
        {
            var result = _evaluator.Evaluate(1d, "<>");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateNotEqualToBlankExpressionEqualsZero()
        {
            var result = _evaluator.Evaluate(0d, "<>");
            Assert.IsTrue(result);
        }
#endregion

#region Character Expression Tests
        [TestMethod]
        public void EvaluateCharacterExpressionEqualNull()
        {
            var result = _evaluator.Evaluate(null, "a");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateCharacterExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, "a");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateCharacterExpressionEqualsNumeral()
        {
            var result = _evaluator.Evaluate(1d, "a");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateCharacterExpressionEqualsSameCharacter()
        {
            var result = _evaluator.Evaluate("a", "a");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateCharacterExpressionEqualsDifferentCharacter()
        {
            var result = _evaluator.Evaluate("b", "a");
            Assert.IsFalse(result);
        }
#endregion

#region CharacterWithOperator Expression Tests
        [TestMethod]
        public void EvaluateCharacterWithOperatorExpressionEqualNull()
        {
            var result = _evaluator.Evaluate(null, ">a");
            Assert.IsFalse(result);
            result = _evaluator.Evaluate(null, "<a");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateCharacterWithOperatorExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, ">a");
            Assert.IsFalse(result);
            result = _evaluator.Evaluate(string.Empty, "<a");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateCharacterWithOperatorExpressionEqualsNumeral()
        {
            var result = _evaluator.Evaluate(1d, ">a");
            Assert.IsFalse(result);
            result = _evaluator.Evaluate(1d, "<a");
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void EvaluateCharacterWithOperatorExpressionEqualsSameCharacter()
        {
            var result = _evaluator.Evaluate("a", ">a");
            Assert.IsFalse(result);
            result = _evaluator.Evaluate("a", ">=a");
            Assert.IsTrue(result);
            result = _evaluator.Evaluate("a", "<a");
            Assert.IsFalse(result);
            result = _evaluator.Evaluate("a", ">=a");
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateCharacterWithOperatorExpressionEqualsDifferentCharacter()
        {
            var result = _evaluator.Evaluate("b", ">a");
            Assert.IsTrue(result);
            result = _evaluator.Evaluate("b", "<a");
            Assert.IsFalse(result);
        }
#endregion
    }
}
