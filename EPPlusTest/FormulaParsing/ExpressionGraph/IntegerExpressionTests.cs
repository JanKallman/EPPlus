using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Operators;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class IntegerExpressionTests
    {
        [TestMethod]
        public void MergeWithNextWithPlusOperatorShouldCalulateSumCorrectly()
        {
            var exp1 = new IntegerExpression("1");
            exp1.Operator = Operator.Plus;
            var exp2 = new IntegerExpression("2");
            exp1.Next = exp2;

            var result = exp1.MergeWithNext();

            Assert.AreEqual(3d, result.Compile().Result);
        }

        [TestMethod]
        public void MergeWithNextWithPlusOperatorShouldSetNextPointer()
        {
            var exp1 = new IntegerExpression("1");
            exp1.Operator = Operator.Plus;
            var exp2 = new IntegerExpression("2");
            exp1.Next = exp2;

            var result = exp1.MergeWithNext();

            Assert.IsNull(result.Next);
        }

        //[TestMethod]
        //public void CompileShouldHandlePercent()
        //{
        //    var exp1 = new IntegerExpression("1");
        //    exp1.Operator = Operator.Percent;
        //    exp1.Next = ConstantExpressions.Percent;
        //    var result = exp1.Compile();
        //    Assert.AreEqual(0.01, result.Result);
        //    Assert.AreEqual(DataType.Decimal, result.DataType);
        //}
    }
}
