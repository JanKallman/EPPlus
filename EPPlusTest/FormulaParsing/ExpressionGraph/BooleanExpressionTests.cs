using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class BooleanExpressionTests
    {
        [TestMethod]
        public void CompileShouldHandlePercent()
        {
            var exp1 = new BooleanExpression("TRUE");
            exp1.SetPercentage();
            var result = exp1.Compile();
            Assert.AreEqual(0.01, result.Result);
            Assert.AreEqual(DataType.Decimal, result.DataType);
        }
    }
}
