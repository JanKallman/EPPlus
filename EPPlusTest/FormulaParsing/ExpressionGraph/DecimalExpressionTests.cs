using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class DecimalExpressionTests
    {
        [TestMethod]
        public void CompileShouldHandlePercent()
        {
            var exp1 = new DecimalExpression("1");
            exp1.SetPercentage();
            var result = exp1.Compile();
            Assert.AreEqual(0.01, result.Result);
        }
    }
}
