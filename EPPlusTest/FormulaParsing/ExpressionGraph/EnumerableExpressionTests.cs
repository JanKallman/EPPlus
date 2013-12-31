using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class EnumerableExpressionTests
    {
        [TestMethod]
        public void CompileShouldReturnEnumerableOfCompiledChildExpressions()
        {
            var expression = new EnumerableExpression();
            expression.AddChild(new IntegerExpression("2"));
            expression.AddChild(new IntegerExpression("3"));
            var result = expression.Compile();

            Assert.IsInstanceOfType(result.Result, typeof(IEnumerable<object>));
            var resultList = (IEnumerable<object>)result.Result;
            Assert.AreEqual(2d, resultList.ElementAt(0));
            Assert.AreEqual(3d, resultList.ElementAt(1));
        }
    }
}
