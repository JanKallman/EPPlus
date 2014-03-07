using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class CompileResultTests
    {
        [TestMethod]
        public void Test1()
        {
            var result = new CompileResult(0.031180782681731412d, DataType.Decimal);
            Assert.AreEqual(311807826817314.0d, result.ResultNumeric);
        }
    }
}
