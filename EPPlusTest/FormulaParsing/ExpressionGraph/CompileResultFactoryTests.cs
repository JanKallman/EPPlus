using System.Globalization;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestClass]
    public class CompileResultFactoryTests
    {
#if (!Core)
        [TestMethod]
        public void CalculateUsingEuropeanDates()
        {
            var us = new CultureInfo("en-US");
            Thread.CurrentThread.CurrentCulture = us;
            var crf = new CompileResultFactory();
            var result = crf.Create("1/15/2014");
            var numeric = result.ResultNumeric;
            Assert.AreEqual(41654, numeric);
            var gb = new CultureInfo("en-GB");
            Thread.CurrentThread.CurrentCulture = gb;
            var euroResult = crf.Create("15/1/2014");
            var eNumeric = euroResult.ResultNumeric;
            Assert.AreEqual(41654, eNumeric);
        }
#endif
    }
}
