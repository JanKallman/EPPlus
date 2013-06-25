using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.IntegrationTests.ExcelDataProviderTests
{
    [TestClass]
    public class ExcelDataProviderIntegrationTests
    {
        private ExcelCell CreateItem(object val, int row)
        {
            return new ExcelCell(val, null, 0, row);
        }

        [TestMethod]
        public void ShouldCallProviderInSumFunctionAndCalculateResult()
        {
            var expectedAddres = "A1:A2";
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            provider
                .Stub(x => x.GetRangeValues(expectedAddres))
                .Return(new ExcelCell[] { CreateItem(1, 0), CreateItem(2, 1) });
            var parser = new FormulaParser(provider);
            var result = parser.Parse(string.Format("sum({0})", expectedAddres));
            Assert.AreEqual(3d, result);
        }

        [TestMethod]
        public void ShouldExecuteFormulaInRange()
        {
            var expectedAddres = "A1:A2";
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            provider
                .Stub(x => x.GetRangeValues(expectedAddres))
                .Return(new ExcelCell[] { CreateItem(1, 0), new ExcelCell(null, "SUM(1,2)", 0, 1) });
            var parser = new FormulaParser(provider);
            var result = parser.Parse(string.Format("sum({0})", expectedAddres));
            Assert.AreEqual(4d, result);
        }

        [TestMethod, ExpectedException(typeof(CircularReferenceException))]
        public void ShouldHandleCircularReference2()
        {
            var expectedAddres = "A1:A2";
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            provider
                .Stub(x => x.GetRangeValues(expectedAddres))
                .Return(new ExcelCell[] { CreateItem(1, 0), new ExcelCell(null, "SUM(A1:A2)",0, 1) });
            var parser = new FormulaParser(provider);
            var result = parser.Parse(string.Format("sum({0})", expectedAddres));
        }
    }
}
