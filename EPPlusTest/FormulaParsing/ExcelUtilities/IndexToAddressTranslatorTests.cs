using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rhino.Mocks;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class IndexToAddressTranslatorTests
    {
        private ExcelDataProvider _excelDataProvider;
        private IndexToAddressTranslator _indexToAddressTranslator;

        [TestInitialize]
        public void Setup()
        {
            SetupTranslator(12345, ExcelReferenceType.RelativeRowAndColumn);
        }

        private void SetupTranslator(int maxRows, ExcelReferenceType refType)
        {
            _excelDataProvider = MockRepository.GenerateStub<ExcelDataProvider>();
            _excelDataProvider.Stub(x => x.ExcelMaxRows).Return(maxRows);
            _indexToAddressTranslator = new IndexToAddressTranslator(_excelDataProvider, refType);
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ShouldThrowIfExcelDataProviderIsNull()
        {
            new IndexToAddressTranslator(null);
        }

        [TestMethod]
        public void ShouldTranslate0And0ToA1()
        {
            var result = _indexToAddressTranslator.ToAddress(0, 0);
            Assert.AreEqual("A1", result);
        }

        [TestMethod]
        public void ShouldTranslate26And0ToAA1()
        {
            var result = _indexToAddressTranslator.ToAddress(26, 0);
            Assert.AreEqual("AA1", result);
        }

        [TestMethod]
        public void ShouldTranslate26x26plus25And0ToZZ1()
        {
            var result = _indexToAddressTranslator.ToAddress(26*26+25, 0);
            Assert.AreEqual("ZZ1", result);
        }

        [TestMethod]
        public void ShouldTranslate26x26plus26And4ToAAA5()
        {
            var result = _indexToAddressTranslator.ToAddress(26 * 26 + 26, 4);
            Assert.AreEqual("AAA5", result);
        }

        [TestMethod]
        public void ShouldTranslateToEntireColumnWhenRowIsEqualToMaxRows()
        {
            _excelDataProvider.Stub(x => x.ExcelMaxRows).Return(123456);
            var result = _indexToAddressTranslator.ToAddress(0, 123456);
            Assert.AreEqual("A", result);
        }

        [TestMethod]
        public void ShouldTranslateToAbsoluteAddress()
        {
            SetupTranslator(123456, ExcelReferenceType.AbsoluteRowAndColumn);
            var result = _indexToAddressTranslator.ToAddress(0, 0);
            Assert.AreEqual("$A$1", result);
        }

        [TestMethod]
        public void ShouldTranslateToAbsoluteRowAndRelativeCol()
        {
            SetupTranslator(123456, ExcelReferenceType.AbsoluteRowRelativeColumn);
            var result = _indexToAddressTranslator.ToAddress(0, 0);
            Assert.AreEqual("A$1", result);
        }

        [TestMethod]
        public void ShouldTranslateToRelativeRowAndAbsoluteCol()
        {
            SetupTranslator(123456, ExcelReferenceType.RelativeRowAbsolutColumn);
            var result = _indexToAddressTranslator.ToAddress(0, 0);
            Assert.AreEqual("$A1", result);
        }

        [TestMethod]
        public void ShouldTranslateToRelativeRowAndCol()
        {
            SetupTranslator(123456, ExcelReferenceType.RelativeRowAndColumn);
            var result = _indexToAddressTranslator.ToAddress(0, 0);
            Assert.AreEqual("A1", result);
        }
    }
}
