﻿using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class AddressTranslatorTests
    {
        private AddressTranslator _addressTranslator;
        private ExcelDataProvider _excelDataProvider;
        private const int ExcelMaxRows = 1356;

        [TestInitialize]
        public void Setup()
        {
            _excelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => _excelDataProvider.ExcelMaxRows).Returns(ExcelMaxRows);
            _addressTranslator = new AddressTranslator(_excelDataProvider);
        }

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ConstructorShouldThrowIfProviderIsNull()
        {
            new AddressTranslator(null);
        }

        [TestMethod]
        public void ShouldTranslateRowNumber()
        {
            int col, row;
            _addressTranslator.ToColAndRow("A2", out col, out row);
            Assert.AreEqual(2, row);
        }

        [TestMethod]
        public void ShouldTranslateLettersToColumnIndex()
        {
            int col, row;
            _addressTranslator.ToColAndRow("C1", out col, out row);
            Assert.AreEqual(3, col);
            _addressTranslator.ToColAndRow("AA2", out col, out row);
            Assert.AreEqual(27, col);
            _addressTranslator.ToColAndRow("BC1", out col, out row);
            Assert.AreEqual(55, col);
        }

        [TestMethod]
        public void ShouldTranslateLetterAddressUsingMaxRowsFromProviderLower()
        {
            int col, row;
            _addressTranslator.ToColAndRow("A", out col, out row);
            Assert.AreEqual(1, row);
        }

        [TestMethod]
        public void ShouldTranslateLetterAddressUsingMaxRowsFromProviderUpper()
        {
            int col, row;
            _addressTranslator.ToColAndRow("A", out col, out row, AddressTranslator.RangeCalculationBehaviour.LastPart);
            Assert.AreEqual(ExcelMaxRows, row);
        }
    }
}
