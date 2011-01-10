using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelPackageTest.DataValidation
{
    [TestClass]
    public class RangeBaseTests : ValidationTestBase
    {
        [TestInitialize]
        public void Setup()
        {
            SetupTestData();
        }

        [TestCleanup]
        public void Cleanup()
        {
            CleanupTestData();
        }

        [TestMethod]
        public void RangeBase_AddIntegerValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].AddIntegerDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidation.Count);
        }

        [TestMethod]
        public void RangeBase_AddIntegerValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].AddIntegerDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidation[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AddDecimalValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].AddDecimalDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidation.Count);
        }

        [TestMethod]
        public void RangeBase_AddDecimalValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].AddDecimalDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidation[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AddTextLengthValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].AddTextLengthDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidation.Count);
        }

        [TestMethod]
        public void RangeBase_AddTextLengthValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].AddTextLengthDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidation[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AddDateTimeValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].AddDateTimeDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidation.Count);
        }

        [TestMethod]
        public void RangeBase_AddDateTimeValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].AddDateTimeDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidation[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AddListValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].AddListDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidation.Count);
        }

        [TestMethod]
        public void RangeBase_AddListValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].AddListDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidation[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AdTimeValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].AddTimeDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidation.Count);
        }

        [TestMethod]
        public void RangeBase_AddTimeValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].AddTimeDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidation[0].Address.Address);
        }
    }
}
