using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EPPlusTest.DataValidation
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
            _sheet.Cells["A1:A2"].DataValidation.AddIntegerDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidations.Count);
        }

        [TestMethod]
        public void RangeBase_AddIntegerValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddIntegerDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidations[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AddDecimalValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddDecimalDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidations.Count);
        }

        [TestMethod]
        public void RangeBase_AddDecimalValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddDecimalDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidations[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AddTextLengthValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddTextLengthDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidations.Count);
        }

        [TestMethod]
        public void RangeBase_AddTextLengthValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddTextLengthDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidations[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AddDateTimeValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddDateTimeDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidations.Count);
        }

        [TestMethod]
        public void RangeBase_AddDateTimeValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddDateTimeDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidations[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AddListValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddListDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidations.Count);
        }

        [TestMethod]
        public void RangeBase_AddListValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddListDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidations[0].Address.Address);
        }

        [TestMethod]
        public void RangeBase_AdTimeValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddTimeDataValidation();

            // Assert
            Assert.AreEqual(1, _sheet.DataValidations.Count);
        }

        [TestMethod]
        public void RangeBase_AddTimeValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddTimeDataValidation();

            // Assert
            Assert.AreEqual("A1:A2", _sheet.DataValidations[0].Address.Address);
        }
    }
}
