using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class ValidationCollectionTests : ValidationTestBase
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

        [TestMethod, ExpectedException(typeof(ArgumentNullException))]
        public void ExcelDataValidationCollection_AddDecimal_ShouldThrowWhenAddressIsNullOrEmpty()
        {
            // Act
            _sheet.DataValidations.AddDecimalValidation(string.Empty);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelDataValidationCollection_AddDecimal_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            // Act
            _sheet.DataValidations.AddDecimalValidation("A1");
            _sheet.DataValidations.AddDecimalValidation("A1");
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelDataValidationCollection_AddInteger_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            // Act
            _sheet.DataValidations.AddIntegerValidation("A1");
            _sheet.DataValidations.AddIntegerValidation("A1");
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelDataValidationCollection_AddTextLength_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            // Act
            _sheet.DataValidations.AddTextLengthValidation("A1");
            _sheet.DataValidations.AddTextLengthValidation("A1");
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelDataValidationCollection_AddDateTime_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            // Act
            _sheet.DataValidations.AddDateTimeValidation("A1");
            _sheet.DataValidations.AddDateTimeValidation("A1");
        }

        [TestMethod]
        public void ExcelDataValidationCollection_Index_ShouldReturnItemAtIndex()
        {
            // Arrange
            _sheet.DataValidations.AddDateTimeValidation("A1");
            _sheet.DataValidations.AddDateTimeValidation("A2");
            _sheet.DataValidations.AddDateTimeValidation("B1");

            // Act
            var result = _sheet.DataValidations[1];

            // Assert
            Assert.AreEqual("A2", result.Address.Address);
        }

        [TestMethod]
        public void ExcelDataValidationCollection_FindAll_ShouldReturnValidationInColumnAonly()
        {
            // Arrange
            _sheet.DataValidations.AddDateTimeValidation("A1");
            _sheet.DataValidations.AddDateTimeValidation("A2");
            _sheet.DataValidations.AddDateTimeValidation("B1");

            // Act
            var result = _sheet.DataValidations.FindAll(x => x.Address.Address.StartsWith("A"));

            // Assert
            Assert.AreEqual(2, result.Count());

        }

        [TestMethod]
        public void ExcelDataValidationCollection_Find_ShouldReturnFirstMatchOnly()
        {
            // Arrange
            _sheet.DataValidations.AddDateTimeValidation("A1");
            _sheet.DataValidations.AddDateTimeValidation("A2");

            // Act
            var result = _sheet.DataValidations.Find(x => x.Address.Address.StartsWith("A"));

            // Assert
            Assert.AreEqual("A1", result.Address.Address);

        }

        [TestMethod]
        public void ExcelDataValidationCollection_Clear_ShouldBeEmpty()
        {
            // Arrange
            _sheet.DataValidations.AddDateTimeValidation("A1");

            // Act
            _sheet.DataValidations.Clear();

            // Assert
            Assert.AreEqual(0, _sheet.DataValidations.Count);

        }

        [TestMethod]
        public void ExcelDataValidationCollection_RemoveAll_ShouldRemoveMatchingEntries()
        {
            // Arrange
            _sheet.DataValidations.AddIntegerValidation("A1");
            _sheet.DataValidations.AddIntegerValidation("A2");
            _sheet.DataValidations.AddIntegerValidation("B1");

            // Act
            _sheet.DataValidations.RemoveAll(x => x.Address.Address.StartsWith("B"));

            // Assert
            Assert.AreEqual(2, _sheet.DataValidations.Count);
        }
    }
}
