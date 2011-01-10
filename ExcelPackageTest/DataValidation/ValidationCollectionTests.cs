using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;

namespace ExcelPackageTest.DataValidation
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
            _sheet.DataValidation.AddDecimalValidation(string.Empty);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelDataValidationCollection_AddDecimal_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            // Act
            _sheet.DataValidation.AddDecimalValidation("A1");
            _sheet.DataValidation.AddDecimalValidation("A1");
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelDataValidationCollection_AddInteger_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            // Act
            _sheet.DataValidation.AddWholeValidation("A1");
            _sheet.DataValidation.AddWholeValidation("A1");
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelDataValidationCollection_AddTextLength_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            // Act
            _sheet.DataValidation.AddTextLengthValidation("A1");
            _sheet.DataValidation.AddTextLengthValidation("A1");
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void ExcelDataValidationCollection_AddDateTime_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            // Act
            _sheet.DataValidation.AddDateTimeValidation("A1");
            _sheet.DataValidation.AddDateTimeValidation("A1");
        }

        [TestMethod]
        public void ExcelDataValidationCollection_Index_ShouldReturnItemAtIndex()
        {
            // Arrange
            _sheet.DataValidation.AddDateTimeValidation("A1");
            _sheet.DataValidation.AddDateTimeValidation("A2");
            _sheet.DataValidation.AddDateTimeValidation("B1");

            // Act
            var result = _sheet.DataValidation[1];

            // Assert
            Assert.AreEqual("A2", result.Address.Address);
        }

        [TestMethod]
        public void ExcelDataValidationCollection_FindAll_ShouldReturnValidationInColumnAonly()
        {
            // Arrange
            _sheet.DataValidation.AddDateTimeValidation("A1");
            _sheet.DataValidation.AddDateTimeValidation("A2");
            _sheet.DataValidation.AddDateTimeValidation("B1");

            // Act
            var result = _sheet.DataValidation.FindAll(x => x.Address.Address.StartsWith("A"));

            // Assert
            Assert.AreEqual(2, result.Count());

        }

        [TestMethod]
        public void ExcelDataValidationCollection_Find_ShouldReturnFirstMatchOnly()
        {
            // Arrange
            _sheet.DataValidation.AddDateTimeValidation("A1");
            _sheet.DataValidation.AddDateTimeValidation("A2");

            // Act
            var result = _sheet.DataValidation.Find(x => x.Address.Address.StartsWith("A"));

            // Assert
            Assert.AreEqual("A1", result.Address.Address);

        }

        [TestMethod]
        public void ExcelDataValidationCollection_Clear_ShouldBeEmpty()
        {
            // Arrange
            _sheet.DataValidation.AddDateTimeValidation("A1");

            // Act
            _sheet.DataValidation.Clear();

            // Assert
            Assert.AreEqual(0, _sheet.DataValidation.Count);

        }

        [TestMethod]
        public void ExcelDataValidationCollection_RemoveAll_ShouldRemoveMatchingEntries()
        {
            // Arrange
            _sheet.DataValidation.AddWholeValidation("A1");
            _sheet.DataValidation.AddWholeValidation("A2");
            _sheet.DataValidation.AddWholeValidation("B1");

            // Act
            _sheet.DataValidation.RemoveAll(x => x.Address.Address.StartsWith("B"));

            // Assert
            Assert.AreEqual(2, _sheet.DataValidation.Count);
        }
    }
}
