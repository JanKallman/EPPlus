using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using System.IO;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class DataValidationTests : ValidationTestBase
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

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfOperatorIsEqualAndFormula1IsEmpty()
        {
            var validations = _sheet.DataValidations.AddIntegerValidation("A1");
            validations.Operator = ExcelDataValidationOperator.equal;
            validations.Validate();
        }

        [TestMethod]
        public void DataValidations_ShouldSetShowErrorMessageFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", true, false);
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.IsTrue(validation.ShowErrorMessage ?? false);
        }

        [TestMethod]
        public void DataValidations_ShouldSetShowInputMessageFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", false, true);
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.IsTrue(validation.ShowInputMessage ?? false);
        }

        [TestMethod]
        public void DataValidations_ShouldSetPromptFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.AreEqual("Prompt", validation.Prompt);
        }

        [TestMethod]
        public void DataValidations_ShouldSetPromptTitleFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.AreEqual("PromptTitle", validation.PromptTitle);
        }

        [TestMethod]
        public void DataValidations_ShouldSetErrorFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.AreEqual("Error", validation.Error);
        }

        [TestMethod]
        public void DataValidations_ShouldSetErrorTitleFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.AreEqual("ErrorTitle", validation.ErrorTitle);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfOperatorIsBetweenAndFormula2IsEmpty()
        {
            var validation = _sheet.DataValidations.AddIntegerValidation("A1");
            validation.Formula.Value = 1;
            validation.Operator = ExcelDataValidationOperator.between;
            validation.Validate();
        }

        [TestMethod]
        public void DataValidations_ShouldAcceptOneItemOnly()
        {
            var validation = _sheet.DataValidations.AddListValidation("A1");
            validation.Formula.Values.Add("1");
            validation.Validate();
        }

        [TestMethod]
        public void ExcelDataValidation_ShouldReplaceLastPartInWholeColumnRangeWithMaxNumberOfRowsOneColumn()
        {
            // Act
            var validation = _sheet.DataValidations.AddIntegerValidation("A:A");

            // Assert
            Assert.AreEqual("A1:A" + ExcelPackage.MaxRows.ToString(), validation.Address.Address);
        }

        [TestMethod]
        public void ExcelDataValidation_ShouldReplaceLastPartInWholeColumnRangeWithMaxNumberOfRowsDifferentColumns()
        {
            // Act
            var validation = _sheet.DataValidations.AddIntegerValidation("A:B");

            // Assert
            Assert.AreEqual(string.Format("A1:B{0}", ExcelPackage.MaxRows), validation.Address.Address);
        }

    }
}
