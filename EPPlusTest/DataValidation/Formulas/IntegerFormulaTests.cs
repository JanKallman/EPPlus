using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestClass]
    public class IntegerFormulaTests : ValidationTestBase
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
            _dataValidationNode = null;
        }

        [TestMethod]
        public void IntegerFormula_FormulaValueIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "decimal", "1");
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            Assert.AreEqual(1, validation.Formula.Value);
        }

        [TestMethod]
        public void IntegerFormula_FormulasFormulaIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "decimal", "A1");

            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);

            // Assert
            Assert.AreEqual("A1", validation.Formula.ExcelFormula);
        }
    }
}
