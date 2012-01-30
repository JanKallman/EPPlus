using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;
using System.Collections;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestClass]
    public class ListFormulaTests : ValidationTestBase
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
        public void ListFormula_FormulaValueIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "list", "\"1,2\"");
            // Act
            var validation = new ExcelDataValidationList(_sheet, "A1", ExcelDataValidationType.List, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.AreEqual(2, validation.Formula.Values.Count);
        }

        [TestMethod]
        public void ListFormula_FormulaValueIsSetFromXmlNodeInConstructorOrderIsCorrect()
        {
            // Arrange
            LoadXmlTestData("A1", "list", "\"1,2\"");
            // Act
            var validation = new ExcelDataValidationList(_sheet, "A1", ExcelDataValidationType.List, _dataValidationNode, _namespaceManager);
            // Assert
            CollectionAssert.AreEquivalent(new List<string>{ "1", "2"}, (ICollection)validation.Formula.Values);
        }

        [TestMethod]
        public void ListFormula_FormulasExcelFormulaIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "list", "A1");
            // Act
            var validation = new ExcelDataValidationList(_sheet, "A1", ExcelDataValidationType.List, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.AreEqual("A1", validation.Formula.ExcelFormula);
        }
    }
}
