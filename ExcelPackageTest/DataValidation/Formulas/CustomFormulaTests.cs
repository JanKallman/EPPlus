using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestClass]
    public class CustomFormulaTests : ValidationTestBase
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
        public void CustomFormula_FormulasFormulaIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "decimal", "A1");

            // Act
            var validation = new ExcelDataValidationCustom(_sheet, "A1", ExcelDataValidationType.Custom, _dataValidationNode, _namespaceManager);

            // Assert
            Assert.AreEqual("A1", validation.Formula.ExcelFormula);
        }
    }
}
