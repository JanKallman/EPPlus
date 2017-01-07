using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation.Formulas;
using System.Xml;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestClass]
    public class DecimalFormulaTests : ValidationTestBase
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
        public void DecimalFormula_FormulaValueIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "decimal", "1.3");
            // Act
            var validation = new ExcelDataValidationDecimal(_sheet, "A1", ExcelDataValidationType.Decimal, _dataValidationNode, _namespaceManager);
            Assert.AreEqual(1.3D, validation.Formula.Value);
        }

        [TestMethod]
        public void DecimalFormula_FormulasFormulaIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "decimal", "A1");

            // Act
            var validation = new ExcelDataValidationDecimal(_sheet, "A1", ExcelDataValidationType.Decimal, _dataValidationNode, _namespaceManager);

            // Assert
            Assert.AreEqual("A1", validation.Formula.ExcelFormula);
        }
    }
}
