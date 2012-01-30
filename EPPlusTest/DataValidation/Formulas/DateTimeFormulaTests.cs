using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestClass]
    public class DateTimeFormulaTests : ValidationTestBase
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
        public void DateTimeFormula_FormulaValueIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            var date = DateTime.Parse("2011-01-08");
            var dateAsString = date.ToOADate().ToString(_cultureInfo);
            LoadXmlTestData("A1", "decimal", dateAsString);
            // Act
            var validation = new ExcelDataValidationDateTime(_sheet, "A1", ExcelDataValidationType.Decimal, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.AreEqual(date, validation.Formula.Value);
        }

        [TestMethod]
        public void DateTimeFormula_FormulasFormulaIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            var date = DateTime.Parse("2011-01-08");
            LoadXmlTestData("A1", "decimal", "A1");

            // Act
            var validation = new ExcelDataValidationDateTime(_sheet, "A1", ExcelDataValidationType.Decimal, _dataValidationNode, _namespaceManager);

            // Assert
            Assert.AreEqual("A1", validation.Formula.ExcelFormula);
        }
    }
}
