using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestClass]
    public class TimeFormulaTests : ValidationTestBase
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
        public void TimeFormula_ValueIsSetFromConstructorValidateHour()
        {
            // Arrange
            var time = new ExcelTime(0.675M);
            LoadXmlTestData("A1", "time", "0.675");

            // Act
            var formula = new ExcelDataValidationTime(_sheet, "A1", ExcelDataValidationType.Time, _dataValidationNode, _namespaceManager);
            
            // Assert
            Assert.AreEqual(time.Hour, formula.Formula.Value.Hour);
        }

        [TestMethod]
        public void TimeFormula_ValueIsSetFromConstructorValidateMinute()
        {
            // Arrange
            var time = new ExcelTime(0.395M);
            LoadXmlTestData("A1", "time", "0.395");

            // Act
            var formula = new ExcelDataValidationTime(_sheet, "A1", ExcelDataValidationType.Time, _dataValidationNode, _namespaceManager);

            // Assert
            Assert.AreEqual(time.Minute, formula.Formula.Value.Minute);
        }

        [TestMethod]
        public void TimeFormula_ValueIsSetFromConstructorValidateSecond()
        {
            // Arrange
            var time = new ExcelTime(0.812M);
            LoadXmlTestData("A1", "time", "0.812");

            // Act
            var formula = new ExcelDataValidationTime(_sheet, "A1", ExcelDataValidationType.Time, _dataValidationNode, _namespaceManager);

            // Assert
            Assert.AreEqual(time.Second.Value, formula.Formula.Value.Second.Value);
        }
    }
}
