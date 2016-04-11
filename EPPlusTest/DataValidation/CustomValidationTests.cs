using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class CustomValidationTests : ValidationTestBase
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
        public void CustomValidation_FormulaIsSet()
        {
            // Act
            var validation = _sheet.DataValidations.AddCustomValidation("A1");

            // Assert
            Assert.IsNotNull(validation.Formula);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void CustomValidation_ShouldThrowExceptionIfFormulaIsTooLong()
        {
            // Arrange
            var sb = new StringBuilder();
            for (var x = 0; x < 257; x++) sb.Append("x");
            
            // Act
            var validation = _sheet.DataValidations.AddCustomValidation("A1");
            validation.Formula.ExcelFormula = sb.ToString();
        }
    }
}
