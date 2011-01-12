using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelPackageTest.DataValidation
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
    }
}
