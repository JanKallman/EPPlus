using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class DecimaDataValidationTests : ValidationTestBase
    {
        private IExcelDataValidationDecimal _validation;

        [TestInitialize]
        public void Setup()
        {
            SetupTestData();
            _validation = _package.Workbook.Worksheets[1].DataValidations.AddDecimalValidation("A1");
        }

        [TestCleanup]
        public void Cleanup()
        {
            CleanupTestData();
            _validation = null;
        }

        [TestMethod]
        public void DecimalDataValidation_Formula1IsSet()
        {
            Assert.IsNotNull(_validation.Formula);
        }

        [TestMethod]
        public void DecimalDataValidation_Formula2IsSet()
        {
            Assert.IsNotNull(_validation.Formula2);
        }
    }
}
