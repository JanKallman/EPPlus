using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using System.IO;

namespace ExcelPackageTest.DataValidation
{
    [TestClass]
    public class DataValidationTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;

        [TestInitialize]
        public void Setup()
        {
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            _package = null;
            _sheet = null;
        }

        private string GetTestOutputPath(string fileName)
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
        }

        private void SaveTestOutput(string fileName)
        {
            var path = GetTestOutputPath(fileName);
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            _package.SaveAs(new FileInfo(path));
        }



        [TestMethod, Ignore]
        public void DataValidations_AddOneValidationOfTypeWhole()
        {
            var validation = _sheet.DataValidation.Create("A1", ExcelDataValidationType.Whole);
            validation.ShowErrorMessage = true;
            validation.ShowInputMessage = true;
            validation.Operator = ExcelDataValidationOperator.greaterThan;
            validation.Formula1 = "1";
            _sheet.DataValidation.Add(validation);

            _package.SaveAs(new FileInfo(GetTestOutputPath("AddOneValidationOfTypeWhole.xlsx")));
        }

        [TestMethod, Ignore]
        public void DataValidations_AddOneValidationOfTypeListOfTypeList()
        {
            var validation = _sheet.DataValidation.Create("A1", ExcelDataValidationType.List);
            validation.ShowErrorMessage = true;
            validation.ShowInputMessage = true;
            validation.Operator = ExcelDataValidationOperator.greaterThan;
            validation.Formula1 = "\"1, 2, 3, 4\"";
            _sheet.DataValidation.Add(validation);

            _package.SaveAs(new FileInfo(GetTestOutputPath("AddOneValidation.xlsx")));
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfOperatorIsEqualAndFormula1IsEmpty()
        {
            var validations = _sheet.DataValidation.Create("A1", ExcelDataValidationType.Whole);
            validations.Operator = ExcelDataValidationOperator.equal;
            _sheet.DataValidation.Add(validations);
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfOperatorIsBetweenAndFormula2IsEmpty()
        {
            var validation = _sheet.DataValidation.Create("A1", ExcelDataValidationType.Whole);
            validation.Operator = ExcelDataValidationOperator.between;
            _sheet.DataValidation.Add(validation);
        }

        [TestMethod, ExpectedException(typeof(FormatException))]
        public void DataValidations_ShouldThrowIfValidationTypeIsListAndFormula1DoesNotContainCommas()
        {
            var validation = _sheet.DataValidation.Create("A1", ExcelDataValidationType.List);
            validation.Formula1 = "1";
            _sheet.DataValidation.Add(validation);
        }
    }
}
