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



        [TestMethod]
        public void DataValidations_AddOneValidationOfTypeWhole()
        {
            var validation = _sheet.DataValidation.AddDecimalValidation("A1");
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validation.ErrorTitle = "Invalid value was entered";
            validation.Error = "Value must be greater than 1.4";
            validation.PromptTitle = "Enter value here";
            validation.Prompt = "Enter a value that is greater than 1.4";
            validation.ShowInputMessage = true;
            validation.Operator = ExcelDataValidationOperator.greaterThan;
            validation.Value = (decimal)1.4;

            _package.SaveAs(new FileInfo(GetTestOutputPath("AddOneValidationOfTypeWhole.xlsx")));
        }

        [TestMethod]
        public void DataValidations_AddOneValidationOfTypeListOfTypeList()
        {
            var validation = _sheet.DataValidation.AddListValidation("A1");
            validation.ShowErrorMessage = true;
            validation.ShowInputMessage = true;
            validation.Values.Add("1");
            validation.Values.Add("2");
            validation.Values.Add("3");
            validation.Validate();

            _package.SaveAs(new FileInfo(GetTestOutputPath("AddOneValidationOfTypeList.xlsx")));
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfOperatorIsEqualAndFormula1IsEmpty()
        {
            var validations = _sheet.DataValidation.AddWholeValidation("A1");
            validations.Operator = ExcelDataValidationOperator.equal;
            validations.Validate();
        }

        [TestMethod, ExpectedException(typeof(InvalidOperationException))]
        public void DataValidations_ShouldThrowIfOperatorIsBetweenAndFormula2IsEmpty()
        {
            var validation = _sheet.DataValidation.AddWholeValidation("A1");
            validation.Operator = ExcelDataValidationOperator.between;
            validation.Validate();
        }

        [TestMethod, ExpectedException(typeof(FormatException))]
        public void DataValidations_ShouldThrowIfValidationTypeIsListAndFormula1DoesNotContainCommas()
        {
            var validation = _sheet.DataValidation.AddListValidation("A1");
            validation.Values.Add("1");
            validation.Validate();
        }
    }
}
