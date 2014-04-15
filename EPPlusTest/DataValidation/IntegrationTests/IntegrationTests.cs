using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.DataValidation;
using System.IO;
using OfficeOpenXml;

namespace EPPlusTest.DataValidation.IntegrationTests
{
    /// <summary>
    /// Remove the Ignore attributes from the testmethods if you want to run any of these tests
    /// </summary>
    [TestClass]
    public class IntegrationTests : ValidationTestBase
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

        [TestMethod, Ignore]
        public void DataValidations_AddOneValidationOfTypeWhole()
        {
            _sheet.Cells["B1"].Value = 2;
            var validation = _sheet.DataValidations.AddIntegerValidation("A1");
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validation.ErrorTitle = "Invalid value was entered";
            validation.PromptTitle = "Enter value here";
            validation.Operator = ExcelDataValidationOperator.greaterThan;
            //validation.Value.Value = 3;
            validation.Formula.ExcelFormula = "B1";

            _package.SaveAs(new FileInfo(GetTestOutputPath("AddOneValidationOfTypeWhole.xlsx")));
        }
        [TestMethod, Ignore]
        public void DataValidations_AddOneValidationOfTypeDecimal()
        {
            var validation = _sheet.DataValidations.AddDecimalValidation("A1");
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validation.ErrorTitle = "Invalid value was entered";
            validation.Error = "Value must be greater than 1.4";
            validation.PromptTitle = "Enter value here";
            validation.Prompt = "Enter a value that is greater than 1.4";
            validation.ShowInputMessage = true;
            validation.Operator = ExcelDataValidationOperator.greaterThan;
            validation.Formula.Value = 1.4;

            _package.SaveAs(new FileInfo(GetTestOutputPath("AddOneValidationOfTypeDecimal.xlsx")));
        }

        [TestMethod]
        public void DataValidations_AddOneValidationOfTypeListOfTypeList()
        {
            var validation = _sheet.DataValidations.AddListValidation("A:A");
            validation.ShowErrorMessage = true;
            validation.ShowInputMessage = true;
            validation.Formula.Values.Add("1");
            validation.Formula.Values.Add("2");
            validation.Formula.Values.Add("3");
            validation.Validate();

            _package.SaveAs(new FileInfo(GetTestOutputPath("AddOneValidationOfTypeList.xlsx")));
        }

        [TestMethod]
        public void DataValidations_AddOneValidationOfTypeListOfTypeTime()
        {
            var validation = _sheet.DataValidations.AddTimeValidation("A1");
            validation.ShowErrorMessage = true;
            validation.ShowInputMessage = true;
            validation.Formula.Value.Hour = 14;
            validation.Formula.Value.Minute = 30;
            validation.Operator = ExcelDataValidationOperator.greaterThan;
            validation.Prompt = "Enter a time greater than 14:30";
            validation.Error = "Invalid time was entered";
            validation.Validate();

            _package.SaveAs(new FileInfo(GetTestOutputPath("AddOneValidationOfTypeTime.xlsx")));
        }

        [TestMethod, Ignore]
        public void DataValidations_ReadExistingWorkbookWithDataValidations()
        {
            using (var package = new ExcelPackage(new FileInfo(GetTestOutputPath("DVTest.xlsx"))))
            {
                Assert.AreEqual(3, package.Workbook.Worksheets[1].DataValidations.Count);
            }
        }

        //[TestMethod]
        //public void Bug_OpenOffice()
        //{
        //    var xlsPath = GetTestOutputPath("OpenOffice.xlsx");
        //    var newFile = new FileInfo(xlsPath);
        //    if( newFile.Exists)
        //    {
        //        newFile.Delete();
        //        //ensures we create a new workbook
        //        newFile = new FileInfo(xlsPath);

        //    }
        //    using(var package = new ExcelPackage(newFile))
        //    {
        //        // add a new worksheet to the empty workbook
        //        var worksheet = package.Workbook.Worksheets.Add("Inventory");
        //        // Add the headers
        //        worksheet.Cells[1, 1].Value = "ID";
        //        worksheet.Cells[1, 2].Value = "Product";
        ////        worksheet.Cells[1, 3].Value = "Quantity";
        //        worksheet.Cells[1, 4].Value = "Price";

        //        worksheet.Cells[1, 5].Value = "Value";

        //        worksheet.Column(1).Width = 15;
        //        worksheet.Column(2).Width = 12;
        //        worksheet.Column(3).Width = 12;

        //        worksheet.View.PageLayoutView = true;

        //        package.Workbook.Properties.Title = "Invertory";
        //        package.Workbook.Properties.Author = "Jan Källman";
        //        package.Save();
        //    }

    }
}
