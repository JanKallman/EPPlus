/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 *  Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2011-01-08
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;

namespace EPPlusSamples
{
    /// <summary>
    /// This sample shows how to use data validation
    /// </summary>
    class Sample11
    {
        public static string RunSample11(DirectoryInfo outputDir)
        {
            //Create a Sample10 directory...
            if (!Directory.Exists(outputDir.FullName + @"\Sample11"))
            {
                outputDir.CreateSubdirectory("Sample11");
            }
            outputDir = new DirectoryInfo(outputDir + @"\Sample11");

            //create FileInfo object...
            FileInfo output = new FileInfo(outputDir.FullName + @"\Output.xlsx");
            if (output.Exists)
            {
                output.Delete();
                output = new FileInfo(outputDir.FullName + @"\Output.xlsx");
            }

            using (var package = new ExcelPackage(output))
            {
                AddIntegerValidation(package);
                AddListValidationFormula(package);
                AddListValidationValues(package);
                AddTimeValidation(package);
                AddDateTimeValidation(package);
                ReadExistingValidationsFromPackage(package);
                package.SaveAs(output);
            }
            return output.FullName;
        }

        /// <summary>
        /// Adds integer validation
        /// </summary>
        /// <param name="file"></param>
        private static void AddIntegerValidation(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("integer");
            // add a validation and set values
            var validation = sheet.DataValidations.AddIntegerValidation("A1:A2");
            // Alternatively:
            //var validation = sheet.Cells["A1:A2"].DataValidation.AddIntegerDataValidation();
            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validation.PromptTitle = "Enter a integer value here";
            validation.Prompt = "Value should be between 1 and 5";
            validation.ShowInputMessage = true;
            validation.ErrorTitle = "An invalid value was entered";
            validation.Error = "Value must be between 1 and 5";
            validation.ShowErrorMessage = true;
            validation.Operator = ExcelDataValidationOperator.between;
            validation.Formula.Value = 1;
            validation.Formula2.Value = 5;

            Console.WriteLine("Added sheet for integer validation");
        }

        /// <summary>
        /// Adds a list validation where the list source is a formula
        /// </summary>
        /// <param name="package"></param>
        private static void AddListValidationFormula(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("list formula");
            sheet.Cells["B1"].Style.Font.Bold = true;
            sheet.Cells["B1"].Value = "Source values";
            sheet.Cells["B2"].Value = 1;
            sheet.Cells["B3"].Value = 2;
            sheet.Cells["B4"].Value = 3;
            
            // add a validation and set values
            var validation = sheet.DataValidations.AddListValidation("A1");
            // Alternatively:
            // var validation = sheet.Cells["A1"].DataValidation.AddListDataValidation();
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
            validation.ErrorTitle = "An invalid value was entered";
            validation.Error = "Select a value from the list";
            validation.Formula.ExcelFormula = "B2:B4";

            Console.WriteLine("Added sheet for list validation with formula");
            
        }

        /// <summary>
        /// Adds a list validation where the selectable values are set
        /// </summary>
        /// <param name="package"></param>
        private static void AddListValidationValues(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("list values");

            // add a validation and set values
            var validation = sheet.DataValidations.AddListValidation("A1");
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
            validation.ErrorTitle = "An invalid value was entered";
            validation.Error = "Select a value from the list";
            for (var i = 1; i <= 5; i++)
            {
                validation.Formula.Values.Add(i.ToString());
            }
            Console.WriteLine("Added sheet for list validation with values");

        }

        /// <summary>
        /// Adds a time validation
        /// </summary>
        /// <param name="package"></param>
        private static void AddTimeValidation(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("time");
            // add a validation and set values
            var validation = sheet.DataValidations.AddTimeValidation("A1");
            // Alternatively:
            // var validation = sheet.Cells["A1"].DataValidation.AddTimeDataValidation();
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validation.ShowInputMessage = true;
            validation.PromptTitle = "Enter time in format HH:MM:SS";
            validation.Prompt = "Should be greater than 13:30:10";
            validation.Operator = ExcelDataValidationOperator.greaterThan;
            var time = validation.Formula.Value;
            time.Hour = 13;
            time.Minute = 30;
            time.Second = 10;
            Console.WriteLine("Added sheet for time validation");
        }

        private static void AddDateTimeValidation(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("datetime");
            // add a validation and set values
            var validation = sheet.DataValidations.AddDateTimeValidation("A1");
            // Alternatively:
            // var validation = sheet.Cells["A1"].DataValidation.AddDateTimeDataValidation();
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = ExcelDataValidationWarningStyle.stop;
            validation.Error = "Invalid date!";
            validation.ShowInputMessage = true;
            validation.Prompt = "Enter a date greater than todays date here";
            validation.Operator = ExcelDataValidationOperator.greaterThan;
            validation.Formula.Value = DateTime.Now.Date;
            Console.WriteLine("Added sheet for date time validation");

        }

        /// <summary>
        /// shows details about all existing validations in the entire workbook
        /// </summary>
        /// <param name="package"></param>
        private static void ReadExistingValidationsFromPackage(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("Package validations");
            // print headers
            sheet.Cells["A1:E1"].Style.Font.Bold = true;
            sheet.Cells["A1"].Value = "Type";
            sheet.Cells["B1"].Value = "Address";
            sheet.Cells["C1"].Value = "Operator";
            sheet.Cells["D1"].Value = "Formula1";
            sheet.Cells["E1"].Value = "Formula2";

            int row = 2;
            foreach (var otherSheet in package.Workbook.Worksheets)
            {
                if(otherSheet == sheet)
                {
                    continue;
                }
                foreach (var dataValidation in otherSheet.DataValidations)
                {
                    sheet.Cells["A" + row.ToString()].Value = dataValidation.ValidationType.Type.ToString();
                    sheet.Cells["B" + row.ToString()].Value = dataValidation.Address.Address;
                    if (dataValidation.AllowsOperator)
                    {
                        sheet.Cells["C" + row.ToString()].Value = ((IExcelDataValidationWithOperator)dataValidation).Operator.ToString();
                    }
                    // type casting is needed to get validationtype-specific values
                    switch(dataValidation.ValidationType.Type)
                    {
                        case eDataValidationType.Whole:
                            PrintWholeValidationDetails(sheet, (IExcelDataValidationInt)dataValidation, row);
                            break;
                        case eDataValidationType.List:
                            PrintListValidationDetails(sheet, (IExcelDataValidationList)dataValidation, row);
                            break;
                        case eDataValidationType.Time:
                            PrintTimeValidationDetails(sheet, (ExcelDataValidationTime)dataValidation, row);
                            break;
                        default:
                            // the rest of the types are not supported in this sample, but I hope you get the picture...
                            break;
                    }
                    row++;
                }
            }
        }

        private static void PrintWholeValidationDetails(ExcelWorksheet sheet, IExcelDataValidationInt wholeValidation, int row)
        {
            sheet.Cells["D" + row.ToString()].Value = wholeValidation.Formula.Value.HasValue ? wholeValidation.Formula.Value.Value.ToString() : wholeValidation.Formula.ExcelFormula;
            sheet.Cells["E" + row.ToString()].Value = wholeValidation.Formula2.Value.HasValue ? wholeValidation.Formula2.Value.Value.ToString() : wholeValidation.Formula2.ExcelFormula;
        }

        private static void PrintListValidationDetails(ExcelWorksheet sheet, IExcelDataValidationList listValidation, int row)
        {
            string value = string.Empty;
            // if formula is used - show it...
            if(!string.IsNullOrEmpty(listValidation.Formula.ExcelFormula))
            {
                value = listValidation.Formula.ExcelFormula;
            }
            else
            {
                // otherwise - show the values from the list collection
                var sb = new StringBuilder();
                foreach(var listValue in listValidation.Formula.Values)
                {
                    if(sb.Length > 0)
                    {
                        sb.Append(",");
                    }
                    sb.Append(listValue);
                }
                value = sb.ToString();
            }
            sheet.Cells["D" + row.ToString()].Value = value;
        }

        private static void PrintTimeValidationDetails(ExcelWorksheet sheet, ExcelDataValidationTime validation, int row)
        {
            var value1 = string.Empty;
            if(!string.IsNullOrEmpty(validation.Formula.ExcelFormula))
            {
                value1 = validation.Formula.ExcelFormula;
            }
            else
            {
                value1 = string.Format("{0}:{1}:{2}", validation.Formula.Value.Hour, validation.Formula.Value.Minute, validation.Formula.Value.Second ?? 0);
            }
            sheet.Cells["D" + row.ToString()].Value = value1;
        }
    }
}
