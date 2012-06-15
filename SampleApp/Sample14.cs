/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 *
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
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Eyal Seagull				Added							2012-04-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.ConditionalFormatting;

namespace EPPlusSamples
{
  class Sample14
  {
    /// <summary>
    /// Sample 14 - Conditional formatting example
    /// </summary>
    public static string RunSample14(DirectoryInfo outputDir)
    {
      FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample14.xlsx");

      if (newFile.Exists)
      {
        newFile.Delete();  // ensures we create a new workbook
        newFile = new FileInfo(outputDir.FullName + @"\sample14.xlsx");
      }

      using (ExcelPackage package = new ExcelPackage(newFile))
      {
        // add a new worksheet to the empty workbook
        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Conditional Formatting");

        // Create 4 columns of samples data
        for (int col = 1; col < 5; col++)
        {
          // Add the headers
          worksheet.Cells[1, col].Value = "Sample " + col;

          for (int row = 2; row < 21; row++)
          {
            // Add some items...
            worksheet.Cells[row, col].Value = row;
          }
        }

        // -------------------------------------------------------------------
        // TwoColorScale Conditional Formatting example
        // -------------------------------------------------------------------
        ExcelAddress cfAddress1 = new ExcelAddress("A2:A10");
        var cfRule1 = worksheet.ConditionalFormatting.AddTwoColorScale(cfAddress1);

        // Now, lets change some properties:
        cfRule1.LowValue.Type = eExcelConditionalFormattingValueObjectType.Num;
        cfRule1.LowValue.Value = 4;
        cfRule1.LowValue.Color = ColorTranslator.FromHtml("#FFFFEB84");
        cfRule1.HighValue.Type = eExcelConditionalFormattingValueObjectType.Formula;
        cfRule1.HighValue.Formula = "IF($G$1=\"A</x:&'cfRule>\",1,5)";
        cfRule1.StopIfTrue = true;
        cfRule1.Style.Font.Bold = true;

        // But others you can't (readonly)
        // cfRule1.Type = eExcelConditionalFormattingRuleType.ThreeColorScale;

        // -------------------------------------------------------------------
        // ThreeColorScale Conditional Formatting example
        // -------------------------------------------------------------------
        ExcelAddress cfAddress2 = new ExcelAddress(2, 2, 10, 2);  //="B2:B10"
        var cfRule2 = worksheet.ConditionalFormatting.AddThreeColorScale(cfAddress2);

        // Changing some properties again
        cfRule2.Priority = 1;
        cfRule2.MiddleValue.Type = eExcelConditionalFormattingValueObjectType.Percentile;
        cfRule2.MiddleValue.Value = 30;
        cfRule2.StopIfTrue = true;

        // You can access a rule by its Priority
        var cfRule2Priority = cfRule2.Priority;
        var cfRule2_1 = worksheet.ConditionalFormatting.RulesByPriority(cfRule2Priority);

        // And you can even change the rule's Address
        cfRule2_1.Address = new ExcelAddress("Z1:Z3");

        // -------------------------------------------------------------------
        // Adding another ThreeColorScale in a different way (observe that we are
        // pointing to the same range as the first rule we entered. Excel allows it to
        // happen and group the rules in one <conditionalFormatting> node)
        // -------------------------------------------------------------------
        var cfRule3 = worksheet.Cells[cfAddress1.Address].ConditionalFormatting.AddThreeColorScale();
        cfRule3.LowValue.Color = Color.LemonChiffon;

        // -------------------------------------------------------------------
        // Change the rules priorities to change their execution order
        // -------------------------------------------------------------------
        cfRule3.Priority = 1;
        cfRule1.Priority = 2;
        cfRule2.Priority = 3;

        // -------------------------------------------------------------------
        // Create an Above Average rule
        // -------------------------------------------------------------------
        var cfRule5 = worksheet.ConditionalFormatting.AddAboveAverage(
          new ExcelAddress("B11:B20"));
        cfRule5.Style.Font.Bold = true;
        cfRule5.Style.Font.Color.Color = Color.Red;
        cfRule5.Style.Font.Strike = true;

        // -------------------------------------------------------------------
        // Create an Above Or Equal Average rule
        // -------------------------------------------------------------------
        var cfRule6 = worksheet.ConditionalFormatting.AddAboveOrEqualAverage(
          new ExcelAddress("B11:B20"));

        // -------------------------------------------------------------------
        // Create a Below Average rule
        // -------------------------------------------------------------------
        var cfRule7 = worksheet.ConditionalFormatting.AddBelowAverage(
          new ExcelAddress("B11:B20"));

        // -------------------------------------------------------------------
        // Create a Below Or Equal Average rule
        // -------------------------------------------------------------------
        var cfRule8 = worksheet.ConditionalFormatting.AddBelowOrEqualAverage(
          new ExcelAddress("B11:B20"));

        // -------------------------------------------------------------------
        // Create a Above StdDev rule
        // -------------------------------------------------------------------
        var cfRule9 = worksheet.ConditionalFormatting.AddAboveStdDev(
          new ExcelAddress("B11:B20"));
          cfRule9.StdDev = 0;

        // -------------------------------------------------------------------
        // Create a Below StdDev rule
        // -------------------------------------------------------------------
        var cfRule10 = worksheet.ConditionalFormatting.AddBelowStdDev(
          new ExcelAddress("B11:B20"));

        cfRule10.StdDev = 2;

        // -------------------------------------------------------------------
        // Create a Bottom rule
        // -------------------------------------------------------------------
        var cfRule11 = worksheet.ConditionalFormatting.AddBottom(
          new ExcelAddress("B11:B20"));

        cfRule11.Rank = 4;

        // -------------------------------------------------------------------
        // Create a Bottom Percent rule
        // -------------------------------------------------------------------
        var cfRule12 = worksheet.ConditionalFormatting.AddBottomPercent(
          new ExcelAddress("B11:B20"));

        cfRule12.Rank = 15;

        // -------------------------------------------------------------------
        // Create a Top rule
        // -------------------------------------------------------------------
        var cfRule13 = worksheet.ConditionalFormatting.AddTop(
          new ExcelAddress("B11:B20"));

        // -------------------------------------------------------------------
        // Create a Top Percent rule
        // -------------------------------------------------------------------
        var cfRule14 = worksheet.ConditionalFormatting.AddTopPercent(
          new ExcelAddress("B11:B20"));
        
        cfRule14.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        cfRule14.Style.Border.Left.Color.Theme = 3;
        cfRule14.Style.Border.Bottom.Style = ExcelBorderStyle.DashDot;
        cfRule14.Style.Border.Bottom.Color.Index=8;
        cfRule14.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        cfRule14.Style.Border.Right.Color.Color=Color.Blue;
        cfRule14.Style.Border.Top.Style = ExcelBorderStyle.Hair;
        cfRule14.Style.Border.Top.Color.Auto=true;

        // -------------------------------------------------------------------
        // Create a Last 7 Days rule
        // -------------------------------------------------------------------
        ExcelAddress timePeriodAddress = new ExcelAddress("D21:G34 C11:C20");
        var cfRule15 = worksheet.ConditionalFormatting.AddLast7Days(
          timePeriodAddress);

        cfRule15.Style.Fill.PatternType = ExcelFillStyle.LightTrellis;
        cfRule15.Style.Fill.PatternColor.Color = Color.BurlyWood;
        cfRule15.Style.Fill.BackgroundColor.Color = Color.LightCyan;

        // -------------------------------------------------------------------
        // Create a Last Month rule
        // -------------------------------------------------------------------
        var cfRule16 = worksheet.ConditionalFormatting.AddLastMonth(
          timePeriodAddress);

        cfRule16.Style.NumberFormat.Format = "YYYY";
        // -------------------------------------------------------------------
        // Create a Last Week rule
        // -------------------------------------------------------------------
        var cfRule17 = worksheet.ConditionalFormatting.AddLastWeek(
          timePeriodAddress);
        cfRule17.Style.NumberFormat.Format = "YYYY";

        // -------------------------------------------------------------------
        // Create a Next Month rule
        // -------------------------------------------------------------------
        var cfRule18 = worksheet.ConditionalFormatting.AddNextMonth(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a Next Week rule
        // -------------------------------------------------------------------
        var cfRule19 = worksheet.ConditionalFormatting.AddNextWeek(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a This Month rule
        // -------------------------------------------------------------------
        var cfRule20 = worksheet.ConditionalFormatting.AddThisMonth(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a This Week rule
        // -------------------------------------------------------------------
        var cfRule21 = worksheet.ConditionalFormatting.AddThisWeek(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a Today rule
        // -------------------------------------------------------------------
        var cfRule22 = worksheet.ConditionalFormatting.AddToday(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a Tomorrow rule
        // -------------------------------------------------------------------
        var cfRule23 = worksheet.ConditionalFormatting.AddTomorrow(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a Yesterday rule
        // -------------------------------------------------------------------
        var cfRule24 = worksheet.ConditionalFormatting.AddYesterday(
          timePeriodAddress);

        // -------------------------------------------------------------------
        // Create a BeginsWith rule
        // -------------------------------------------------------------------
        ExcelAddress cellIsAddress = new ExcelAddress("E11:E20");
        var cfRule25 = worksheet.ConditionalFormatting.AddBeginsWith(
          cellIsAddress);

        cfRule25.Text = "SearchMe";

        // -------------------------------------------------------------------
        // Create a Between rule
        // -------------------------------------------------------------------
        var cfRule26 = worksheet.ConditionalFormatting.AddBetween(
          cellIsAddress);

        cfRule26.Formula = "IF(E11>5,10,20)";
        cfRule26.Formula2 = "IF(E11>5,30,50)";

        // -------------------------------------------------------------------
        // Create a ContainsBlanks rule
        // -------------------------------------------------------------------
        var cfRule27 = worksheet.ConditionalFormatting.AddContainsBlanks(
          cellIsAddress);

        // -------------------------------------------------------------------
        // Create a ContainsErrors rule
        // -------------------------------------------------------------------
        var cfRule28 = worksheet.ConditionalFormatting.AddContainsErrors(
          cellIsAddress);

        // -------------------------------------------------------------------
        // Create a ContainsText rule
        // -------------------------------------------------------------------
        var cfRule29 = worksheet.ConditionalFormatting.AddContainsText(
          cellIsAddress);

        cfRule29.Text = "Me";

        // -------------------------------------------------------------------
        // Create a DuplicateValues rule
        // -------------------------------------------------------------------
        var cfRule30 = worksheet.ConditionalFormatting.AddDuplicateValues(
          cellIsAddress);

        // -------------------------------------------------------------------
        // Create an EndsWith rule
        // -------------------------------------------------------------------
        var cfRule31 = worksheet.ConditionalFormatting.AddEndsWith(
          cellIsAddress);

        cfRule31.Text = "EndText";

        // -------------------------------------------------------------------
        // Create an Equal rule
        // -------------------------------------------------------------------
        var cfRule32 = worksheet.ConditionalFormatting.AddEqual(
          cellIsAddress);

        cfRule32.Formula = "6";

        // -------------------------------------------------------------------
        // Create an Expression rule
        // -------------------------------------------------------------------
        var cfRule33 = worksheet.ConditionalFormatting.AddExpression(
          cellIsAddress);

        cfRule33.Formula = "E11=E12";

        // -------------------------------------------------------------------
        // Create a GreaterThan rule
        // -------------------------------------------------------------------
        var cfRule34 = worksheet.ConditionalFormatting.AddGreaterThan(
          cellIsAddress);

        cfRule34.Formula = "SE(E11<10,10,65)";

        // -------------------------------------------------------------------
        // Create a GreaterThanOrEqual rule
        // -------------------------------------------------------------------
        var cfRule35 = worksheet.ConditionalFormatting.AddGreaterThanOrEqual(
          cellIsAddress);

        cfRule35.Formula = "35";

        // -------------------------------------------------------------------
        // Create a LessThan rule
        // -------------------------------------------------------------------
        var cfRule36 = worksheet.ConditionalFormatting.AddLessThan(
          cellIsAddress);

        cfRule36.Formula = "36";

        // -------------------------------------------------------------------
        // Create a LessThanOrEqual rule
        // -------------------------------------------------------------------
        var cfRule37 = worksheet.ConditionalFormatting.AddLessThanOrEqual(
          cellIsAddress);

        cfRule37.Formula = "37";

        // -------------------------------------------------------------------
        // Create a NotBetween rule
        // -------------------------------------------------------------------
        var cfRule38 = worksheet.ConditionalFormatting.AddNotBetween(
          cellIsAddress);

        cfRule38.Formula = "333";
        cfRule38.Formula2 = "999";

        // -------------------------------------------------------------------
        // Create a NotContainsBlanks rule
        // -------------------------------------------------------------------
        var cfRule39 = worksheet.ConditionalFormatting.AddNotContainsBlanks(
          cellIsAddress);

        // -------------------------------------------------------------------
        // Create a NotContainsErrors rule
        // -------------------------------------------------------------------
        var cfRule40 = worksheet.ConditionalFormatting.AddNotContainsErrors(
          cellIsAddress);

        // -------------------------------------------------------------------
        // Create a NotContainsText rule
        // -------------------------------------------------------------------
        var cfRule41 = worksheet.ConditionalFormatting.AddNotContainsText(
          cellIsAddress);

        cfRule41.Text = "NotMe";

        // -------------------------------------------------------------------
        // Create an NotEqual rule
        // -------------------------------------------------------------------
        var cfRule42 = worksheet.ConditionalFormatting.AddNotEqual(
          cellIsAddress);

        cfRule42.Formula = "14";

        // -----------------------------------------------------------
        // Removing Conditional Formatting rules
        // -----------------------------------------------------------
        // Remove one Rule by its object
        //worksheet.ConditionalFormatting.Remove(cfRule1);

        // Remove one Rule by index
        //worksheet.ConditionalFormatting.RemoveAt(1);

        // Remove one Rule by its Priority
        //worksheet.ConditionalFormatting.RemoveByPriority(2);

        // Remove all the Rules
        //worksheet.ConditionalFormatting.RemoveAll();

        // set some document properties
        package.Workbook.Properties.Title = "Conditional Formatting";
        package.Workbook.Properties.Author = "Eyal Seagull";
        package.Workbook.Properties.Comments = "This sample demonstrates how to add Conditional Formatting to an Excel 2007 worksheet using EPPlus";

        // set some custom property values
        package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Eyal Seagull");
        package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");

        // save our new workbook and we are done!
        package.Save();
      }

      return newFile.FullName;
    }
  }
}
