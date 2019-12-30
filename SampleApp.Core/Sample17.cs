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
 * See https://github.com/JanKallman/EPPlus for details.
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
 * Jan Källman		Added		10-SEP-2009
 *******************************************************************************/

using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using OfficeOpenXml;
using System.IO;
using System.Data.SqlClient;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing;
using System.Linq;

namespace EPPlusSamples
{
    class EvaluateFunction : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var args = arguments.ToArray();
            // renderer.customFunctionEvaluated = true;

           double? hasDefaultVal = null;
            if (args.Length > 1 && args[1].Value is double)
            {
                hasDefaultVal = (double)args[1].Value;
            }
            return new CompileResult(args[0].Value.ToString(), DataType.String);
        }
    }
    class Sample17
    {
        /// <summary>
        /// This sample creates a new workbook from a template file containing a chart and populates it with Exchangrates from 
        /// the Adventureworks database and set the three series on the chart.
        /// </summary>
        /// <param name="connectionString">Connectionstring to the Adventureworks db</param>
        /// <param name="template">the template</param>
        /// <param name="outputdir">output dir</param>
        /// <returns></returns>
        public static string RunSample17(FileInfo template)
        {
            using (var stream = File.Open(template.FullName, FileMode.Open))
            using (ExcelPackage p = new ExcelPackage(stream))
            {
                p.Workbook.FormulaParserManager.AddOrReplaceFunction("e", new EvaluateFunction());
                //Set up the headers
                ExcelWorksheet ws = p.Workbook.Worksheets[0];
                for (int i=1; i<= ws.Dimension.Rows; i++)
                {
                    for (int j=1; j<=ws.Dimension.Columns; j++)
                    {
                        var cell = ws.Cells[i, j];
                        if (cell.Merge)
                            continue;
                        var val = ws.Cells[i, j].Value;
                        var formual = ws.Cells[i, j].Formula;
                        if (!string.IsNullOrEmpty(formual))
                        {
                            var result = p.Workbook.FormulaParserManager.Parse(formual);
                            var val2 = ws.Calculate(formual);
                            ws.Cells[i, j].Formula = string.Empty;
                            ws.Cells[i, j].Value = val2;
                            Console.WriteLine($"{i},{j}: {val} and f={formual} and eval={val2}");
                        }
                    }
                }
                p.Workbook.Worksheets[0].Select();
                
                foreach (var a in ws.MergedCells)
                {
                    var cell = ws.Cells[a];
                    var formual = cell.Formula;
                    if (!string.IsNullOrEmpty(formual))
                    {
                        var tokens = p.Workbook.FormulaParser.Lexer.Tokenize(formual);
                            //var result = p.Workbook.FormulaParserManager.Parse();
                        var val2 = ws.Calculate(formual);
                        cell.Formula = string.Empty;
                        cell.Value = val2;
                        Console.WriteLine($"{cell.Value} and f={formual} and eval={val2}");
                    }
                }
                Byte[] bin = p.GetAsByteArray();

                FileInfo file = Utils.GetFileInfo("sample17.xlsx");
                File.WriteAllBytes(file.FullName, bin);
                return file.FullName;
            }
        }
    }
}