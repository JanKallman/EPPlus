using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace EPPlusSamples
{
    class Sample_FormulaCalc
    {
        public static void RunSampleFormulaCalc()
        {
            using (var package = new ExcelPackage(new MemoryStream()))
            {
                var ws1 = package.Workbook.Worksheets.Add("ws1");
                // Add some values to sum
                ws1.Cells["A1"].Formula = "(2*2)/2";
                ws1.Cells["A2"].Value = 4;
                ws1.Cells["A3"].Value = 6;
                ws1.Cells["A4"].Formula = "SUM(A1:A3)";
                
                // calculate all formulas on  the worksheet
                ws1.Calculate();

                // Print the calculated value
                Console.WriteLine("SUM(A1:A3) evaluated to {0}", ws1.Cells["A4"].Value);

                // Add another worksheet
                var ws2 = package.Workbook.Worksheets.Add("ws2");
                ws2.Cells["A1"].Value = 3;
                ws2.Cells["A2"].Formula = "SUM(A1,ws1!A4)";

                // calculate all formulas in the entire workbook
                package.Workbook.Calculate();

                // Print the calculated value
                Console.WriteLine("SUM(A1,ws1!A4) evaluated to {0}", ws2.Cells["A2"].Value);

                // calculate a range
                ws1.Cells["B1"].Formula = "IF(TODAY()<DATE(2013;6;1);\"BEFORE\" &\" FIRST\";CONCATENATE(\"FIRST\";\" OF\";\" JUNE 2013 OR LATER\"))";
                ws1.Cells["B1"].Calculate();
                
                // Print the calculated value
                Console.WriteLine("IF(TODAY()<DATE(2014;6;1);\"BEFORE\" &\" FIRST\";CONCATENATE(\"FIRST\";\" OF\";\" JUNE OR LATER\")) evaluated to {0}", ws1.Cells["B1"].Value);

                // evaluate a formula string
                const string formula = "(2+4)*ws1!A2";
                var result = package.Workbook.FormulaParserManager.Parse(formula);

                // print the calculated value
                Console.WriteLine("(2+4)*ws1!A2 evaluated to {0}", result);
                ws1.Calculate("(2+4)*A2");
                
            }
        }
    }
}
