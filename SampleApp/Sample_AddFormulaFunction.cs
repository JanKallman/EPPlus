using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace EPPlusSamples
{
    /// <summary>
    /// This sample shows how to add functions to the FormulaParser of EPPlus.
    /// 
    /// For further details on how to build functions, have a look in the EPPlus.FormulaParsing.Excel.Functions namespace
    /// </summary>
    class Sample_AddFormulaFunction
    {
        public static void RunSample_AddFormulaFunction()
        {
            using (var package = new ExcelPackage(new MemoryStream()))
            {
                // add your function module to the parser
                package.Workbook.FormulaParserManager.LoadFunctionModule(new MyFunctionModule());

                // Note that if you dont want to write a module, you can also
                // add new functions to the parser this way:
                // package.Workbook.FormulaParserManager.AddOrReplaceFunction("sum.addtwo", new SumAddTwo());-
                // package.Workbook.FormulaParserManager.AddOrReplaceFunction("seanconneryfy", new SeanConneryfy());
                

                //Override the buildin Text function to handle swedish date formatting strings. Excel has localized date format strings with is now supported by EPPlus.
                package.Workbook.FormulaParserManager.AddOrReplaceFunction("text", new TextSwedish());

                // add a worksheet with some dummy data
                var ws = package.Workbook.Worksheets.Add("Test");
                ws.Cells["A1"].Value = 1;
                ws.Cells["A2"].Value = 2;
                ws.Cells["P3"].Formula = "SUM(A1:A2)";
                ws.Cells["B1"].Value = "Hello";
                ws.Cells["C1"].Value = new DateTime(2013,12,31);
                ws.Cells["C2"].Formula="Text(C1,\"åååå-MM-dd\")";   //Swedish formatting
                // use the added "sum.addtwo" function
                ws.Cells["A4"].Formula = "SUM.ADDTWO(A1:A2,P3)";
                // use the other function "seanconneryfy"
                ws.Cells["B2"].Formula = "seanconneryfy(B1)";

                // calculate
                ws.Calculate();
                                
                // show result
                Console.WriteLine("sum.addtwo(A1:A2,P3) evaluated to {0}", ws.Cells["A4"].Value);
                Console.WriteLine("seanconneryfy(B1) evaluated to {0}", ws.Cells["B2"].Value);
            }
        }
    }

    class MyFunctionModule : IFunctionModule
    {
        public MyFunctionModule()
        {
            Functions = new Dictionary<string, ExcelFunction>()
                {
                    {"sum.addtwo", new SumAddTwo()},
                    {"seanconneryfy", new SeanConneryfy()}
                };    
        }

        public IDictionary<string, ExcelFunction> Functions { get; private set; }
    }

    /// <summary>
    /// A really unnecessary function. Adds two to all numbers in the supplied range and calculates the sum.
    /// </summary>
    class SumAddTwo : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            // Sanity check, will set excel VALUE error if min length is not met
            ValidateArguments(arguments, 1);
            
            // Helper method that converts function arguments to an enumerable of doubles
            var numbers = ArgsToDoubleEnumerable(arguments, context);
            
            // Do the work
            var result = 0d;
            numbers.ToList().ForEach(x => result += (x + 2));

            // return the result
            return CreateResult(result, DataType.Decimal);
        }
    }
    /// <summary>
    /// This function handles Swedish formatting strings.
    /// </summary>
    class TextSwedish : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            // Sanity check, will set excel VALUE error if min length is not met
            ValidateArguments(arguments, 2);

            //Replace swedish year format with invariant for parameter 2.
            var format = arguments.ElementAt(1).Value.ToString().Replace("åååå", "yyyy");   
            var newArgs = new List<FunctionArgument> { arguments.ElementAt(0) };
            newArgs.Add(new FunctionArgument(format));

            //Use the build-in Text function.
            var func = new Text();
            return func.Execute(newArgs, context);
        }
    }

    /// <summary>
    /// An even more unnecessary function, inspired by the Sean Connery keyboard;) Will add 'sh' at the end of the supplied string.
    /// </summary>
    class SeanConneryfy : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            // Sanity check, will set excel VALUE error if min length is not met
            ValidateArguments(arguments, 1);
            // Get the first arg
            var input = ArgToString(arguments, 0);

            // return the result
            return CreateResult(input + "sh", DataType.String);
        }
    }
}
