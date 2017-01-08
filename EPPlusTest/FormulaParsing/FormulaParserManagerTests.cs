using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class FormulaParserManagerTests
    {
        #region test classes

        private class MyFunction : ExcelFunction
        {
            public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
            {
                throw new NotImplementedException();
            }
        }

        private class MyModule : IFunctionModule
        {
            public MyModule()
            {
                Functions = new Dictionary<string, ExcelFunction>();
                Functions.Add("MyFunction", new MyFunction());

                CustomCompilers = new Dictionary<Type, FunctionCompiler>();
            }
            public IDictionary<string, ExcelFunction> Functions { get; }
            public IDictionary<Type, FunctionCompiler> CustomCompilers { get; }
        }
        #endregion

        [TestMethod]
        public void FunctionsShouldBeCopied()
        {
            using (var package1 = new ExcelPackage())
            {
                package1.Workbook.FormulaParserManager.LoadFunctionModule(new MyModule());
                using (var package2 = new ExcelPackage())
                {
                    var origNumberOfFuncs = package2.Workbook.FormulaParserManager.GetImplementedFunctionNames().Count();

                    // replace functions including the custom functions from package 1
                    package2.Workbook.FormulaParserManager.CopyFunctionsFrom(package1.Workbook);

                    // Assertions: number of functions are increased with 1, and the list of function names contains the custom function.
                    Assert.AreEqual(origNumberOfFuncs + 1, package2.Workbook.FormulaParserManager.GetImplementedFunctionNames().Count());
                    Assert.IsTrue(package2.Workbook.FormulaParserManager.GetImplementedFunctionNames().Contains("myfunction"));
                }
            }
        }
    }
}
