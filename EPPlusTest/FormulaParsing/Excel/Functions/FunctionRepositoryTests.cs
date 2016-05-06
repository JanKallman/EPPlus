using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class FunctionRepositoryTests
    {
        #region LoadModule Tests
        [TestMethod]
        public void LoadModulePopulatesFunctionsAndCustomCompilers()
        {
            var functionRepository = FunctionRepository.Create();
            Assert.IsFalse(functionRepository.IsFunctionName(MyFunction.Name));
            Assert.IsFalse(functionRepository.CustomCompilers.ContainsKey(typeof(MyFunction)));
            functionRepository.LoadModule(new TestFunctionModule());
            Assert.IsTrue(functionRepository.IsFunctionName(MyFunction.Name));
            Assert.IsTrue(functionRepository.CustomCompilers.ContainsKey(typeof(MyFunction)));
        }
        #endregion

        #region Nested Classes
        public class TestFunctionModule : FunctionsModule
        {
            public TestFunctionModule()
            {
                var myFunction = new MyFunction();
                var customCompiler = new MyFunctionCompiler(myFunction);
                base.Functions.Add(MyFunction.Name, myFunction);
                base.CustomCompilers.Add(typeof(MyFunction), customCompiler);
            }
        }

        public class MyFunction : ExcelFunction
        {
            public const string Name = "MyFunction";
            public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
            {
                throw new NotImplementedException();
            }
        }

        public class MyFunctionCompiler : FunctionCompiler
        {
            public MyFunctionCompiler(MyFunction function) : base(function) { }
            public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
            {
                throw new NotImplementedException();
            }
        }
        #endregion
    }
}
