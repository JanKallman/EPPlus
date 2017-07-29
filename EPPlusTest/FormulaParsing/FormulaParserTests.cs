using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;
using ExGraph = OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class FormulaParserTests
    {
        private FormulaParser _parser;

        [TestInitialize]
        public void Setup()
        {
            var provider = A.Fake<ExcelDataProvider>();
            _parser = new FormulaParser(provider);

        }

        [TestCleanup]
        public void Cleanup()
        {

        }

        [TestMethod]
        public void ParserShouldCallLexer()
        {
            var lexer = A.Fake<ILexer>();
            A.CallTo(() => lexer.Tokenize("ABC")).Returns(Enumerable.Empty<Token>());
            _parser.Configure(x => x.SetLexer(lexer));

            _parser.Parse("ABC");

            A.CallTo(() => lexer.Tokenize("ABC")).MustHaveHappened();
        }

        [TestMethod]
        public void ParserShouldCallGraphBuilder()
        {
            var lexer = A.Fake<ILexer>();
            var tokens = new List<Token>();
            A.CallTo(() => lexer.Tokenize("ABC")).Returns(tokens);
            var graphBuilder = A.Fake<IExpressionGraphBuilder>();
            A.CallTo(() => graphBuilder.Build(tokens)).Returns(new ExGraph());

            _parser.Configure(config =>
                {
                    config
                        .SetLexer(lexer)
                        .SetGraphBuilder(graphBuilder);
                });

            _parser.Parse("ABC");

            A.CallTo(() => graphBuilder.Build(tokens)).MustHaveHappened();
        }

        [TestMethod]
        public void ParserShouldCallCompiler()
        {
            var lexer = A.Fake<ILexer>();
            var tokens = new List<Token>();
            A.CallTo(() => lexer.Tokenize("ABC")).Returns(tokens);
            var expectedGraph = new ExGraph();
            expectedGraph.Add(new StringExpression("asdf"));
            var graphBuilder = A.Fake<IExpressionGraphBuilder>();
            A.CallTo(() => graphBuilder.Build(tokens)).Returns(expectedGraph);
            var compiler = A.Fake<IExpressionCompiler>();
            A.CallTo(() => compiler.Compile(expectedGraph.Expressions)).Returns(new CompileResult(0, DataType.Integer));

            _parser.Configure(config =>
            {
                config
                    .SetLexer(lexer)
                    .SetGraphBuilder(graphBuilder)
                    .SetExpresionCompiler(compiler);
            });

            _parser.Parse("ABC");

            A.CallTo(() => compiler.Compile(expectedGraph.Expressions)).MustHaveHappened();
        }

        [TestMethod]
        public void ParseAtShouldCallExcelDataProvider()
        {
            var excelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => excelDataProvider.GetRangeFormula(string.Empty, 1, 1)).Returns("Sum(1,2)");
            var parser = new FormulaParser(excelDataProvider);
            var result = parser.ParseAt("A1");
            Assert.AreEqual(3d, result);
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void ParseAtShouldThrowIfAddressIsNull()
        {
            _parser.ParseAt(null);
        }
    }
}
