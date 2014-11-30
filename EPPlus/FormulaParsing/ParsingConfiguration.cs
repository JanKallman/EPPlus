using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing
{
    public class ParsingConfiguration
    {
        public virtual ILexer Lexer { get; private set; }

        public virtual IdProvider IdProvider { get; private set; }

        public IExpressionGraphBuilder GraphBuilder { get; private set; }

        public IExpressionCompiler ExpressionCompiler{ get; private set; }

        public FunctionRepository FunctionRepository{ get; private set; }

        private ParsingConfiguration() 
        {
            FunctionRepository = FunctionRepository.Create();
        }

        internal static ParsingConfiguration Create()
        {
            return new ParsingConfiguration();
        }

        public ParsingConfiguration SetIdProvider(IdProvider idProvider)
        {
            IdProvider = idProvider;
            return this;
        }

        public ParsingConfiguration SetLexer(ILexer lexer)
        {
            Lexer = lexer;
            return this;
        }

        public ParsingConfiguration SetGraphBuilder(IExpressionGraphBuilder graphBuilder)
        {
            GraphBuilder = graphBuilder;
            return this;
        }

        public ParsingConfiguration SetExpresionCompiler(IExpressionCompiler expressionCompiler)
        {
            ExpressionCompiler = expressionCompiler;
            return this;
        }
    }
}
