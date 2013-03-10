using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionFactory : IExpressionFactory
    {
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ParsingContext _parsingContext;

        public ExpressionFactory(ExcelDataProvider excelDataProvider, ParsingContext context)
        {
            _excelDataProvider = excelDataProvider;
            _parsingContext = context;
        }


        public Expression Create(Token token)
        {
            switch (token.TokenType)
            {
                case TokenType.Integer:
                    return new IntegerExpression(token.Value);
                case TokenType.String:
                    return new StringExpression(token.Value);
                case TokenType.Decimal:
                    return new DecimalExpression(token.Value);
                case TokenType.Boolean:
                    return new BooleanExpression(token.Value);
                case TokenType.ExcelAddress:
                    return new ExcelAddressExpression(token.Value, _excelDataProvider, _parsingContext);
                case TokenType.NameValue:
                    return new NamedValueExpression(token.Value, _parsingContext);
                default:
                    return new StringExpression(token.Value);
            }
        }
    }
}
