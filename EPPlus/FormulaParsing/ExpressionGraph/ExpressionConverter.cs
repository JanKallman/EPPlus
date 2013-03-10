using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionConverter : IExpressionConverter
    {
        public StringExpression ToStringExpression(Expression expression)
        {
            var result = expression.Compile();
            var newExp = new StringExpression(result.Result.ToString());
            newExp.Operator = expression.Operator;
            return newExp;
        }

        public Expression FromCompileResult(CompileResult compileResult)
        {
            switch (compileResult.DataType)
            {
                case DataType.Integer:
                    return new IntegerExpression(compileResult.Result.ToString());
                case DataType.String:
                    return new StringExpression(compileResult.Result.ToString());
                case DataType.Decimal:
                    return new DecimalExpression(compileResult.Result.ToString());
                case DataType.Boolean:
                    return new BooleanExpression(compileResult.Result.ToString());
            }
            return null;
        }

        private static IExpressionConverter _instance;
        public static IExpressionConverter Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ExpressionConverter();
                }
                return _instance;
            }
        }
    }
}
