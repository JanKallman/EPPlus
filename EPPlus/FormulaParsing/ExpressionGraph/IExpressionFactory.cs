using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public interface IExpressionFactory
    {
        Expression Create(Token token);
    }
}
