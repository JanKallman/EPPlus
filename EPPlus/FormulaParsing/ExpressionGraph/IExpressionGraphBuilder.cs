using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public interface IExpressionGraphBuilder
    {
        ExpressionGraph Build(IEnumerable<Token> tokens);
    }
}
