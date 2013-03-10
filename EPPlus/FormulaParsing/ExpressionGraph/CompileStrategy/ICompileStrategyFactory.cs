using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy
{
    public interface ICompileStrategyFactory
    {
        CompileStrategy Create(Expression expression);
    }
}
