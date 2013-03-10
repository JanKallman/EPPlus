using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    public interface IOperator
    {
        Operators Operator { get; }

        CompileResult Apply(CompileResult left, CompileResult right);

        int Precedence { get; }
    }
}
