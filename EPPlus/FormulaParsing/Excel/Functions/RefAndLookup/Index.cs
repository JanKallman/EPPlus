using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public class Index : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var arg1 = arguments.ElementAt(0);
            var args = arg1.Value as IEnumerable<FunctionArgument>;
            if (args != null)
            {
                var index = ArgToInt(arguments, 1);
                if (index > args.Count())
                {
                    throw new ExcelErrorValueException(eErrorType.Ref);
                }
                var candidate = args.ElementAt(index - 1);
                if (!IsNumber(candidate.Value))
                {
                    throw new ExcelErrorValueException(eErrorType.Value);
                }
                return CreateResult(ConvertUtil.GetValueDouble(candidate.Value), DataType.Decimal);
            }
            if (arg1.IsExcelRange)
            {
                var index = ArgToInt(arguments, 1);
                if(index < arg1.ValueAsRangeInfo.Count())
                {
                    ThrowExcelErrorValueException(eErrorType.Ref);
                }
                var candidate = arg1.ValueAsRangeInfo.ElementAt(index - 1);
                if (!IsNumber(candidate.Value))
                {
                    throw new ExcelErrorValueException(eErrorType.Value);
                }
                return CreateResult(ConvertUtil.GetValueDouble(candidate.Value), DataType.Decimal);
            }
            throw new NotImplementedException();
        }
    }
}
