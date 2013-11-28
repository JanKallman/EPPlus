using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Second : TimeBaseFunction
    {
        public Second()
            : base()
        {

        }
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var firstArg = arguments.ElementAt(0).Value.ToString();
            if (arguments.Count() == 1 && TimeStringParser.CanParse(firstArg))
            {
                var result = TimeStringParser.Parse(firstArg);
                return CreateResult(GetSecondFromSerialNumber(result), DataType.Integer);
            }
            ValidateAndInitSerialNumber(arguments);
            return CreateResult(GetSecondFromSerialNumber(SerialNumber), DataType.Integer);
        }

        private int GetSecondFromSerialNumber(double serialNumber)
        {
            return (int)System.Math.Round(GetSecond(serialNumber), 0);
        }
    }
}
