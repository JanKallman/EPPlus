using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Time : TimeBaseFunction
    {
        public Time()
            : base()
        {

        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var firstArg = arguments.ElementAt(0).Value.ToString();
            if(arguments.Count() == 1 && TimeStringParser.CanParse(firstArg))
            {
                var result = TimeStringParser.Parse(firstArg);
                return new CompileResult(result, DataType.Time);
            }
            ValidateArguments(arguments, 3);
            var hour = ArgToInt(arguments, 0);
            var min = ArgToInt(arguments, 1);
            var sec = ArgToInt(arguments, 2);

            ThrowArgumentExceptionIf(() => sec < 0 || sec > 59, "Invalid second: " + sec);
            ThrowArgumentExceptionIf(() => min < 0 || min > 59, "Invalid minute: " + min);
            ThrowArgumentExceptionIf(() => min < 0 || hour > 23, "Invalid hour: " + hour);


            var secondsOfThisTime = (double)(hour * 60 * 60 + min * 60 + sec);
            return CreateResult(GetTimeSerialNumber(secondsOfThisTime), DataType.Time);
        }
    }
}
