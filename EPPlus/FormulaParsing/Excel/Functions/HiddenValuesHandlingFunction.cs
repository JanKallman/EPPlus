using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public abstract class HiddenValuesHandlingFunction : ExcelFunction
    {
        public bool IgnoreHiddenValues
        {
            get;
            set;
        }

        protected override IEnumerable<double> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments)
        {
            if (!arguments.Any())
            {
                return Enumerable.Empty<double>();
            }
            if (IgnoreHiddenValues)
            {
                var nonHidden = arguments.Where(x => !x.ExcelStateFlagIsSet(ExcelCellState.HiddenCell));
                return base.ArgsToDoubleEnumerable(nonHidden);
            }
            return base.ArgsToDoubleEnumerable(arguments);
        }

        protected bool ShouldIgnore(ExcelDataProvider.ICellInfo c, ParsingContext context)
        {
            return (IgnoreHiddenValues  && c.IsHiddenRow) || (context.Scopes.Current.IsSubtotal && IsSubTotal(c));
        }
        protected bool ShouldIgnore(FunctionArgument arg)
        {
            if (IgnoreHiddenValues && arg.ExcelStateFlagIsSet(ExcelCellState.HiddenCell))
            {
                return true;
            }
            return false;
        }
        protected bool IsSubTotal(ExcelDataProvider.ICellInfo c)
        {
            var tokens = c.Tokens;
            if (tokens == null) return false;
            foreach (var token in c.Tokens)
            {
                if (token.TokenType == LexicalAnalysis.TokenType.Function && token.Value.Equals("SUBTOTAL", StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

    }
}
