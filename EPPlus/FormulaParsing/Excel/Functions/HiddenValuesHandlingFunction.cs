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

        protected override IEnumerable<double> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if (!arguments.Any())
            {
                return Enumerable.Empty<double>();
            }
            if (IgnoreHiddenValues)
            {
                var nonHidden = arguments.Where(x => !x.ExcelStateFlagIsSet(ExcelCellState.HiddenCell));
                return base.ArgsToDoubleEnumerable(nonHidden, context);
            }
            return base.ArgsToDoubleEnumerable(arguments, context);
        }

        protected bool ShouldIgnore(ExcelDataProvider.ICellInfo c, ParsingContext context)
        {
            return CellStateHelper.ShouldIgnore(IgnoreHiddenValues, c, context);
        }
        protected bool ShouldIgnore(FunctionArgument arg)
        {
            if (IgnoreHiddenValues && arg.ExcelStateFlagIsSet(ExcelCellState.HiddenCell))
            {
                return true;
            }
            return false;
        }

    }
}
