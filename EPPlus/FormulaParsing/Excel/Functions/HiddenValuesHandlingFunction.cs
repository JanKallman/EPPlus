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
            if (IgnoreHiddenValues)
            {
                var nonHidden = arguments.Where(x => !x.ExcelStateFlagIsSet(ExcelCellState.HiddenCell));
                return base.ArgsToDoubleEnumerable(nonHidden);
            }
            return base.ArgsToDoubleEnumerable(arguments);
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
