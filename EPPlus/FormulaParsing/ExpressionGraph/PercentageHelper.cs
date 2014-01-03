using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    internal static class PercentageHelper
    {
        internal static bool SupportsPercentage(DataType dataType)
        {
            switch (dataType)
            {
                case DataType.Decimal:
                case DataType.Integer:
                case DataType.Boolean:
                    return true;
                default:
                    return false;
            }
        }

        internal static double ApplyPercent(int numberOfPercentageSigns, double val)
        {
            double result = val;
            var nPercentSigns = numberOfPercentageSigns;
            while (nPercentSigns > 0)
            {
                result *= 0.01;
                nPercentSigns--;
            }
            return result;
        }
    }
}
