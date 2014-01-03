using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public abstract class PercentHandlingExpression : AtomicExpression
    {
        protected PercentHandlingExpression(string expression)
            : base(expression)
        {

        }

        protected int NumberOfPercentSigns { get; private set; }

        public override void SetPercentage()
        {
            NumberOfPercentSigns++;
        }

        protected double ApplyPercent(double val)
        {
            return PercentageHelper.ApplyPercent(NumberOfPercentSigns, val);
        }
    }
}
