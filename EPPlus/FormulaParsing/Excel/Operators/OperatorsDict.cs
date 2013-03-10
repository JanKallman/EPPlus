using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    public class OperatorsDict : Dictionary<string, IOperator>
    {
        public OperatorsDict()
        {
            Add("+", Operator.Plus);
            Add("-", Operator.Minus);
            Add("*", Operator.Multiply);
            Add("/", Operator.Divide);
            Add("^", Operator.Exp);
            Add("=", Operator.Eq);
            Add(">", Operator.GreaterThan);
            Add(">=", Operator.GreaterThanOrEqual);
            Add("<", Operator.LessThan);
            Add("<=", Operator.LessThanOrEqual);
            Add("<>", Operator.NotEqualsTo);
            Add("&", Operator.Concat);
            Add("mod", Operator.Modulus);
        }

        public static IDictionary<string, IOperator> _instance;

        public static IDictionary<string, IOperator> Instance
        {
            get 
            {
                if (_instance == null)
                {
                    _instance = new OperatorsDict();
                }
                return _instance;
            }
        }
    }
}
