using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class ArgumentCollectionUtil
    {
        private readonly DoubleEnumerableArgConverter _doubleEnumerableArgConverter;
        private readonly ObjectEnumerableArgConverter _objectEnumerableArgConverter;

        public ArgumentCollectionUtil()
            : this(new DoubleEnumerableArgConverter(), new ObjectEnumerableArgConverter())
        {

        }

        public ArgumentCollectionUtil(
            DoubleEnumerableArgConverter doubleEnumerableArgConverter, 
            ObjectEnumerableArgConverter objectEnumerableArgConverter)
        {
            _doubleEnumerableArgConverter = doubleEnumerableArgConverter;
            _objectEnumerableArgConverter = objectEnumerableArgConverter;
        }

        public virtual IEnumerable<double> ArgsToDoubleEnumerable(bool ignoreHidden, IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return _doubleEnumerableArgConverter.ConvertArgs(ignoreHidden, arguments, context);
        }

        public virtual IEnumerable<object> ArgsToObjectEnumerable(bool ignoreHidden,
                                                                  IEnumerable<FunctionArgument> arguments,
                                                                  ParsingContext context)
        {
            return _objectEnumerableArgConverter.ConvertArgs(ignoreHidden, arguments, context);
        }

        public virtual double CalculateCollection(IEnumerable<FunctionArgument> collection, double result, Func<FunctionArgument, double, double> action)
        {
            foreach (var item in collection)
            {
                if (item.Value is IEnumerable<FunctionArgument>)
                {
                    result = CalculateCollection((IEnumerable<FunctionArgument>)item.Value, result, action);
                }
                else
                {
                    result = action(item, result);
                }
            }
            return result;
        }
    }
}
