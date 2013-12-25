using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class CompileResultValidators
    {
        private readonly Dictionary<DataType, CompileResultValidator> _validators = new Dictionary<DataType, CompileResultValidator>(); 

        private CompileResultValidator CreateOrGet(DataType dataType)
        {
            if (_validators.ContainsKey(dataType))
            {
                return _validators[dataType];
            }
            if (dataType == DataType.Decimal)
            {
                return _validators[DataType.Decimal] = new DecimalCompileResultValidator();
            }
            return CompileResultValidator.Empty;
        }

        public CompileResultValidator GetValidator(DataType dataType)
        {
            return CreateOrGet(dataType);
        }
    }
}
