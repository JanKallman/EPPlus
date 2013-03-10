using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExcelAddressExpression : Expression
    {
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ParsingContext _parsingContext;
        private readonly RangeAddressFactory _rangeAddressFactory;

        public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
            : this(expression, excelDataProvider, parsingContext, new RangeAddressFactory(excelDataProvider))
        {

        }

        public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext, RangeAddressFactory rangeAddressFactory)
            : base(expression)
        {
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            Require.That(parsingContext).Named("parsingContext").IsNotNull();
            Require.That(rangeAddressFactory).Named("rangeAddressFactory").IsNotNull();
            _excelDataProvider = excelDataProvider;
            _parsingContext = parsingContext;
            _rangeAddressFactory = rangeAddressFactory;
        }

        public override bool IsGroupedExpression
        {
            get { return false; }
        }

        public override CompileResult Compile()
        {
            if (ParentIsLookupFunction)
            {
                return new CompileResult(ExpressionString, DataType.ExcelAddress);
            }
            else
            {
                return CompileRangeValues();
            }
        }

        private CompileResult CompileRangeValues()
        {
            var result = _excelDataProvider.GetRangeValues(ExpressionString);
            if (result == null || result.Count() == 0)
            {
                return null;
            }
            var rangeValueList = HandleRangeValues(result);
            if (rangeValueList.Count > 1)
            {
                return new CompileResult(rangeValueList, DataType.Enumerable);
            }
            else
            {
                var factory = new CompileResultFactory();
                return factory.Create(rangeValueList.First());
            }
        }

        private List<object> HandleRangeValues(IEnumerable<ExcelCell> result)
        {
            var rangeValueList = new List<object>();
            for (int x = 0; x < result.Count(); x++)
            {
                var dataItem = result.ElementAt(x);
                if (dataItem.Value != null)
                {
                    rangeValueList.Add(dataItem.Value);
                }
                else if (!string.IsNullOrEmpty(dataItem.Formula))
                {
                    var address = _rangeAddressFactory.Create(dataItem.ColIndex - 1, dataItem.RowIndex);
                    rangeValueList.Add(_parsingContext.Parser.Parse(dataItem.Formula, address));
                }
            }
            return rangeValueList;
        }
    }
}
