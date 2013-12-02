using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class ArgumentParsers
    {
        private static object _syncRoot = new object();
        private readonly Dictionary<DataType, ArgumentParser> _parsers = new Dictionary<DataType, ArgumentParser>();
        private readonly ArgumentParserFactory _parserFactory;

        public ArgumentParsers()
            : this(new ArgumentParserFactory())
        {

        }

        public ArgumentParsers(ArgumentParserFactory factory)
        {
            Require.That(factory).Named("argumentParserfactory").IsNotNull();
            _parserFactory = factory;
        }

        public ArgumentParser GetParser(DataType dataType)
        {
            if (!_parsers.ContainsKey(dataType))
            {
                lock (_syncRoot)
                {
                    if (!_parsers.ContainsKey(dataType))
                    {
                        _parsers.Add(dataType, _parserFactory.CreateArgumentParser(dataType));
                    }
                }
            }
            return _parsers[dataType];
        }
    }
}
