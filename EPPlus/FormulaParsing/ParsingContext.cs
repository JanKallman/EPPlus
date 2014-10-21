using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Parsing context
    /// </summary>
    public class ParsingContext : IParsingLifetimeEventHandler
    {
        private ParsingContext() { }

        /// <summary>
        /// The <see cref="FormulaParser"/> of the current context.
        /// </summary>
        public FormulaParser Parser { get; set; }

        /// <summary>
        /// The <see cref="ExcelDataProvider"/> is an abstraction on top of
        /// Excel, in this case EPPlus.
        /// </summary>
        public ExcelDataProvider ExcelDataProvider { get; set; }

        /// <summary>
        /// Utility for handling addresses
        /// </summary>
        public RangeAddressFactory RangeAddressFactory { get; set; }

        /// <summary>
        /// <see cref="INameValueProvider"/> of the current context
        /// </summary>
        public INameValueProvider NameValueProvider { get; set; }

        /// <summary>
        /// Configuration
        /// </summary>
        public ParsingConfiguration Configuration { get; set; }

        /// <summary>
        /// Scopes, a scope represents the parsing of a cell or a value.
        /// </summary>
        public ParsingScopes Scopes { get; private set; }

        /// <summary>
        /// Factory method.
        /// </summary>
        /// <returns></returns>
        public static ParsingContext Create()
        {
            var context = new ParsingContext();
            context.Configuration = ParsingConfiguration.Create();
            context.Scopes = new ParsingScopes(context);
            return context;
        }

        void IParsingLifetimeEventHandler.ParsingCompleted()
        {
            
        }
    }
}
