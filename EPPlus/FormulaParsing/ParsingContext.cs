/* Copyright (C) 2011  Jan Källman
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied.
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author Change                      Date
 *******************************************************************************
 * Mats Alm Added		                2016-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Logging;

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

        public ExcelAddressCache AddressCache { get; private set; }

        /// <summary>
        /// Returns true if a <see cref="IFormulaParserLogger"/> is attached to the parser.
        /// </summary>
        public bool Debug
        {
            get { return Configuration.Logger != null; }
        }

        /// <summary>
        /// Factory method.
        /// </summary>
        /// <returns></returns>
        public static ParsingContext Create()
        {
            var context = new ParsingContext();
            context.Configuration = ParsingConfiguration.Create();
            context.Scopes = new ParsingScopes(context);
            context.AddressCache = new ExcelAddressCache();
            return context;
        }

        void IParsingLifetimeEventHandler.ParsingCompleted()
        {
            AddressCache.Clear();
        }
    }
}
