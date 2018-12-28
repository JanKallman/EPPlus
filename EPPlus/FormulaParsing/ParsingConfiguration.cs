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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing
{
    public class ParsingConfiguration
    {
        public virtual ILexer Lexer { get; private set; }

        public IFormulaParserLogger Logger { get; private set; }

        public IExpressionGraphBuilder GraphBuilder { get; private set; }

        public IExpressionCompiler ExpressionCompiler{ get; private set; }

        public FunctionRepository FunctionRepository{ get; private set; }

        private ParsingConfiguration() 
        {
            FunctionRepository = FunctionRepository.Create();
        }

        internal static ParsingConfiguration Create()
        {
            return new ParsingConfiguration();
        }

        public ParsingConfiguration SetLexer(ILexer lexer)
        {
            Lexer = lexer;
            return this;
        }

        public ParsingConfiguration SetGraphBuilder(IExpressionGraphBuilder graphBuilder)
        {
            GraphBuilder = graphBuilder;
            return this;
        }

        public ParsingConfiguration SetExpresionCompiler(IExpressionCompiler expressionCompiler)
        {
            ExpressionCompiler = expressionCompiler;
            return this;
        }

        /// <summary>
        /// Attaches a logger, errors and log entries will be written to the logger during the parsing process.
        /// </summary>
        /// <param name="logger"></param>
        /// <returns></returns>
        public ParsingConfiguration AttachLogger(IFormulaParserLogger logger)
        {
            Require.That(logger).Named("logger").IsNotNull();
            Logger = logger;
            return this;
        }

        /// <summary>
        /// if a logger is attached it will be removed.
        /// </summary>
        /// <returns></returns>
        public ParsingConfiguration DetachLogger()
        {
            Logger = null;
            return this;
        }
    }
}
