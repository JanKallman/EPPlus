/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

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
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.FormulaParsing.Utilities;
namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Provides access to various functionality regarding 
    /// excel formula evaluation.
    /// </summary>
    public class FormulaParserManager
    {
        private readonly FormulaParser _parser;

        internal FormulaParserManager(FormulaParser parser)
        {
            Require.That(parser).Named("parser").IsNotNull();
            _parser = parser;
        }

        /// <summary>
        /// Loads a module containing custom functions to the formula parser. By using
        /// this method you can add your own implementations of Excel functions, by
        /// implementing a <see cref="IFunctionModule"/>.
        /// </summary>
        /// <param name="module">A <see cref="IFunctionModule"/> containing <see cref="ExcelFunction"/>s.</param>
        public void LoadFunctionModule(IFunctionModule module)
        {
            _parser.Configure(x => x.FunctionRepository.LoadModule(module));
        }

        /// <summary>
        /// If the supplied <paramref name="functionName"/> does not exist, the supplied
        /// <paramref name="functionImpl"/> implementation will be added to the formula parser.
        /// If it exists, the existing function will be replaced by the supplied <paramref name="functionImpl">function implementation</paramref>
        /// </summary>
        /// <param name="functionName"></param>
        /// <param name="functionImpl"></param>
        public void AddOrReplaceFunction(string functionName, ExcelFunction functionImpl)
        {
            _parser.Configure(x => x.FunctionRepository.AddOrReplaceFunction(functionName, functionImpl));
        }

        /// <summary>
        /// Returns an enumeration of all functions implemented, both the built in functions
        /// and functions added using the LoadFunctionModule method of this class.
        /// </summary>
        /// <returns>Function names in lower case</returns>
        public IEnumerable<string> GetImplementedFunctionNames()
        {
            var fnList = _parser.FunctionNames.ToList();
            fnList.Sort((x, y) => String.Compare(x, y, System.StringComparison.Ordinal));
            return fnList;
        }

        /// <summary>
        /// Parses the supplied <paramref name="formula"/> and returns the result.
        /// </summary>
        /// <param name="formula"></param>
        /// <returns></returns>
        public object Parse(string formula)
        {
            return _parser.Parse(formula);
        }

        /// <summary>
        /// Attaches a logger to the <see cref="FormulaParser"/>.
        /// </summary>
        /// <param name="logger">An instance of <see cref="IFormulaParserLogger"/></param>
        /// <see cref="OfficeOpenXml.FormulaParsing.Logging.LoggerFactory"/>
        public void AttachLogger(IFormulaParserLogger logger)
        {
            _parser.Configure(c => c.AttachLogger(logger));
        }

        /// <summary>
        /// Attaches a logger to the formula parser that produces output to the supplied logfile.
        /// </summary>
        /// <param name="logfile"></param>
        public void AttachLogger(FileInfo logfile)
        {
            _parser.Configure(c => c.AttachLogger(LoggerFactory.CreateTextFileLogger(logfile)));
        }
        /// <summary>
        /// Detaches any attached logger from the formula parser.
        /// </summary>
        public void DetachLogger()
        {
            _parser.Configure(c => c.DetachLogger());
        }
    }
}
