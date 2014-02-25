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
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;
using System.Diagnostics;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing
{
    public class FormulaParser
    {
        private readonly ParsingContext _parsingContext;
        private readonly ExcelDataProvider _excelDataProvider;

        public FormulaParser(ExcelDataProvider excelDataProvider)
            : this(excelDataProvider, ParsingContext.Create())
        {
           
        }

        public FormulaParser(ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
        {
            parsingContext.Parser = this;
            parsingContext.ExcelDataProvider = excelDataProvider;
            parsingContext.NameValueProvider = new EpplusNameValueProvider(excelDataProvider);
            parsingContext.RangeAddressFactory = new RangeAddressFactory(excelDataProvider);
            _parsingContext = parsingContext;
            _excelDataProvider = excelDataProvider;
            Configure(configuration =>
            {
                configuration
                    .SetLexer(new Lexer(_parsingContext.Configuration.FunctionRepository, _parsingContext.NameValueProvider))
                    .SetGraphBuilder(new ExpressionGraphBuilder(excelDataProvider, _parsingContext))
                    .SetExpresionCompiler(new ExpressionCompiler())
                    .SetIdProvider(new IntegerIdProvider())
                    .FunctionRepository.LoadModule(new BuiltInFunctions());
            });
        }

        public void Configure(Action<ParsingConfiguration> configMethod)
        {
            configMethod.Invoke(_parsingContext.Configuration);
            _lexer = _parsingContext.Configuration.Lexer ?? _lexer;
            _graphBuilder = _parsingContext.Configuration.GraphBuilder ?? _graphBuilder;
            _compiler = _parsingContext.Configuration.ExpressionCompiler ?? _compiler;
        }

        private ILexer _lexer;
        private IExpressionGraphBuilder _graphBuilder;
        private IExpressionCompiler _compiler;

        public ILexer Lexer { get { return _lexer; } }
        public IEnumerable<string> FunctionNames { get { return _parsingContext.Configuration.FunctionRepository.FunctionNames; } } 

        internal virtual object Parse(string formula, RangeAddress rangeAddress)
        {
            using (var scope = _parsingContext.Scopes.NewScope(rangeAddress))
            {
                var tokens = _lexer.Tokenize(formula);
                var graph = _graphBuilder.Build(tokens);
                if (graph.Expressions.Count() == 0)
                {
                    return null;
                }
                return _compiler.Compile(graph.Expressions).Result;
            }
        }

        internal virtual object Parse(IEnumerable<Token> tokens, string worksheet, string address)
        {
            var rangeAddress = _parsingContext.RangeAddressFactory.Create(address);
            using (var scope = _parsingContext.Scopes.NewScope(rangeAddress))
            {
                var graph = _graphBuilder.Build(tokens);
                if (graph.Expressions.Count() == 0)
                {
                    return null;
                }
                return _compiler.Compile(graph.Expressions).Result;
            }
        }
        internal virtual object ParseCell(IEnumerable<Token> tokens, string worksheet, int row, int column)
        {
            var rangeAddress = _parsingContext.RangeAddressFactory.Create(worksheet, column, row);
            using (var scope = _parsingContext.Scopes.NewScope(rangeAddress))
            {
                //    _parsingContext.Dependencies.AddFormulaScope(scope);
                var graph = _graphBuilder.Build(tokens);
                if (graph.Expressions.Count() == 0)
                {
                    return 0d;
                }
                try
                {
                    var compileResult = _compiler.Compile(graph.Expressions);
                    // quick solution for the fact that an excelrange can be returned.
                    var rangeInfo = compileResult.Result as ExcelDataProvider.IRangeInfo;
                    if (rangeInfo == null)
                    {
                        return compileResult.Result ?? 0d;
                    }
                    else
                    {
                        if (rangeInfo.IsEmpty)
                        {
                            return 0d;
                        }
                        if (!rangeInfo.IsMulti)
                        {
                            return rangeInfo.First().Value ?? 0d;
                        }
                        throw new ExcelErrorValueException(eErrorType.Value);
                    }
                }
                catch(ExcelErrorValueException ex)
                {
                    return ex.ErrorValue;
                }
            }
        }

        public virtual object Parse(string formula, string address)
        {
            return Parse(formula, _parsingContext.RangeAddressFactory.Create(address));
        }
        
        public virtual object Parse(string formula)
        {
            return Parse(formula, RangeAddress.Empty);
        }

        public virtual object ParseAt(string address)
        {
            Require.That(address).Named("address").IsNotNullOrEmpty();
            var rangeAddress = _parsingContext.RangeAddressFactory.Create(address);
            return ParseAt(rangeAddress.Worksheet, rangeAddress.FromRow, rangeAddress.FromCol);
        }

        public virtual object ParseAt(string worksheetName, int row, int col)
        {
            var f = _excelDataProvider.GetRangeFormula(worksheetName, row, col);
            if (string.IsNullOrEmpty(f))
            {
                return _excelDataProvider.GetRangeValue(worksheetName, row, col);
            }
            else
            {
                return Parse(f, _parsingContext.RangeAddressFactory.Create(worksheetName,col,row));
            }
            //var dataItem = _excelDataProvider.GetRangeValues(address).FirstOrDefault();
            //if (dataItem == null /*|| (dataItem.Value == null && dataItem.Formula == null)*/) return null;
            //if (!string.IsNullOrEmpty(dataItem.Formula))
            //{
            //    return Parse(dataItem.Formula, _parsingContext.RangeAddressFactory.Create(address));
            //}
            //return Parse(dataItem.Value.ToString(), _parsingContext.RangeAddressFactory.Create(address));
        }

    }
}
