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
//using ExcelFormulaParser.Engine.Excel.Functions;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public class Lexer : ILexer
    {
        public Lexer(Dictionary<string, object> functionRepository, Dictionary<string, object> nameValueProvider)
            :this(new SourceCodeTokenizer(functionRepository, nameValueProvider), new SyntacticAnalyzer())
        {

        }

        public Lexer(ISourceCodeTokenizer tokenizer, ISyntacticAnalyzer analyzer)
        {
            _tokenizer = tokenizer;
            _analyzer = analyzer;
        }

        private readonly ISourceCodeTokenizer _tokenizer;
        private readonly ISyntacticAnalyzer _analyzer;

        public IEnumerable<Token> Tokenize(string input)
        {
            var tokens = _tokenizer.Tokenize(input);
            _analyzer.Analyze(tokens);
            return tokens;
        }
    }
}
