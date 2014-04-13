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

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class TokenizerContext
    {
        public TokenizerContext(string formula)
        {
            if (!string.IsNullOrEmpty(formula))
            {
                _chars = formula.ToArray();
            }
            _result = new List<Token>();
            _currentToken = new StringBuilder();
        }

        private char[] _chars;
        private List<Token> _result;
        private StringBuilder _currentToken;

        public char[] FormulaChars
        {
            get { return _chars; }
        }

        public IList<Token> Result
        {
            get { return _result; }
        }

        public bool IsInString
        {
            get;
            private set;
        }

        public void ToggleIsInString()
        {
            IsInString = !IsInString;
        }

        internal int BracketCount
        {
            get;
            set;
        }

        public string CurrentToken
        {
            get { return _currentToken.ToString(); }
        }

        public bool CurrentTokenHasValue
        {
            get { return !string.IsNullOrEmpty(CurrentToken.Trim()); }
        }

        public void NewToken()
        {
            _currentToken = new StringBuilder();
        }

        public void AddToken(Token token)
        {
            _result.Add(token);
        }

        public void AppendToCurrentToken(char c)
        {
            _currentToken.Append(c.ToString());
        }

        public void AppendToLastToken(string stringToAppend)
        {
            _result.Last().Append(stringToAppend);
        }

        public void ReplaceLastToken(Token newToken)
        {
            if (_result.Count > 0)
            {
                _result.RemoveAt(_result.Count - 1);   
            }
            _result.Add(newToken);
        }

        public Token LastToken
        {
            get { return _result.Count > 0 ? _result.Last() : null; }
        }

    }
}
