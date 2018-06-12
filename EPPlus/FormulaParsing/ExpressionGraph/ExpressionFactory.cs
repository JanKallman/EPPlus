﻿/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionFactory : IExpressionFactory
    {
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ParsingContext _parsingContext;

        public ExpressionFactory(ExcelDataProvider excelDataProvider, ParsingContext context)
        {
            _excelDataProvider = excelDataProvider;
            _parsingContext = context;
        }


        public Expression Create(Token token)
        {
            switch (token.TokenType)
            {
                case TokenType.Integer:
                    return new IntegerExpression(token.Value, token.IsNegated);
                case TokenType.String:
                    return new StringExpression(token.Value);
                case TokenType.Decimal:
                    return new DecimalExpression(token.Value, token.IsNegated);
                case TokenType.Boolean:
                    return new BooleanExpression(token.Value);
                case TokenType.ExcelAddress:
                    return new ExcelAddressExpression(token.Value, _excelDataProvider, _parsingContext, token.IsNegated);
                case TokenType.InvalidReference:
                    return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Ref));
                case TokenType.NumericError:
                    return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Num));
                case TokenType.ValueDataTypeError:
                    return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Value));
                case TokenType.Null:
                    return new ExcelErrorExpression(token.Value, ExcelErrorValue.Create(eErrorType.Null));
                case TokenType.NameValue:
                    return new NamedValueExpression(token.Value, _parsingContext);
                default:
                    return new StringExpression(token.Value);
            }
        }
    }
}
