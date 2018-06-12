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
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class IntegerExpression : AtomicExpression
    {
        private readonly double? _compiledValue;
        private readonly bool _negate;

        public IntegerExpression(string expression)
            : this(expression, false)
        {

        }

        public IntegerExpression(string expression, bool negate)
            : base(expression)
        {
            _negate = negate;
        }

        public IntegerExpression(double val)
            : base(val.ToString(CultureInfo.InvariantCulture))
        {
            _compiledValue = Math.Floor(val);
        }

        public override CompileResult Compile()
        {
            double result = _compiledValue.HasValue ? _compiledValue.Value : double.Parse(ExpressionString, CultureInfo.InvariantCulture);
            result = _negate ? result*-1 : result;
            return new CompileResult(result, DataType.Integer);
        }
    }
}
