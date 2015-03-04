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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class CompileResultFactory
    {
        public virtual CompileResult Create(object obj)
        {
            if ((obj is ExcelDataProvider.INameInfo))
            {
                obj = ((ExcelDataProvider.INameInfo)obj).Value;
            }
            if (obj is ExcelDataProvider.IRangeInfo)
            {
                obj = ((ExcelDataProvider.IRangeInfo)obj).GetOffset(0, 0);
            }
            if (obj == null) return new CompileResult(null, DataType.Empty);
            if (obj.GetType().Equals(typeof(string)))
            {
                return new CompileResult(obj, DataType.String);
            }
            if (obj.GetType().Equals(typeof(double)) || obj is decimal)
            {
                return new CompileResult(obj, DataType.Decimal);
            }
            if (obj.GetType().Equals(typeof(int)) || obj is long || obj is short)
            {
                return new CompileResult(obj, DataType.Integer);
            }
            if (obj.GetType().Equals(typeof(bool)))
            {
                return new CompileResult(obj, DataType.Boolean);
            }
            if (obj.GetType().Equals(typeof (ExcelErrorValue)))
            {
                return new CompileResult(obj, DataType.ExcelError);
            }
            if (obj.GetType().Equals(typeof(System.DateTime)))
            {
                return new CompileResult(((System.DateTime)obj).ToOADate(), DataType.Date);
            }
            throw new ArgumentException("Non supported type " + obj.GetType().FullName);
        }
    }
}
