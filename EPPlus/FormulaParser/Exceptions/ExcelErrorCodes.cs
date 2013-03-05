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

namespace ExcelFormulaParser.Engine.Exceptions
{
    public class ExcelErrorCodes
    {
        private ExcelErrorCodes(string code)
        {
            Code = code;
        }

        public string Code
        {
            get;
            private set;
        }

        public override int GetHashCode()
        {
            return Code.GetHashCode();
        }

        public override bool  Equals(object obj)
        {
            if (obj is ExcelErrorCodes)
            {
                return ((ExcelErrorCodes)obj).Code.Equals(Code);
            }
 	        return false;
        }

        public static bool operator == (ExcelErrorCodes c1, ExcelErrorCodes c2)
        {
            return c1.Code.Equals(c2.Code);
        }

        public static bool operator !=(ExcelErrorCodes c1, ExcelErrorCodes c2)
        {
            return !c1.Code.Equals(c2.Code);
        }

        private static readonly IEnumerable<string> Codes = new List<string> { Value.Code, Name.Code, NoValueAvaliable.Code };

        public static bool IsErrorCode(object valueToTest)
        {
            if (valueToTest == null)
            {
                return false;
            }
            var candidate = valueToTest.ToString();
            if (Codes.FirstOrDefault(x => x == candidate) != null)
            {
                return true;
            }
            return false;
        }

        public static ExcelErrorCodes Value
        {
            get { return new ExcelErrorCodes("#VALUE!"); }
        }

        public static ExcelErrorCodes Name
        {
            get { return new ExcelErrorCodes("#NAME?"); }
        }

        public static ExcelErrorCodes NoValueAvaliable
        {
            get { return new ExcelErrorCodes("#N/A"); }
        }
    }
}
