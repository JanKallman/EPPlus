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

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class ValueMatcher
    {
        public const int IncompatibleOperands = -2;

        public virtual int IsMatch(object o1, object o2)
        {
            if (o1 != null && o2 == null) return 1;
            if (o1 == null && o2 != null) return -1;
            if (o1 == null && o2 == null) return 0;
            if (o1 is string && o2 is string)
            {
                return CompareStringToString(o1.ToString().ToLower(), o2.ToString().ToLower());
            }
            else if( o1.GetType() == typeof(string))
            {
                return CompareStringToObject(o1.ToString(), o2);
            }
            else if (o2.GetType() == typeof(string))
            {
                return CompareObjectToString(o1, o2.ToString());
            }
            return Convert.ToDouble(o1).CompareTo(Convert.ToDouble(o2));
        }

        protected virtual int CompareStringToString(string s1, string s2)
        {
            return s1.CompareTo(s2);
        }

        protected virtual int CompareStringToObject(string o1, object o2)
        {
            double d1;
            if (double.TryParse(o1, out d1))
            {
                return d1.CompareTo(Convert.ToDouble(o2));
            }
            return IncompatibleOperands;
        }

        protected virtual int CompareObjectToString(object o1, string o2)
        {
            double d2;
            if (double.TryParse(o2, out d2))
            {
                return Convert.ToDouble(o1).CompareTo(d2);
            }
            return IncompatibleOperands;
        }
    }
}
