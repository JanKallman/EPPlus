/*******************************************************************************
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
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public static class ExcelAddressUtil
    {
        static char[] SheetNameInvalidChars = new char[] { '?', ':', '*', '/', '\\' };
        public static bool IsValidAddress(string token)
        {
            int ix;
            if (token[0] == '\'')
            {
                ix = token.LastIndexOf('\'');
                if (ix > 0 && ix < token.Length - 1 && token[ix + 1] == '!')
                {
                    if (token.IndexOfAny(SheetNameInvalidChars, 1, ix - 1) > 0)
                    {
                        return false;
                    }
                    token = token.Substring(ix + 2);
                }
                else
                {
                    return false;
                }
            }
            else if ((ix = token.IndexOf('!')) > 1)
            {
                if (token.IndexOfAny(SheetNameInvalidChars, 0, token.IndexOf('!')) > 0)
                {
                    return false;
                }
                token = token.Substring(token.IndexOf('!') + 1);
            }
            return OfficeOpenXml.ExcelAddress.IsValidAddress(token);
        }
        readonly static char[] NameInvalidChars = new char[] { '!', '@', '#', '$', '£', '%', '&', '/', '(', ')', '[', ']', '{', '}', '<', '>', '=', '+', '*', '-', '~', '^', ':', ';', '|', ',', ' ' };
        public static bool IsValidName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }
            var fc = name[0];
            if (!(char.IsLetter(fc) || fc == '_' || (fc == '\\' && name.Length > 2)))
            {
                return false;
            }

            if (name.IndexOfAny(NameInvalidChars, 1) > 0)
            {
                return false;
            }

            if(ExcelCellBase.IsValidAddress(name))
            {
                return false;
            }

            //TODO:Add check for functionnames.
            return true;
        }
        public static string GetValidName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return name;
            }

            var fc = name[0];
            if (!(char.IsLetter(fc) || fc == '_' || (fc == '\\' && name.Length > 2)))
            {
                name = "_" + name.Substring(1);
            }

            name=NameInvalidChars.Aggregate(name, (c1, c2) => c1.Replace(c2, '_'));
            return name;
        }
    }
}
