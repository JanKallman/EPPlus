/* Copyright (C) 2011  Jan Källman
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
 *******************************************************************************
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public abstract class CollectionFlattener<T>
    {
        public virtual IEnumerable<T> FuncArgsToFlatEnumerable(IEnumerable<FunctionArgument> arguments, Action<FunctionArgument, IList<T>> convertFunc)
        {
            var argList = new List<T>();
            FuncArgsToFlatEnumerable(arguments, argList, convertFunc);
            return argList;
        }

        private void FuncArgsToFlatEnumerable(IEnumerable<FunctionArgument> arguments, List<T> argList, Action<FunctionArgument, IList<T>> convertFunc)
        {
            foreach (var arg in arguments)
            {
                if (arg.Value is IEnumerable<FunctionArgument>)
                {
                    FuncArgsToFlatEnumerable((IEnumerable<FunctionArgument>)arg.Value, argList, convertFunc);
                }
                else
                {
                    convertFunc(arg, argList);
                }
            }
        }
    }
}
