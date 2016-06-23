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
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public interface IFunctionModule
    {
        /// <summary>
        /// Gets a dictionary of custom function implementations.
        /// </summary>
        IDictionary<string, ExcelFunction> Functions { get; }

        /// <summary>
        /// Gets a dictionary of custom function compilers. A function compiler is not 
        /// necessary for a custom function, unless the default expression evaluation is not
        /// sufficient for the implementation of the custom function. When a FunctionCompiler instance
        /// is created, it should be given a reference to the same function instance that exists
        /// in the Functions collection of this module.
        /// </summary>
        IDictionary<Type, FunctionCompiler> CustomCompilers { get; }
  }
}
