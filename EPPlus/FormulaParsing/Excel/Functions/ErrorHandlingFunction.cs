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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Base class for functions that handles an error that occurs during the
    /// normal execution of the function.
    /// If an exception occurs during the Execute-call that exception will be
    /// caught by the compiler, then the HandleError-method will be called.
    /// </summary>
    public abstract class ErrorHandlingFunction : ExcelFunction
    {
        /// <summary>
        /// Indicates that the function is an ErrorHandlingFunction.
        /// </summary>
        public override bool IsErrorHandlingFunction
        {
            get
            {
                return true;
            }
        }

        /// <summary>
        /// Method that should be implemented to handle the error.
        /// </summary>
        /// <param name="errorCode"></param>
        /// <returns></returns>
        public abstract CompileResult HandleError(string errorCode);
    }
}
