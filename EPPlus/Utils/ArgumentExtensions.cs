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
 * Mats Alm   		                Added       		        2011-01-01
 * Jan Källman		    License changed GPL-->LGPL  2011-12-27
 *******************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils
{
    /// <summary>
    /// Extension methods for guarding
    /// </summary>
    public static class ArgumentExtensions
    {

        /// <summary>
        /// Throws an ArgumentNullException if argument is null
        /// </summary>
        /// <typeparam name="T">Argument type</typeparam>
        /// <param name="argument">Argument to check</param>
        /// <param name="argumentName">parameter/argument name</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void IsNotNull<T>(this IArgument<T> argument, string argumentName)
            where T : class
        {
            argumentName = string.IsNullOrEmpty(argumentName) ? "value" : argumentName;
            if (argument.Value == null)
            {
                throw new ArgumentNullException(argumentName);
            }
        }

        /// <summary>
        /// Throws an <see cref="ArgumentNullException"/> if the string argument is null or empty
        /// </summary>
        /// <param name="argument">Argument to check</param>
        /// <param name="argumentName">parameter/argument name</param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void IsNotNullOrEmpty(this IArgument<string> argument, string argumentName)
        {
            if (string.IsNullOrEmpty(argument.Value))
            {
                throw new ArgumentNullException(argumentName);
            }
        }

        /// <summary>
        /// Throws an ArgumentOutOfRangeException if the value of the argument is out of the supplied range
        /// </summary>
        /// <typeparam name="T">Type implementing <see cref="IComparable"/></typeparam>
        /// <param name="argument">The argument to check</param>
        /// <param name="min">Min value of the supplied range</param>
        /// <param name="max">Max value of the supplied range</param>
        /// <param name="argumentName">parameter/argument name</param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public static void IsInRange<T>(this IArgument<T> argument, T min, T max, string argumentName)
            where T : IComparable
        {
            if (!(argument.Value.CompareTo(min) >= 0 && argument.Value.CompareTo(max) <= 0))
            {
                throw new ArgumentOutOfRangeException(argumentName);
            }
        }
    }
}
