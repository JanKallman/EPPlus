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
 * Jan Källman		    Added       		        2017-11-02
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.CompatibilityExtensions
{
    
    public static class DateTimeExtensions
    {
 #if Core
        // System.DateTime
        /// <summary>Converts the value of this instance to the equivalent OLE Automation date.</summary>
        /// <returns>A double-precision floating-point number that contains an OLE Automation date equivalent to the value of this instance.</returns>
        /// <exception cref="T:System.OverflowException">The value of this instance cannot be represented as an OLE Automation Date. </exception>
        /// <filterpriority>2</filterpriority>
        public static double ToOADate(this DateTime dt)
        {
            return TicksToOADate(dt.Ticks);
        }
        private static double TicksToOADate(long value)
        {
            if (value == 0L)
            {
                return 0.0;
            }
            if (value < 864000000000L)
            {
                value += 599264352000000000L;
            }
            if (value < 31241376000000000L)
            {
                throw new OverflowException("Value is too big");
                //throw new OverflowException(Environment.GetResourceString("Arg_OleAutDateInvalid"));
            }
            long num = (value - 599264352000000000L) / 10000L;
            if (num < 0L)
            {
                long num2 = num % 86400000L;
                if (num2 != 0L)
                {
                    num -= (86400000L + num2) * 2L;
                }
            }
            return (double)num / 86400000.0;
        }
#endif
        // System.DateTime
        /// <summary>Returns a <see cref="T:System.DateTime" /> equivalent to the specified OLE Automation Date.</summary>
        /// <returns>An object that represents the same date and time as <paramref name="d" />.</returns>
        /// <param name="d">An OLE Automation Date value. </param>
        /// <exception cref="T:System.ArgumentException">The date is not a valid OLE Automation Date value. </exception>
        /// <filterpriority>1</filterpriority>
        public static DateTime FromOADate(double d)
        {
#if Core
            return new DateTime(DoubleDateToTicks(d), DateTimeKind.Unspecified);
#else
            return DateTime.FromOADate(d);
#endif
        }
#if Core
        // System.DateTime
        internal static long DoubleDateToTicks(double value)
        {
            if (value >= 2958466.0 || value <= -657435.0)
            {
                throw new ArgumentException("Out of range");
                //throw new ArgumentException(Environment.GetResourceString("Arg_OleAutDateInvalid"));
            }
            long num = (long)(value * 86400000.0 + ((value >= 0.0) ? 0.5 : -0.5));
            if (num < 0L)
            {
                num -= num % 86400000L * 2L;
            }
            num += 59926435200000L;
            if (num < 0L || num >= 315537897600000L)
            {
                throw new ArgumentException("Out of range");
                //throw new ArgumentException(Environment.GetResourceString("Arg_OleAutDateScale"));
            }
            return num * 10000L;
        }
#endif
    }
}
