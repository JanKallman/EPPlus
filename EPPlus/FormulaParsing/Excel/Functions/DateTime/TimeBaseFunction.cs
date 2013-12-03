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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public abstract class TimeBaseFunction : ExcelFunction
    {
        public TimeBaseFunction()
        {
            TimeStringParser = new TimeStringParser();
        }

        protected TimeStringParser TimeStringParser
        {
            get;
            private set;
        }

        protected double SerialNumber
        {
            get;
            private set;
        }

        public void ValidateAndInitSerialNumber(IEnumerable<FunctionArgument> arguments)
        {
            ValidateArguments(arguments, 1);
            SerialNumber = (double)ArgToDecimal(arguments, 0);
        }

        protected double SecondsInADay
        {
            get{ return 24 * 60 * 60; }
        }

        protected double GetTimeSerialNumber(double seconds)
        {
            return seconds / SecondsInADay;
        }

        protected double GetSeconds(double serialNumber)
        {
            return serialNumber * SecondsInADay;
        }

        protected double GetHour(double serialNumber)
        {
            var seconds = GetSeconds(serialNumber);
            return (int)seconds / (60 * 60);
        }

        protected double GetMinute(double serialNumber)
        {
            var seconds = GetSeconds(serialNumber);
            seconds -= GetHour(serialNumber) * 60 * 60;
            return (seconds - (seconds % 60)) / 60;
        }

        protected double GetSecond(double serialNumber)
        {
            return GetSeconds(serialNumber) % 60;
        }
    }
}
