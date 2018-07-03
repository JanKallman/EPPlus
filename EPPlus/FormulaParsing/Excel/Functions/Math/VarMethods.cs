﻿/* Copyright (C) 2011  Jan Källman
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
 * Mats Alm   		                Added		                2015-04-19
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    internal static class VarMethods
    {
        private static double Divide(double left, double right)
        {
            if (System.Math.Abs(right - 0d) < double.Epsilon)
            {
                throw new ExcelErrorValueException(eErrorType.Div0);
            }
            return left / right;
        }

        public static double Var(IEnumerable<ExcelDoubleCellValue> args)
        {
            return Var(args.Select(x => (double)x));
        }

        public static double Var(IEnumerable<double> args)
        {
            double avg = args.Select(x => (double)x).Average();
            double d = args.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
            return Divide(d, (args.Count() - 1));
        }

        public static double VarP(IEnumerable<ExcelDoubleCellValue> args)
        {
            return VarP(args.Select(x => (double)x));
        }

        public static double VarP(IEnumerable<double> args)
        {
            double avg = args.Select(x => (double)x).Average();
            double d = args.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
            return Divide(d, args.Count()); 
        }
    }
}
