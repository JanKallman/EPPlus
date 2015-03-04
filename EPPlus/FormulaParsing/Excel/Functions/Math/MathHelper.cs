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
 * Mats Alm   		                Added		                2015-01-11
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MathObj = System.Math;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    /// <summary>
    /// Thanks to the guys in this thread: http://stackoverflow.com/questions/2840798/c-sharp-math-class-question
    /// </summary>
    public static class MathHelper
    {
        // Secant 
        public static double Sec(double x)
        {
            return 1 / MathObj.Cos(x);
        }

        // Cosecant
        public static double Cosec(double x)
        {
            return 1 / MathObj.Sin(x);
        }

        // Cotangent 
        public static double Cotan(double x)
        {
            return 1 / MathObj.Tan(x);
        }

        // Inverse Sine 
        public static double Arcsin(double x)
        {
            return MathObj.Atan(x / MathObj.Sqrt(-x * x + 1));
        }

        // Inverse Cosine 
        public static double Arccos(double x)
        {
            return MathObj.Atan(-x / MathObj.Sqrt(-x * x + 1)) + 2 * MathObj.Atan(1);
        }


        // Inverse Secant 
        public static double Arcsec(double x)
        {
            return 2 * MathObj.Atan(1) - MathObj.Atan(MathObj.Sign(x) / MathObj.Sqrt(x * x - 1));
        }

        // Inverse Cosecant 
        public static double Arccosec(double x)
        {
            return MathObj.Atan(MathObj.Sign(x) / MathObj.Sqrt(x * x - 1));
        }

        // Inverse Cotangent 
        public static double Arccotan(double x)
        {
            return 2 * MathObj.Atan(1) - MathObj.Atan(x);
        }

        // Hyperbolic Sine 
        public static double HSin(double x)
        {
            return (MathObj.Exp(x) - MathObj.Exp(-x)) / 2;
        }

        // Hyperbolic Cosine 
        public static double HCos(double x)
        {
            return (MathObj.Exp(x) + MathObj.Exp(-x)) / 2;
        }

        // Hyperbolic Tangent 
        public static double HTan(double x)
        {
            return (MathObj.Exp(x) - MathObj.Exp(-x)) / (MathObj.Exp(x) + MathObj.Exp(-x));
        }

        // Hyperbolic Secant 
        public static double HSec(double x)
        {
            return 2 / (MathObj.Exp(x) + MathObj.Exp(-x));
        }

        // Hyperbolic Cosecant 
        public static double HCosec(double x)
        {
            return 2 / (MathObj.Exp(x) - MathObj.Exp(-x));
        }

        // Hyperbolic Cotangent 
        public static double HCotan(double x)
        {
            return (MathObj.Exp(x) + MathObj.Exp(-x)) / (MathObj.Exp(x) - MathObj.Exp(-x));
        }

        // Inverse Hyperbolic Sine 
        public static double HArcsin(double x)
        {
            return MathObj.Log(x + MathObj.Sqrt(x * x + 1));
        }

        // Inverse Hyperbolic Cosine 
        public static double HArccos(double x)
        {
            return MathObj.Log(x + MathObj.Sqrt(x * x - 1));
        }

        // Inverse Hyperbolic Tangent 
        public static double HArctan(double x)
        {
            return MathObj.Log((1 + x) / (1 - x)) / 2;
        }

        // Inverse Hyperbolic Secant 
        public static double HArcsec(double x)
        {
            return MathObj.Log((MathObj.Sqrt(-x * x + 1) + 1) / x);
        }

        // Inverse Hyperbolic Cosecant 
        public static double HArccosec(double x)
        {
            return MathObj.Log((MathObj.Sign(x) * MathObj.Sqrt(x * x + 1) + 1) / x);
        }

        // Inverse Hyperbolic Cotangent 
        public static double HArccotan(double x)
        {
            return MathObj.Log((x + 1) / (x - 1)) / 2;
        }

        // Logarithm to base N 
        public static double LogN(double x, double n)
        {
            return MathObj.Log(x) / MathObj.Log(n);
        }
    }
}
