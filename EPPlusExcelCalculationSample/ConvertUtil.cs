using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace EPPlusExcelFormDemo
{
    internal static class ConvertUtil
    {
        internal static bool IsNumeric(object candidate)
        {
            if (candidate == null) return false;
            return (candidate.GetType().IsPrimitive || candidate is double || candidate is decimal || candidate is DateTime || candidate is TimeSpan || candidate is long);
        }

        internal static bool IsNumericString(object candidate)
        {
            if (candidate != null)
            {
                return Regex.IsMatch(candidate.ToString(), @"^[\d]+(\,[\d])?");
            }
            return false;
        }

        /// <summary>
        /// Convert an object value to a double 
        /// </summary>
        /// <param name="v"></param>
        /// <param name="ignoreBool"></param>
        /// <returns></returns>
        internal static double GetValueDouble(object v, bool ignoreBool = false)
        {
            double d;
            try
            {
                if (ignoreBool && v is bool)
                {
                    return 0;
                }
                if (IsNumeric(v))
                {
                    if (v is DateTime)
                    {
                        d = ((DateTime)v).ToOADate();
                    }
                    else if (v is TimeSpan)
                    {
                        d = new DateTime(((TimeSpan)v).Ticks).ToOADate();
                    }
                    else
                    {
                        d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                    }
                }
                else
                {
                    d = 0;
                }
            }

            catch
            {
                d = 0;
            }
            return d;
        }
    }
}
