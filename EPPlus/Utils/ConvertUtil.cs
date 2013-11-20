using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils
{
    internal static class ConvertUtil
    {
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
                if ((v.GetType().IsPrimitive || v is double || v is decimal || v is DateTime || v is TimeSpan))
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
