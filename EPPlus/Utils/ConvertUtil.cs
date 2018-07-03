using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.IO;
using OfficeOpenXml.Compatibility;

namespace OfficeOpenXml.Utils
{
    internal static class ConvertUtil
    {
        internal static bool IsNumeric(object candidate)
        {
            if (candidate == null) return false;
            return (TypeCompat.IsPrimitive(candidate) || candidate is double || candidate is decimal || candidate is DateTime || candidate is TimeSpan || candidate is long);
        }
		/// <summary>
		/// Tries to parse a double from the specified <paramref name="candidate"/> which is expected to be a string value.
		/// </summary>
		/// <param name="candidate">The string value.</param>
		/// <param name="result">The double value parsed from the specified <paramref name="candidate"/>.</param>
		/// <returns>True if <paramref name="candidate"/> could be parsed to a double; otherwise, false.</returns>
		internal static bool TryParseNumericString(object candidate, out double result)
		{
			if (candidate != null)
			{
				// If a number is stored in a string, Excel will not convert it to the invariant format, so assume that it is in the current culture's number format.
				// This may not always be true, but it is a better assumption than assuming it is always in the invariant culture, which will probably never be true
				// for locales outside the United States.
				var style = NumberStyles.Float | NumberStyles.AllowThousands;
				return double.TryParse(candidate.ToString(), style, CultureInfo.CurrentCulture, out result);
			}
			result = 0;
			return false;
		}
		/// <summary>
		/// Tries to parse a boolean value from the specificed <paramref name="candidate"/>.
		/// </summary>
		/// <param name="candidate">The value to check for boolean-ness.</param>
		/// <param name="result">The boolean value parsed from the specified <paramref name="candidate"/>.</param>
		/// <returns>True if <paramref name="candidate"/> could be parsed </returns>
		internal static bool TryParseBooleanString(object candidate, out bool result)
		{
			if (candidate != null)
				return bool.TryParse(candidate.ToString(), out result);
			result = false;
			return false;
		}
		/// <summary>
		/// Tries to parse a <see cref="DateTime"/> from the specified <paramref name="candidate"/> which is expected to be a string value.
		/// </summary>
		/// <param name="candidate">The string value.</param>
		/// <param name="result">The double value parsed from the specified <paramref name="candidate"/>.</param>
		/// <returns>True if <paramref name="candidate"/> could be parsed to a double; otherwise, false.</returns>
		internal static bool TryParseDateString(object candidate, out DateTime result)
		{
			if (candidate != null)
			{
				var style = DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal;
				// If a date is stored in a string, Excel will not convert it to the invariant format, so assume that it is in the current culture's date/time format.
				// This may not always be true, but it is a better assumption than assuming it is always in the invariant culture, which will probably never be true
				// for locales outside the United States.
				return DateTime.TryParse(candidate.ToString(), CultureInfo.CurrentCulture, style, out result);
			}
			result = DateTime.MinValue;
			return false;
		}
		/// <summary>
		/// Convert an object value to a double 
		/// </summary>
		/// <param name="v"></param>
		/// <param name="ignoreBool"></param>
        /// <param name="retNaN">Return NaN if invalid double otherwise 0</param>
		/// <returns></returns>
		internal static double GetValueDouble(object v, bool ignoreBool = false, bool retNaN=false)
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
                        d = DateTime.FromOADate(0).Add((TimeSpan)v).ToOADate();
                    }
                    else
                    {
                        d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                    }
                }
                else
                {
                    d = retNaN ? double.NaN : 0;
                }
            }

            catch
            {
                d = retNaN ? double.NaN : 0;
            }
            return d;
        }
        /// <summary>
        /// OOXML requires that "," , and &amp; be escaped, but ' and " should *not* be escaped, nor should
        /// any extended Unicode characters. This function only encodes the required characters.
        /// System.Security.SecurityElement.Escape() escapes ' and " as  &apos; and &quot;, so it cannot
        /// be used reliably. System.Web.HttpUtility.HtmlEncode overreaches as well and uses the numeric
        /// escape equivalent.
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        internal static string ExcelEscapeString(string s)
        {
            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }
        /// <summary>
        /// Return true if preserve space attribute is set.
        /// </summary>
        /// <param name="sw"></param>
        /// <param name="t"></param>
        /// <returns></returns>
        internal static void ExcelEncodeString(StreamWriter sw, string t)
        {
            if (Regex.IsMatch(t, "(_x[0-9A-F]{4,4}_)"))
            {
                var match = Regex.Match(t, "(_x[0-9A-F]{4,4}_)");
                int indexAdd = 0;
                while (match.Success)
                {
                    t = t.Insert(match.Index + indexAdd, "_x005F");
                    indexAdd += 6;
                    match = match.NextMatch();
                }
            }
            for (int i = 0; i < t.Length; i++)
            {
                if (t[i] <= 0x1f && t[i] != '\t' && t[i] != '\n' && t[i] != '\r') //Not Tab, CR or LF
                {
                    sw.Write("_x00{0}_", (t[i] < 0xf ? "0" : "") + ((int)t[i]).ToString("X"));
                }
                else
                {
                    sw.Write(t[i]);
                }
            }

        }
        /// <summary>
        /// Return true if preserve space attribute is set.
        /// </summary>
        /// <param name="sb"></param>
        /// <param name="t"></param>
        /// <param name="encodeTabCRLF"></param>
        /// <returns></returns>
        internal static void ExcelEncodeString(StringBuilder sb, string t, bool encodeTabCRLF=false)
        {
            if (Regex.IsMatch(t, "(_x[0-9A-F]{4,4}_)"))
            {
                var match = Regex.Match(t, "(_x[0-9A-F]{4,4}_)");
                int indexAdd = 0;
                while (match.Success)
                {
                    t = t.Insert(match.Index + indexAdd, "_x005F");
                    indexAdd += 6;
                    match = match.NextMatch();
                }
            }
            for (int i = 0; i < t.Length; i++)
            {
                if (t[i] <= 0x1f && ((t[i] != '\t' && t[i] != '\n' && t[i] != '\r' && encodeTabCRLF == false) || encodeTabCRLF)) //Not Tab, CR or LF
                {
                    sb.AppendFormat("_x00{0}_", (t[i] < 0xf ? "0" : "") + ((int)t[i]).ToString("X"));
                }
                else
                {
                    sb.Append(t[i]);
                }
            }

        }
        /// <summary>
        /// Return true if preserve space attribute is set.
        /// </summary>
        /// <param name="t"></param>
        /// <returns></returns>
        internal static string ExcelEncodeString(string t)
        {
            StringBuilder sb=new StringBuilder();
            t=t.Replace("\r\n", "\n"); //For some reason can't table name have cr in them. Replace with nl
            ExcelEncodeString(sb, t, true);
            return sb.ToString();
        }
        internal static string ExcelDecodeString(string t)
        {
            var match = Regex.Match(t, "(_x005F|_x[0-9A-F]{4,4}_)");
            if (!match.Success) return t;

            var useNextValue = false;
            var ret = new StringBuilder();
            var prevIndex = 0;
            while (match.Success)
            {
                if (prevIndex < match.Index) ret.Append(t.Substring(prevIndex, match.Index - prevIndex));
                if (!useNextValue && match.Value == "_x005F")
                {
                    useNextValue = true;
                }
                else
                {
                    if (useNextValue)
                    {
                        ret.Append(match.Value);
                        useNextValue = false;
                    }
                    else
                    {
                        ret.Append((char)int.Parse(match.Value.Substring(2, 4), NumberStyles.AllowHexSpecifier));
                    }
                }
                prevIndex = match.Index + match.Length;
                match = match.NextMatch();
            }
            ret.Append(t.Substring(prevIndex, t.Length - prevIndex));
            return ret.ToString();
        }

        /// <summary>
        ///     Convert cell value to desired type, including nullable structs.
        ///     When converting blank string to nullable struct (e.g. ' ' to int?) null is returned.
        ///     When attempted conversion fails exception is passed through.
        /// </summary>
        /// <typeparam name="T">
        ///     The type to convert to.
        /// </typeparam>
        /// <returns>
        ///     The <paramref name="value"/> converted to <typeparamref name="T"/>.
        /// </returns>
        /// <remarks>
        ///     If input is string, parsing is performed for output types of DateTime and TimeSpan, which if fails throws <see cref="FormatException"/>.
        ///     Another special case for output types of DateTime and TimeSpan is when input is double, in which case <see cref="DateTime.FromOADate"/>
        ///     is used for conversion. This special case does not work through other types convertible to double (e.g. integer or string with number).
        ///     In all other cases 'direct' conversion <see cref="Convert.ChangeType(object, Type)"/> is performed.
        /// </remarks>
        /// <exception cref="FormatException">
        ///     <paramref name="value"/> is string and its format is invalid for conversion (parsing fails)
        /// </exception>
        /// <exception cref="InvalidCastException">
        ///     <paramref name="value"/> is not string and direct conversion fails
        /// </exception>
        public static T GetTypedCellValue<T>(object value)
        {
            if (value == null)
                return default(T);

            var fromType = value.GetType();
            var toType = typeof(T);
            var toNullableUnderlyingType = (TypeCompat.IsGenericType(toType) && toType.GetGenericTypeDefinition() == typeof(Nullable<>))
                ? Nullable.GetUnderlyingType(toType)
                : null;

            if (fromType == toType || fromType == toNullableUnderlyingType)
                return (T)value;

            // if converting to nullable struct and input is blank string, return null
            if (toNullableUnderlyingType != null && fromType == typeof(string) && ((string)value).Trim() == string.Empty)
                return default(T);

            toType = toNullableUnderlyingType ?? toType;

            if (toType == typeof(DateTime))
            {
                if (value is double)
                    return (T)(object)(DateTime.FromOADate((double)value));

                if (fromType == typeof(TimeSpan))
                    return ((T)(object)(new DateTime(((TimeSpan)value).Ticks)));

                if (fromType == typeof(string))
                    return (T)(object)DateTime.Parse(value.ToString());
            }
            else if (toType == typeof(TimeSpan))
            {
                if (value is double)
                    return (T)(object)(new TimeSpan(DateTime.FromOADate((double)value).Ticks));

                if (fromType == typeof(DateTime))
                    return ((T)(object)(new TimeSpan(((DateTime)value).Ticks)));

                if (fromType == typeof(string))
                    return (T)(object)TimeSpan.Parse(value.ToString());
            }

            return (T)Convert.ChangeType(value, toType);
        }

        #region internal cache objects
        internal static TextInfo _invariantTextInfo = CultureInfo.InvariantCulture.TextInfo;
        internal static CompareInfo _invariantCompareInfo = CompareInfo.GetCompareInfo(CultureInfo.InvariantCulture.Name);  //TODO:Check that it works
        #endregion
    }
}
