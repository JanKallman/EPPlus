using System;
using System.ComponentModel;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Compatibility;

namespace EPPlusTest.Utils
{
	[TestClass]
	public class ConvertUtilTest
	{
		[TestMethod]
		public void TryParseNumericString()
		{
			double result;
			object numericString = null;
			double expected = 0;
			Assert.IsFalse(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
			expected = 1442.0;
			numericString = expected.ToString("e", CultureInfo.CurrentCulture); // 1.442E+003
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
			numericString = expected.ToString("f0", CultureInfo.CurrentCulture); // 1442
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
			numericString = expected.ToString("f2", CultureInfo.CurrentCulture); // 1442.00
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
			numericString = expected.ToString("n", CultureInfo.CurrentCulture); // 1,442.0
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
			expected = -0.00526;
			numericString = expected.ToString("e", CultureInfo.CurrentCulture); // -5.26E-003
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
			numericString = expected.ToString("f0", CultureInfo.CurrentCulture); // -0
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(0.0, result);
			numericString = expected.ToString("f3", CultureInfo.CurrentCulture); // -0.005
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(-0.005, result);
			numericString = expected.ToString("n6", CultureInfo.CurrentCulture); // -0.005260
			Assert.IsTrue(ConvertUtil.TryParseNumericString(numericString, out result));
			Assert.AreEqual(expected, result);
		}
		
		[TestMethod]
		public void TryParseDateString()
		{
			DateTime result;
			object dateString = null;
			DateTime expected = DateTime.MinValue;
			Assert.IsFalse(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(expected, result);
			expected = new DateTime(2013, 1, 15);
			dateString = expected.ToString("d", CultureInfo.CurrentCulture); // 1/15/2013
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(expected, result);
			dateString = expected.ToString("D", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(expected, result);
			dateString = expected.ToString("F", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013 12:00:00 AM
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(expected, result);
			dateString = expected.ToString("g", CultureInfo.CurrentCulture); // 1/15/2013 12:00 AM
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(expected, result);
			expected = new DateTime(2013, 1, 15, 15, 26, 32);
			dateString = expected.ToString("F", CultureInfo.CurrentCulture); // Tuesday, January 15, 2013 3:26:32 PM
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(expected, result);
			dateString = expected.ToString("g", CultureInfo.CurrentCulture); // 1/15/2013 3:26 PM
			Assert.IsTrue(ConvertUtil.TryParseDateString(dateString, out result));
			Assert.AreEqual(new DateTime(2013, 1, 15, 15, 26, 0), result);
		}

        [TestMethod]
        public void TextToInt()
        {
            var result = ConvertUtil.GetTypedCellValue<int>("204");

            Assert.AreEqual(204, result);
        }
        // This is just illustration of the bug in old implementation
        //[TestMethod]
        public void TextToIntInOldImplementation()
        {
            var result = GetTypedValue<int>("204");

            Assert.AreEqual(204, result);
        }
        [TestMethod]
        public void DoubleToNullableInt()
        {
            var result = ConvertUtil.GetTypedCellValue<int?>(2D);

            Assert.AreEqual(2, result);
        }

        [TestMethod]
        public void StringToDecimal()
        {
            var decimalSign=System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            var result = ConvertUtil.GetTypedCellValue<decimal>($"1{decimalSign}4");

            Assert.AreEqual((decimal)1.4, result);
        }

        [TestMethod]
        public void EmptyStringToNullableDecimal()
        {
            var result = ConvertUtil.GetTypedCellValue<decimal?>("");
            Assert.IsNull(result);
        }

        [TestMethod]
        public void BlankStringToNullableDecimal()
        {
            var result = ConvertUtil.GetTypedCellValue<decimal?>(" ");

            Assert.IsNull(result);
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException))]
        public void EmptyStringToDecimal()
        {
            ConvertUtil.GetTypedCellValue<decimal>("");
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException))]
        public void FloatingPointStringToInt()
        {
            ConvertUtil.GetTypedCellValue<int>("1.4");
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidCastException))]
        public void IntToDateTime()
        {
            ConvertUtil.GetTypedCellValue<DateTime>(122);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidCastException))]
        public void IntToTimeSpan()
        {
            ConvertUtil.GetTypedCellValue<TimeSpan>(122);
        }

        [TestMethod]
        public void IntStringToTimeSpan()
        {
            Assert.AreEqual(TimeSpan.FromDays(122), ConvertUtil.GetTypedCellValue<TimeSpan>("122"));
        }

        [TestMethod]
        public void BoolToInt()
        {
            Assert.AreEqual(1, ConvertUtil.GetTypedCellValue<int>(true));
            Assert.AreEqual(0, ConvertUtil.GetTypedCellValue<int>(false));
        }

        [TestMethod]
        public void BoolToDecimal()
        {
            Assert.AreEqual(1m, ConvertUtil.GetTypedCellValue<decimal>(true));
            Assert.AreEqual(0m, ConvertUtil.GetTypedCellValue<decimal>(false));
        }

        [TestMethod]
        public void BoolToDouble()
        {
            Assert.AreEqual(1d, ConvertUtil.GetTypedCellValue<double>(true));
            Assert.AreEqual(0d, ConvertUtil.GetTypedCellValue<double>(false));
        }

        [TestMethod]
        [ExpectedException(typeof(FormatException))]
        public void BadTextToInt()
        {
            ConvertUtil.GetTypedCellValue<int>("text1");
        }

        // previous implementation
        internal T GetTypedValue<T>(object v)
        {
            if (v == null)
            {
                return default(T);
            }
            Type fromType = v.GetType();
            Type toType = typeof(T);
            
            Type toType2 = (TypeCompat.IsGenericType(toType) && toType.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
                ? Nullable.GetUnderlyingType(toType)
                : null;
            if (fromType == toType || fromType == toType2)
            {
                return (T)v;
            }
            var cnv = TypeDescriptor.GetConverter(fromType);
            if (toType == typeof(DateTime) || toType2 == typeof(DateTime))    //Handle dates
            {
                if (fromType == typeof(TimeSpan))
                {
                    return ((T)(object)(new DateTime(((TimeSpan)v).Ticks)));
                }
                else if (fromType == typeof(string))
                {
                    DateTime dt;
                    if (DateTime.TryParse(v.ToString(), out dt))
                    {
                        return (T)(object)(dt);
                    }
                    else
                    {
                        return default(T);
                    }

                }
                else
                {
                    if (cnv.CanConvertTo(typeof(double)))
                    {
                        return (T)(object)(DateTime.FromOADate((double)cnv.ConvertTo(v, typeof(double))));
                    }
                    else
                    {
                        return default(T);
                    }
                }
            }
            else if (toType == typeof(TimeSpan) || toType2 == typeof(TimeSpan))    //Handle timespan
            {
                if (fromType == typeof(DateTime))
                {
                    return ((T)(object)(new TimeSpan(((DateTime)v).Ticks)));
                }
                else if (fromType == typeof(string))
                {
                    TimeSpan ts;
                    if (TimeSpan.TryParse(v.ToString(), out ts))
                    {
                        return (T)(object)(ts);
                    }
                    else
                    {
                        return default(T);
                    }
                }
                else
                {
                    if (cnv.CanConvertTo(typeof(double)))
                    {

                        return (T)(object)(new TimeSpan(DateTime.FromOADate((double)cnv.ConvertTo(v, typeof(double))).Ticks));
                    }
                    else
                    {
                        try
                        {
                            // Issue 14682 -- "GetValue<decimal>() won't convert strings"
                            // As suggested, after all special cases, all .NET to do it's 
                            // preferred conversion rather than simply returning the default
                            return (T)Convert.ChangeType(v, typeof(T));
                        }
                        catch (Exception)
                        {
                            // This was the previous behaviour -- no conversion is available.
                            return default(T);
                        }
                    }
                }
            }
            else
            {
                if (cnv.CanConvertTo(toType))
                {
                    return (T)cnv.ConvertTo(v, typeof(T));
                }
                else
                {
                    if (toType2 != null)
                    {
                        toType = toType2;
                        if (cnv.CanConvertTo(toType))
                        {
                            return (T)cnv.ConvertTo(v, toType); //Fixes issue 15377
                        }
                    }

                    if (fromType == typeof(double) && toType == typeof(decimal))
                    {
                        return (T)(object)Convert.ToDecimal(v);
                    }
                    else if (fromType == typeof(decimal) && toType == typeof(double))
                    {
                        return (T)(object)Convert.ToDouble(v);
                    }
                    else
                    {
                        return default(T);
                    }
                }
            }
        }

    }
}
