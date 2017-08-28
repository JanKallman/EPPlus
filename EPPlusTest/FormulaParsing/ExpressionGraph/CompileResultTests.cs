using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
	[TestClass]
	public class CompileResultTests
	{
		[TestMethod]
		public void NumericStringCompileResult()
		{
			var expected = 124.24;
			string numericString = expected.ToString("n");
			CompileResult result = new CompileResult(numericString, DataType.String);
			Assert.IsFalse(result.IsNumeric);
			Assert.IsTrue(result.IsNumericString);
			Assert.AreEqual(expected, result.ResultNumeric);
		}

		[TestMethod]
		public void DateStringCompileResult()
		{
			var expected = new DateTime(2013, 1, 15);
			string dateString = expected.ToString("d");
			CompileResult result = new CompileResult(dateString, DataType.String);
			Assert.IsFalse(result.IsNumeric);
			Assert.IsTrue(result.IsDateString);
			Assert.AreEqual(expected.ToOADate(), result.ResultNumeric);
		}
	}
}
