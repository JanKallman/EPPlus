using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class AverageTests
	{
		[TestMethod]
		public void AverageLiterals()
		{
			// In the case of literals, Average DOES parse and include numeric strings, date strings, bools, etc.
			Average average = new Average();
			var date1 = new DateTime(2013, 1, 5);
			var date2 = new DateTime(2013, 1, 15);
			double value1 = 1000;
			double value2 = 2000;
			double value3 = 6000;
			double value4 = 1;
			double value5 = date1.ToOADate();
			double value6 = date2.ToOADate();
			var result = average.Execute(new FunctionArgument[]
			{
				new FunctionArgument(value1.ToString("n")),
				new FunctionArgument(value2),
				new FunctionArgument(value3.ToString("n")),
				new FunctionArgument(true),
				new FunctionArgument(date1),
				new FunctionArgument(date2.ToString("d"))
			}, ParsingContext.Create());
			Assert.AreEqual((value1 + value2 + value3 + value4 + value5 + value6) / 6, result.Result);
		}

		[TestMethod]
		public void AverageCellReferences()
		{
			// In the case of cell references, Average DOES NOT parse and include numeric strings, date strings, bools, unparsable strings, etc.
			ExcelPackage package = new ExcelPackage();
			var worksheet = package.Workbook.Worksheets.Add("Test");
			ExcelRange range1 = worksheet.Cells[1, 1];
			range1.Formula = "\"1000\"";
			range1.Calculate();
			var range2 = worksheet.Cells[1, 2];
			range2.Value = 2000;
			var range3 = worksheet.Cells[1, 3];
			range3.Formula = $"\"{new DateTime(2013, 1, 5).ToString("d")}\"";
			range3.Calculate();
			var range4 = worksheet.Cells[1, 4];
			range4.Value = true;
			var range5 = worksheet.Cells[1, 5];
			range5.Value = new DateTime(2013, 1, 5);
			var range6 = worksheet.Cells[1, 6];
			range6.Value = "Test";
			Average average = new Average();
			var rangeInfo1 = new EpplusExcelDataProvider.RangeInfo(worksheet, 1, 1, 1, 3);
			var rangeInfo2 = new EpplusExcelDataProvider.RangeInfo(worksheet, 1, 4, 1, 4);
			var rangeInfo3 = new EpplusExcelDataProvider.RangeInfo(worksheet, 1, 5, 1, 6);
			var context = ParsingContext.Create();
			var address = new OfficeOpenXml.FormulaParsing.ExcelUtilities.RangeAddress();
			address.FromRow = address.ToRow = address.FromCol = address.ToCol = 2;
			context.Scopes.NewScope(address);
			var result = average.Execute(new FunctionArgument[]
			{
				new FunctionArgument(rangeInfo1),
				new FunctionArgument(rangeInfo2),
				new FunctionArgument(rangeInfo3)
			}, context);
			Assert.AreEqual((2000 + new DateTime(2013, 1, 5).ToOADate()) / 2, result.Result);
		}

		[TestMethod]
		public void AverageArray()
		{
			// In the case of arrays, Average DOES NOT parse and include numeric strings, date strings, bools, unparsable strings, etc.
			Average average = new Average();
			var date1 = new DateTime(2013, 1, 5);
			var date2 = new DateTime(2013, 1, 15);
			double value = 2000;
			var result = average.Execute(new FunctionArgument[]
			{
				new FunctionArgument(new FunctionArgument[]
				{
					new FunctionArgument(1000.ToString("n")),
					new FunctionArgument(value),
					new FunctionArgument(6000.ToString("n")),
					new FunctionArgument(true),
					new FunctionArgument(date1),
					new FunctionArgument(date2.ToString("d")),
					new FunctionArgument("test")
				})
			}, ParsingContext.Create());
			Assert.AreEqual((2000 + date1.ToOADate()) / 2, result.Result);
		}

		[TestMethod]
		[ExpectedException(typeof(ExcelErrorValueException))]
		public void AverageUnparsableLiteral()
		{
			// In the case of literals, any unparsable string literal results in a #VALUE.
			Average average = new Average();
			var result = average.Execute(new FunctionArgument[]
			{
				new FunctionArgument(1000),
				new FunctionArgument("Test")
			}, ParsingContext.Create());
		}
	}
}
