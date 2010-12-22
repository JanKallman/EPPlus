using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace ExcelPackageTest
{
	[TestClass]
	public class ExcelExtensionTests
	{
		#region Test Initialize
		private static ExcelPackage _pck;
		//
		// You can use the following additional attributes as you write your tests:
		//
		// Use ClassInitialize to run code before running the first test in the class
		[ClassInitialize]
		public static void ExcelExtensionTestsInitialize(TestContext testContext)
		{
			_pck = new ExcelPackage(new FileInfo("Test\\Worksheet.xlsx"));
		}

		// Use ClassCleanup to run code after all tests in a class have run
		[ClassCleanup]
		public static void ExcelExtensionTestsCleanup()
		{
			_pck = null;
		}
		#endregion

		#region ExcelRange Tests
		[TestMethod]
		public void SetValueTest()
		{
			// Arrange
			var sheet = _pck.Workbook.Worksheets.Add("newsheet");
			var cell = sheet.GetCell(1, 1);
			var value = "test";

			// Act
			var result = cell.SetValue(value);

			// Assert
			Assert.AreEqual(result.Value, value);
			Assert.IsNotNull(result.Value);
		}

		[TestMethod]
		public void GetColumnTest()
		{
			// Arrange
			var sheet = _pck.Workbook.Worksheets.Add("newsheet");
			var cell = sheet.GetCell(1, 1);
			
			// Act
			var result = cell.Column();

			// Assert
			Assert.AreEqual(result.GetType(), typeof(ExcelColumn));
			Assert.IsNotNull(result);
		}

		[TestMethod]
		public void GetRowTest()
		{
			// Arrange
			var sheet = _pck.Workbook.Worksheets.Add("newsheet");
			var cell = sheet.GetCell(1, 1);
			
			// Act
			var result = cell.Row();

			// Assert
			Assert.AreEqual(result.GetType(), typeof(ExcelRow));
			Assert.IsNotNull(result);
		}

		[TestMethod]
		public void SetFormular()
		{
			// Arrange

			// Act

			// Assert
		}
		#endregion
	}
}
