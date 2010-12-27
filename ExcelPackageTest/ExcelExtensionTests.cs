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
		private const string TestValue = "test";

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

			// Act
			var result = cell.SetValue(TestValue);

			// Assert
			Assert.AreEqual(result.Value, TestValue);
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

		#region ExcelWorksheet Tests
		[TestMethod]
		public void AddIEnumberable()
		{
			// Arrange
			var sheet = _pck.Workbook.Worksheets.Add("newsheet");
			var startingCell = sheet.GetCell(1, 1);
			var listOfData = new List<string> {"test1", "test2", "test3"};

			// Act
			var modifiedCells = sheet.Add(startingCell, InsertDirection.Across, listOfData);

			// Assert
			Assert.AreEqual(modifiedCells.GetType(), typeof(ExcelRange));
			Assert.AreEqual(modifiedCells.Value, listOfData[0]);
		}

		[TestMethod]
		public void GetCellValue()
		{
			// Arrange
			// create new sheet
			var sheet = _pck.Workbook.Worksheets.Add("newsheet");
			// set cell value
			sheet.SetCellValue(1, 1, TestValue);

			// Act
			var result = sheet.GetCellValue<string>(1, 1);

			// Assert
			Assert.AreEqual(result, TestValue);
			Assert.AreNotEqual(result, 123);
		}

		[TestMethod]
		public void GetCell()
		{
			// Arrange
			// create new sheet
			var sheet = _pck.Workbook.Worksheets.Add("newsheet");
			// set cell value
			sheet.SetCellValue(1, 1, TestValue);

			// Act
			var result = sheet.GetCell(1, 1);

			// Assert
			Assert.AreEqual(result.GetType(), typeof (ExcelRange));
			Assert.AreEqual(result.Value, TestValue);
		}

		[TestMethod]
		public void GetColumn()
		{
			// Arrange
			// create new sheet
			var sheet = _pck.Workbook.Worksheets.Add("newsheet");

			// Act
			var result = sheet.GetColumn(1);

			// Assert
			Assert.AreEqual(result.GetType(), typeof (ExcelColumn));
		}

		[TestMethod]
		public void GetRow()
		{
			// Arrange
			// create new sheet
			var sheet = _pck.Workbook.Worksheets.Add("newsheet");

			// Act
			var result = sheet.GetRow(1);

			// Assert
			Assert.AreEqual(result.GetType(), typeof(ExcelRow));
		}

		[TestMethod]
		public void SetCellValue()
		{
			// Arrange
			// create new sheet
			var sheet = _pck.Workbook.Worksheets.Add("newsheet");

			// Act
			var result = sheet.SetCellValue(1, 1, TestValue);

			// Assert
			Assert.AreEqual(result.GetType(), typeof(ExcelRange));
			Assert.AreEqual(result.Value, TestValue);
		}
		#endregion
	}
}
