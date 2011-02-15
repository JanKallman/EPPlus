using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace ExcelPackageTest
{
	[TestClass]
	public class WorksheetsTests
	{
		private ExcelPackage package;
		private ExcelWorkbook workbook;

		[TestInitialize]
		public void TestInitialize()
		{
			package = new ExcelPackage();
			workbook = package.Workbook;
			workbook.Worksheets.Add("NEW1");
		}

		[TestMethod]
		public void ConfirmFileStructure()
		{
			Assert.IsNotNull(package, "Package not created");
			Assert.IsNotNull(workbook, "No workbook found");
		}

		[TestMethod]
		public void ShouldBeAbleToDeleteAndThenAdd()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Delete(1);
			workbook.Worksheets.Add("NEW3");
		}

		[TestMethod]
		public void DeleteByNameWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Delete("NEW2");
		}

		[TestMethod, ExpectedException(typeof(Exception))]
		public void DeleteByNameWhereWorkSheetDoesNotExist()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Delete("NEW3");
		}
	}
}
