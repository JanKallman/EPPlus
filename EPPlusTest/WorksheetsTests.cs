using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Reflection;

namespace EPPlusTest
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

		[TestMethod, ExpectedException(typeof(ArgumentException))]
		public void DeleteByNameWhereWorkSheetDoesNotExist()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Delete("NEW3");
		}

		[TestMethod]
		public void MoveBeforeByNameWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveBefore("NEW4", "NEW2");

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		[TestMethod]
		public void MoveAfterByNameWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveAfter("NEW4", "NEW2");

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		[TestMethod]
		public void MoveBeforeByPositionWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveBefore(4, 2);

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		[TestMethod]
		public void MoveAfterByPositionWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveAfter(4, 2);

			CompareOrderOfWorksheetsAfterSaving(package);
		}
        #region Delete Column with Save Tests

        private const string OutputDirectory = @"d:\temp\";

        [TestMethod,Ignore]
        public void DeleteFirstColumnInRangeColumnShouldBeDeleted()
        {
            // Arrange
            ExcelPackage pck = new ExcelPackage();
            using (
                Stream file =
                    Assembly.GetExecutingAssembly()
                        .GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
            {
                pck.Load(file);
            }
            var wsData = pck.Workbook.Worksheets[1];

            // Act
            wsData.DeleteColumn(1);
            pck.SaveAs(new FileInfo(OutputDirectory + "AfterDeleteColumn.xlsx"));

            // Assert
            Assert.AreEqual("Title", wsData.Cells["A1"].Text);
            Assert.AreEqual("First Name", wsData.Cells["B1"].Text);
            Assert.AreEqual("Family Name", wsData.Cells["C1"].Text);
        }


        [TestMethod, Ignore]
        public void DeleteLastColumnInRangeColumnShouldBeDeleted()
        {
            // Arrange
            ExcelPackage pck = new ExcelPackage();
            using (
                Stream file =
                    Assembly.GetExecutingAssembly()
                        .GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
            {
                pck.Load(file);
            }
            var wsData = pck.Workbook.Worksheets[1];

            // Act
            wsData.DeleteColumn(4);
            pck.SaveAs(new FileInfo(OutputDirectory + "AfterDeleteColumn.xlsx"));

            // Assert
            Assert.AreEqual("Id", wsData.Cells["A1"].Text);
            Assert.AreEqual("Title", wsData.Cells["B1"].Text);
            Assert.AreEqual("First Name", wsData.Cells["C1"].Text);
        }

        [TestMethod, Ignore]
        public void DeleteColumnAfterNormalRangeSheetShouldRemainUnchanged()
        {
            // Arrange
            ExcelPackage pck = new ExcelPackage();
            using (
                Stream file =
                    Assembly.GetExecutingAssembly()
                        .GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
            {
                pck.Load(file);
            }
            var wsData = pck.Workbook.Worksheets[1];

            // Act
            wsData.DeleteColumn(5);
            pck.SaveAs(new FileInfo(OutputDirectory + "AfterDeleteColumn.xlsx"));

            // Assert
            Assert.AreEqual("Id", wsData.Cells["A1"].Text);
            Assert.AreEqual("Title", wsData.Cells["B1"].Text);
            Assert.AreEqual("First Name", wsData.Cells["C1"].Text);
            Assert.AreEqual("Family Name", wsData.Cells["D1"].Text);

        }

        [TestMethod, Ignore]
        [ExpectedException(typeof(ArgumentException))]
        public void DeleteColumnBeforeRangeMimitThrowsArgumentException()
        {
            // Arrange
            ExcelPackage pck = new ExcelPackage();
            using (
                Stream file =
                    Assembly.GetExecutingAssembly()
                        .GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
            {
                pck.Load(file);
            }
            var wsData = pck.Workbook.Worksheets[1];

            // Act
            wsData.DeleteColumn(0);

            // Assert
            Assert.Fail();

        }

        [TestMethod, Ignore]
        [ExpectedException(typeof(ArgumentException))]
        public void DeleteColumnAfterRangeLimitThrowsArgumentException()
        {
            // Arrange
            ExcelPackage pck = new ExcelPackage();
            using (
                Stream file =
                    Assembly.GetExecutingAssembly()
                        .GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
            {
                pck.Load(file);
            }
            var wsData = pck.Workbook.Worksheets[1];

            // Act
            wsData.DeleteColumn(16385);

            // Assert
            Assert.Fail();

        }

        [TestMethod, Ignore]
        public void DeleteFirstTwoColumnsFromRangeColumnsShouldBeDeleted()
        {
            // Arrange
            ExcelPackage pck = new ExcelPackage();
            using (
                Stream file =
                    Assembly.GetExecutingAssembly()
                        .GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
            {
                pck.Load(file);
            }
            var wsData = pck.Workbook.Worksheets[1];

            // Act
            wsData.DeleteColumn(1, 2);
            pck.SaveAs(new FileInfo(OutputDirectory + "AfterDeleteColumn.xlsx"));

            // Assert
            Assert.AreEqual("First Name", wsData.Cells["A1"].Text);
            Assert.AreEqual("Family Name", wsData.Cells["B1"].Text);

        }
        #endregion

        [TestMethod]
        public void RangeClearMethodShouldNotClearSurroundingCells()
        {
            var wks  = workbook.Worksheets.Add("test");
            wks.Cells[2, 2].Value = "something";
            wks.Cells[2, 3].Value = "something";

            wks.Cells[2, 3].Clear();

            Assert.IsNotNull(wks.Cells[2, 2].Value);
            Assert.AreEqual("something", wks.Cells[2, 2].Value);
            Assert.IsNull(wks.Cells[2, 3].Value);
        }

        private static void CompareOrderOfWorksheetsAfterSaving(ExcelPackage editedPackage)
		{
			var packageStream = new MemoryStream();
			editedPackage.SaveAs(packageStream);

			var newPackage = new ExcelPackage(packageStream);
			var positionId = 1;
			foreach (var worksheet in editedPackage.Workbook.Worksheets)
			{
				Assert.AreEqual(worksheet.Name, newPackage.Workbook.Worksheets[positionId].Name, "Worksheets are not in the same order");
				positionId++;
			}
		}
	}
}
