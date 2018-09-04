using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class ExcelRangeBaseTest : TestBase
    {
        [TestMethod]
        public void CopyCopiesCommentsFromSingleCellRanges()
        {
            InitBase();
            var pck = new ExcelPackage();
            var ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
            var sourceExcelRange = ws1.Cells[3, 3];
            Assert.IsNull(sourceExcelRange.Comment);
            sourceExcelRange.AddComment("Testing comment 1", "test1");
            Assert.AreEqual("test1", sourceExcelRange.Comment.Author);
            Assert.AreEqual("Testing comment 1", sourceExcelRange.Comment.Text);
            var destinationExcelRange = ws1.Cells[5, 5];
            Assert.IsNull(destinationExcelRange.Comment);
            sourceExcelRange.Copy(destinationExcelRange);
            // Assert the original comment is intact.
            Assert.AreEqual("test1", sourceExcelRange.Comment.Author);
            Assert.AreEqual("Testing comment 1", sourceExcelRange.Comment.Text);
            // Assert the comment was copied.
            Assert.AreEqual("test1", destinationExcelRange.Comment.Author);
            Assert.AreEqual("Testing comment 1", destinationExcelRange.Comment.Text);
        }

        [TestMethod]
        public void CopyCopiesCommentsFromMultiCellRanges()
        {
            InitBase();
            var pck = new ExcelPackage();
            var ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
            var sourceExcelRangeC3 = ws1.Cells[3, 3];
            var sourceExcelRangeD3 = ws1.Cells[3, 4];
            var sourceExcelRangeE3 = ws1.Cells[3, 5];
            Assert.IsNull(sourceExcelRangeC3.Comment);
            Assert.IsNull(sourceExcelRangeD3.Comment);
            Assert.IsNull(sourceExcelRangeE3.Comment);
            sourceExcelRangeC3.AddComment("Testing comment 1", "test1");
            sourceExcelRangeD3.AddComment("Testing comment 2", "test1");
            sourceExcelRangeE3.AddComment("Testing comment 3", "test1");
            Assert.AreEqual("test1", sourceExcelRangeC3.Comment.Author);
            Assert.AreEqual("Testing comment 1", sourceExcelRangeC3.Comment.Text);
            Assert.AreEqual("test1", sourceExcelRangeD3.Comment.Author);
            Assert.AreEqual("Testing comment 2", sourceExcelRangeD3.Comment.Text);
            Assert.AreEqual("test1", sourceExcelRangeE3.Comment.Author);
            Assert.AreEqual("Testing comment 3", sourceExcelRangeE3.Comment.Text);
            // Copy the full row to capture each cell at once.
            Assert.IsNull(ws1.Cells[5, 3].Comment);
            Assert.IsNull(ws1.Cells[5, 4].Comment);
            Assert.IsNull(ws1.Cells[5, 5].Comment);
            ws1.Cells["3:3"].Copy(ws1.Cells["5:5"]);
            // Assert the original comments are intact.
            Assert.AreEqual("test1", sourceExcelRangeC3.Comment.Author);
            Assert.AreEqual("Testing comment 1", sourceExcelRangeC3.Comment.Text);
            Assert.AreEqual("test1", sourceExcelRangeD3.Comment.Author);
            Assert.AreEqual("Testing comment 2", sourceExcelRangeD3.Comment.Text);
            Assert.AreEqual("test1", sourceExcelRangeE3.Comment.Author);
            Assert.AreEqual("Testing comment 3", sourceExcelRangeE3.Comment.Text);
            // Assert the comments were copied.
            var destinationExcelRangeC5 = ws1.Cells[5, 3];
            var destinationExcelRangeD5 = ws1.Cells[5, 4];
            var destinationExcelRangeE5 = ws1.Cells[5, 5];
            Assert.AreEqual("test1", destinationExcelRangeC5.Comment.Author);
            Assert.AreEqual("Testing comment 1", destinationExcelRangeC5.Comment.Text);
            Assert.AreEqual("test1", destinationExcelRangeD5.Comment.Author);
            Assert.AreEqual("Testing comment 2", destinationExcelRangeD5.Comment.Text);
            Assert.AreEqual("test1", destinationExcelRangeE5.Comment.Author);
            Assert.AreEqual("Testing comment 3", destinationExcelRangeE5.Comment.Text);
        }

        [TestMethod]
        public void LoadFromCollectionPrintsMemberHeaders()
        {
            var kittens = new[]
            {
                new KittenData("Kuro", 0.5, 0.8),
                new KittenData("Mittens", 0.6, 0.9),
            };
            var pck = new ExcelPackage();
            var ws1 = pck.Workbook.Worksheets.Add("Kittens");
            var targetExcelRangeA1 = ws1.Cells[1, 1];
            targetExcelRangeA1.LoadFromCollection(kittens, PrintHeaders: true);
            // Headers
            Assert.AreEqual("Name", ws1.Cells[1, 1].Value);
            Assert.AreEqual("Furriness", ws1.Cells[1, 2].Value);
            Assert.AreEqual("Cuteness", ws1.Cells[1, 3].Value);
            Assert.IsNull(ws1.Cells[1, 4].Value);
            // First kitten
            Assert.AreEqual("Kuro", ws1.Cells[2, 1].Value);
            Assert.AreEqual(0.5, ws1.Cells[2, 2].Value);
            Assert.AreEqual(0.8, ws1.Cells[2, 3].Value);
            Assert.IsNull(ws1.Cells[2, 4].Value);
            // Second kitten
            Assert.AreEqual("Mittens", ws1.Cells[3, 1].Value);
            Assert.AreEqual(0.6, ws1.Cells[3, 2].Value);
            Assert.AreEqual(0.9, ws1.Cells[3, 3].Value);
            Assert.IsNull(ws1.Cells[3, 4].Value);
            // No third kitten
            Assert.IsNull(ws1.Cells[4, 1].Value);
        }

        [TestMethod]
        public void LoadFromCollectionSupportsEmptyCollectionWithoutHeaders()
        {
            var kittens = new KittenData[]
            {
            };
            var pck = new ExcelPackage();
            var ws1 = pck.Workbook.Worksheets.Add("Kittens");
            var targetExcelRangeA1 = ws1.Cells[1, 1];
            targetExcelRangeA1.LoadFromCollection(kittens);
            // Nothing should be written.
            foreach (var cell in ws1.Cells)
            {
                Assert.IsNull(cell.Value);
            }
        }

        [TestMethod]
        public void SettingAddressHandlesMultiAddresses()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                var name = package.Workbook.Names.Add("Test", worksheet.Cells[3, 3]);
                name.Address = "Sheet1!C3";
                name.Address = "Sheet1!D3";
                Assert.IsNull(name.Addresses);
                name.Address = "C3:D3,E3:F3";
                Assert.IsNotNull(name.Addresses);
                name.Address = "Sheet1!C3";
                Assert.IsNull(name.Addresses);
            }
        }

        class KittenData
        {
            public string Name { get; set; }
            public double Furriness { get; set; }
            public double Cuteness { get; set; }
            public KittenData()
            {
            }
            public KittenData(
               string name,
               double furriness,
               double cuteness)
            {
                Name = name;
                Furriness = furriness;
                Cuteness = cuteness;
            }
        }
    }
}