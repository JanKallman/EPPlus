using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class ExcelCellBaseTest
    {
        #region UpdateFormulaReferences Tests
        [TestMethod]
        public void UpdateFormulaReferencesOnTheSameSheet()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("C3", 3, 3, 2, 2, "sheet", "sheet");
            Assert.AreEqual("F6", result);
        }

        [TestMethod]
        public void UpdateFormulaReferencesIgnoresIncorrectSheet()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("C3", 3, 3, 2, 2, "sheet", "other sheet");
            Assert.AreEqual("C3", result);
        }

        [TestMethod]
        public void UpdateFormulaReferencesFullyQualifiedReferenceOnTheSameSheet()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("'sheet name here'!C3", 3, 3, 2, 2, "sheet name here", "sheet name here");
            Assert.AreEqual("'sheet name here'!F6", result);
        }

        [TestMethod]
        public void UpdateFormulaReferencesFullyQualifiedCrossSheetReferenceArray()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("SUM('sheet name here'!B2:D4)", 3, 3, 3, 3, "cross sheet", "sheet name here");
            Assert.AreEqual("SUM('sheet name here'!B2:G7)", result);
        }

        [TestMethod]
        public void UpdateFormulaReferencesFullyQualifiedReferenceOnADifferentSheet()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("'updated sheet'!C3", 3, 3, 2, 2, "boring sheet", "updated sheet");
            Assert.AreEqual("'updated sheet'!F6", result);
        }

        [TestMethod]
        public void UpdateFormulaReferencesReferencingADifferentSheetIsNotUpdated()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("'boring sheet'!C3", 3, 3, 2, 2, "boring sheet", "updated sheet");
            Assert.AreEqual("'boring sheet'!C3", result);
        }
        #endregion
    }
}
