using System;
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

        #region UpdateCrossSheetReferenceNames Tests
        [TestMethod]
        public void UpdateFormulaSheetReferences()
        {
          var result = ExcelCellBase.UpdateFormulaSheetReferences("5+'OldSheet'!$G3+'Some Other Sheet'!C3+SUM(1,2,3)", "OldSheet", "NewSheet");
          Assert.AreEqual("5+'NewSheet'!$G3+'Some Other Sheet'!C3+SUM(1,2,3)", result);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void UpdateFormulaSheetReferencesNullOldSheetThrowsException()
        {
          ExcelCellBase.UpdateFormulaSheetReferences("formula", null, "sheet2");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void UpdateFormulaSheetReferencesEmptyOldSheetThrowsException()
        {
          ExcelCellBase.UpdateFormulaSheetReferences("formula", string.Empty, "sheet2");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void UpdateFormulaSheetReferencesNullNewSheetThrowsException()
        {
          ExcelCellBase.UpdateFormulaSheetReferences("formula", "sheet1", null);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void UpdateFormulaSheetReferencesEmptyNewSheetThrowsException()
        {
          ExcelCellBase.UpdateFormulaSheetReferences("formula", "sheet1", string.Empty);
        }
        #endregion
  }
}
