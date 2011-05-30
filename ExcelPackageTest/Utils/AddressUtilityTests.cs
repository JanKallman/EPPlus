using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Utils;
using OfficeOpenXml;

namespace EPPlusTest.Utils
{
    [TestClass]
    public class AddressUtilityTests
    {
        [TestMethod]
        public void ParseForEntireColumnSelections_ShouldAddMaxRows()
        {
            // Arrange
            var address = "A:A";

            // Act
            var result = AddressUtility.ParseEntireColumnSelections(address);

            // Assert
            Assert.AreEqual("A1:A" + ExcelPackage.MaxRows, result);
        }

        [TestMethod]
        public void ParseForEntireColumnSelections_ShouldAddMaxRowsOnColumnsWithMultipleLetters()
        {
            // Arrange
            var address = "AB:AC";

            // Act
            var result = AddressUtility.ParseEntireColumnSelections(address);

            // Assert
            Assert.AreEqual("AB1:AC" + ExcelPackage.MaxRows, result);
        }

        [TestMethod]
        public void ParseForEntireColumnSelections_ShouldHandleMultipleRanges()
        {
            // Arrange
            var address = "A:A B:B";
            var expected = string.Format("A1:A{0} B1:B{0}", ExcelPackage.MaxRows);

            // Act
            var result = AddressUtility.ParseEntireColumnSelections(address);

            // Assert
            Assert.AreEqual(expected, result);
        }
    }
}
