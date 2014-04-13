using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class ExcelAddressInfoTests
    {
        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void ParseShouldThrowIfAddressIsNull()
        {
            ExcelAddressInfo.Parse(null);
        }

        [TestMethod]
        public void ParseShouldSetWorksheet()
        {
            var info = ExcelAddressInfo.Parse("Worksheet!A1");
            Assert.AreEqual("Worksheet", info.Worksheet);
        }

        [TestMethod]
        public void WorksheetIsSpecifiedShouldBeTrueWhenWorksheetIsSupplied()
        {
            var info = ExcelAddressInfo.Parse("Worksheet!A1");
            Assert.IsTrue(info.WorksheetIsSpecified);
        }

        [TestMethod]
        public void ShouldIndicateMultipleCellsWhenAddressContainsAColon()
        {
            var info = ExcelAddressInfo.Parse("A1:A2");
            Assert.IsTrue(info.IsMultipleCells);
        }

        [TestMethod]
        public void ShouldSetStartCell()
        {
            var info = ExcelAddressInfo.Parse("A1:A2");
            Assert.AreEqual("A1", info.StartCell);
        }

        [TestMethod]
        public void ShouldSetEndCell()
        {
            var info = ExcelAddressInfo.Parse("A1:A2");
            Assert.AreEqual("A2", info.EndCell);
        }

        [TestMethod]
        public void ParseShouldSetAddressOnSheet()
        {
            var info = ExcelAddressInfo.Parse("Worksheet!A1:A2");
            Assert.AreEqual("A1:A2", info.AddressOnSheet);
        }

        [TestMethod]
        public void AddressOnSheetShouldBeSameAsAddressIfNoWorksheetIsSpecified()
        {
            var info = ExcelAddressInfo.Parse("A1:A2");
            Assert.AreEqual("A1:A2", info.AddressOnSheet);
        }
    }
}
