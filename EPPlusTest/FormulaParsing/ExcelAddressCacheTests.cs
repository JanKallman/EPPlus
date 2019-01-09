using OfficeOpenXml.FormulaParsing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class ExcelAddressCacheTests
    {
        [TestMethod]
        public void ShouldGenerateNewIds()
        {
            var cache = new ExcelAddressCache();
            var firstId = cache.GetNewId();
            Assert.AreEqual(1, firstId);

            var secondId = cache.GetNewId();
            Assert.AreEqual(2, secondId);
        }

        [TestMethod]
        public void ShouldReturnCachedAddress()
        {
            var cache = new ExcelAddressCache();
            var id = cache.GetNewId();
            var address = "A1";
            var result = cache.Add(id, address);
            Assert.IsTrue(result);
            Assert.AreEqual(address, cache.Get(id));
        }

        [TestMethod]
        public void AddShouldReturnFalseIfUsedId()
        {
            var cache = new ExcelAddressCache();
            var id = cache.GetNewId();
            var address = "A1";
            var result = cache.Add(id, address);
            Assert.IsTrue(result);
            var result2 = cache.Add(id, address);
            Assert.IsFalse(result2);
        }

        [TestMethod]
        public void ClearShouldResetId()
        {
            var cache = new ExcelAddressCache();
            var id = cache.GetNewId();
            Assert.AreEqual(1, id);
            var address = "A1";
            var result = cache.Add(id, address);
            Assert.AreEqual(1, cache.Count);
            var id2 = cache.GetNewId();
            Assert.AreEqual(2, id2);
            cache.Clear();
            var id3 = cache.GetNewId();
            Assert.AreEqual(1, id3);
            
        }
    }
}
