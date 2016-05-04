using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
    [TestClass]
    public class WildCardValueMatcherTests
    {
        private WildCardValueMatcher _matcher;

        [TestInitialize]
        public void Setup()
        {
            _matcher = new WildCardValueMatcher();
        }

        [TestMethod]
        public void IsMatchShouldReturn0WhenSingleCharWildCardMatches()
        {
            var string1 = "a?c?";
            var string2 = "abcd";
            var result = _matcher.IsMatch(string1, string2);
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void IsMatchShouldReturn0WhenMultipleCharWildCardMatches()
        {
            var string1 = "a*c.";
            var string2 = "abcc.";
            var result = _matcher.IsMatch(string1, string2);
            Assert.AreEqual(0, result);
        }
    }
}
