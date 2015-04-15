using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Database
{
    [TestClass]
    public class RowMatcherTests
    {
        [TestMethod]
        public void IsMatchShouldReturnTrueIfCriteriasMatch()
        {
            var data = new Dictionary<string, object>();
            data["Crit1"] = 1;
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<string, object>();
            crit["Crit1"] = 1;
            crit["Crit3"] = 3;

            var matcher = new RowMatcher();

            Assert.IsTrue(matcher.IsMatch(data, crit));
        }

        [TestMethod]
        public void IsMatchShouldReturnFalseIfCriteriasDoesNotMatch()
        {
            var data = new Dictionary<string, object>();
            data["Crit1"] = 1;
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<string, object>();
            crit["Crit1"] = 1;
            crit["Crit3"] = 4;

            var matcher = new RowMatcher();

            Assert.IsFalse(matcher.IsMatch(data, crit));
        }

        [TestMethod]
        public void IsMatchShouldMatchStrings1()
        {
            var data = new Dictionary<string, object>();
            data["Crit1"] = "1";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<string, object>();
            crit["Crit1"] = "1";
            crit["Crit3"] = 3;

            var matcher = new RowMatcher();

            Assert.IsTrue(matcher.IsMatch(data, crit));
        }

        [TestMethod]
        public void IsMatchShouldMatchStrings2()
        {
            var data = new Dictionary<string, object>();
            data["Crit1"] = "2";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<string, object>();
            crit["Crit1"] = "1";
            crit["Crit3"] = 3;

            var matcher = new RowMatcher();

            Assert.IsFalse(matcher.IsMatch(data, crit));
        }

        [TestMethod]
        public void IsMatchShouldMatchWildcardStrings()
        {
            var data = new Dictionary<string, object>();
            data["Crit1"] = "test";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<string, object>();
            crit["Crit1"] = "t*t";
            crit["Crit3"] = 3;

            var matcher = new RowMatcher();

            Assert.IsTrue(matcher.IsMatch(data, crit));
        }

        [TestMethod]
        public void IsMatchShouldMatchNumericExpression()
        {
            var data = new Dictionary<string, object>();
            data["Crit1"] = "test";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<string, object>();
            crit["Crit2"] = "<3";
            crit["Crit3"] = 3;

            var matcher = new RowMatcher();

            Assert.IsTrue(matcher.IsMatch(data, crit));
        }
    }
}
