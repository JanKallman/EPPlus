using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using Rhino.Mocks;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Database
{
    [TestClass]
    public class RowMatcherTests
    {
        private ExcelDatabaseCriteria GetCriteria(Dictionary<ExcelDatabaseCriteriaField, object> items)
        {
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            var criteria = MockRepository.GenerateStub<ExcelDatabaseCriteria>(provider, string.Empty);
            
            criteria.Stub(x => x.Items).Return(items);
            return criteria;
        }
        [TestMethod]
        public void IsMatchShouldReturnTrueIfCriteriasMatch()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = 1;
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField("Crit1")] = 1;
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.IsTrue(matcher.IsMatch(data, criteria));
        }

        [TestMethod]
        public void IsMatchShouldReturnFalseIfCriteriasDoesNotMatch()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = 1;
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField("Crit1")] = 1;
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 4;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.IsFalse(matcher.IsMatch(data, criteria));
        }

        [TestMethod]
        public void IsMatchShouldMatchStrings1()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = "1";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField("Crit1")] = "1";
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.IsTrue(matcher.IsMatch(data, criteria));
        }

        [TestMethod]
        public void IsMatchShouldMatchStrings2()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = "2";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField("Crit1")] = "1";
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.IsFalse(matcher.IsMatch(data, criteria));
        }

        [TestMethod]
        public void IsMatchShouldMatchWildcardStrings()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = "test";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField("Crit1")] = "t*t";
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.IsTrue(matcher.IsMatch(data, criteria));
        }

        [TestMethod]
        public void IsMatchShouldMatchNumericExpression()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = "test";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField("Crit2")] = "<3";
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.IsTrue(matcher.IsMatch(data, criteria));
        }

        [TestMethod]
        public void IsMatchShouldHandleFieldIndex()
        {
            var data = new ExcelDatabaseRow();
            data["Crit1"] = "test";
            data["Crit2"] = 2;
            data["Crit3"] = 3;

            var crit = new Dictionary<ExcelDatabaseCriteriaField, object>();
            crit[new ExcelDatabaseCriteriaField(2)] = "<3";
            crit[new ExcelDatabaseCriteriaField("Crit3")] = 3;

            var matcher = new RowMatcher();

            var criteria = GetCriteria(crit);

            Assert.IsTrue(matcher.IsMatch(data, criteria));
        }
    }
}
