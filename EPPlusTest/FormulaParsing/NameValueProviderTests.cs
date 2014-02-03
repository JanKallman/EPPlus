using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing;
using Rhino.Mocks;

namespace EPPlusTest.FormulaParsing
{
    [TestClass]
    public class NameValueProviderTests
    {
        //private ExcelDataProvider _excelDataProvider;

        //[TestInitialize]
        //public void Setup()
        //{
        //    _excelDataProvider = MockRepository.GenerateMock<ExcelDataProvider>();
        //}

        //[TestMethod]
        //public void IsNamedValueShouldReturnTrueIfKeyIsANamedValue()
        //{
        //    var dict = new Dictionary<string, object>();
        //    dict.Add("A", "B");
        //    _excelDataProvider.Stub(x => x.GetWorkbookNameValues())
        //        .Return(dict);
        //    var nameValueProvider = new EpplusNameValueProvider(_excelDataProvider);

        //    var result = nameValueProvider.IsNamedValue("A");
        //    Assert.IsTrue(result);
        //}

        //[TestMethod]
        //public void IsNamedValueShouldReturnFalseIfKeyIsNotANamedValue()
        //{
        //    var dict = new Dictionary<string, object>();
        //    dict.Add("A", "B");
        //    _excelDataProvider.Stub(x => x.GetWorkbookNameValues())
        //        .Return(dict);
        //    var nameValueProvider = new EpplusNameValueProvider(_excelDataProvider);

        //    var result = nameValueProvider.IsNamedValue("C");
        //    Assert.IsFalse(result);
        //}

        //[TestMethod]
        //public void GetNamedValueShouldReturnCorrectValueIfKeyExists()
        //{
        //    var dict = new Dictionary<string, object>();
        //    dict.Add("A", "B");
        //    _excelDataProvider.Stub(x => x.GetWorkbookNameValues())
        //        .Return(dict);
        //    var nameValueProvider = new EpplusNameValueProvider(_excelDataProvider);

        //    var result = nameValueProvider.GetNamedValue("A");
        //    Assert.AreEqual("B", result);
        //}

        //[TestMethod]
        //public void ReloadShouldReloadDataFromExcelDataProvider()
        //{
        //    var dict = new Dictionary<string, object>();
        //    dict.Add("A", "B");
        //    _excelDataProvider.Stub(x => x.GetWorkbookNameValues())
        //        .Return(dict);
        //    var nameValueProvider = new EpplusNameValueProvider(_excelDataProvider);

        //    var result = nameValueProvider.GetNamedValue("A");
        //    Assert.AreEqual("B", result);

        //    dict.Clear();
        //    nameValueProvider.Reload();
        //    Assert.IsFalse(nameValueProvider.IsNamedValue("A"));
        //}
    }
}
