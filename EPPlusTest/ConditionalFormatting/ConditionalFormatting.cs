using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System.IO;

namespace EPPlusTest
{
  /// <summary>
  /// Test the Conditional Formatting feature
  /// </summary>
  [TestClass]
  public class ConditionalFormatting
  {
    private TestContext testContextInstance;
    private static ExcelPackage _pck;

    /// <summary>
    ///Gets or sets the test context which provides
    ///information about and functionality for the current test run.
    ///</summary>
    public TestContext TestContext
    {
      get
      {
        return testContextInstance;
      }
      set
      {
        testContextInstance = value;
      }
    }

    #region Additional test attributes
    // You can use the following additional attributes as you write your tests:
    // Use ClassInitialize to run code before running the first test in the class
    [ClassInitialize()]
    public static void MyClassInitialize(TestContext testContext)
    {
      if (Directory.Exists("Test"))
      {
        Directory.Delete("Test", true);
      }

      Directory.CreateDirectory(string.Format("Test"));
      _pck = new ExcelPackage(new FileInfo(@"Test\ConditionalFormatting.xlsx"));
    }

    // Use ClassCleanup to run code after all tests in a class have run
    [ClassCleanup()]
    public static void MyClassCleanup()
    {
      _pck = null;
    }

    // //Use TestInitialize to run code before running each test 
    // [TestInitialize()]
    // public void MyTestInitialize() 
    // {
    // }

    //// Use TestCleanup to run code after each test has run
    // [TestCleanup()]
    // public void MyTestCleanup() 
    // {
    // }
    #endregion

    /// <summary>
    /// 
    /// </summary>
    [TestMethod]
    public void TwoColorScale()
    {
      var ws = _pck.Workbook.Worksheets.Add("TwoColorScale");
    }

    /// <summary>
    /// 
    /// </summary>
    [TestMethod]
    public void ReadConditionalFormatting()
    {
      var pck = new ExcelPackage(new FileInfo(@"c:\temp\cf.xlsx"));

      Assert.IsTrue(pck.Workbook.Worksheets[1].ConditionalFormatting.Count > 0);
    }
  }
}