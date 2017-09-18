using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Utils.CompundDocument;

namespace EPPlusTest
{
    /// <summary>
    /// Summary description for CompoundDoc
    /// </summary>
    [TestClass]
    public class CompoundDoc
    {
        public CompoundDoc()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

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
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod, Ignore]
        public void Read()
        {
           //var doc = File.ReadAllBytes(@"c:\temp\vbaProject.bin");
           var doc = File.ReadAllBytes(@"c:\temp\vba.bin");
           var cd = new CompoundDocumentFile(doc);
           var ms = new MemoryStream();
           cd.Write(ms);
           printitems(cd.RootItem);
           File.WriteAllBytes(@"c:\temp\vba.bin", ms.ToArray());
        }

        private void printitems(CompoundDocumentItem item)
        {
            File.AppendAllText(@"c:\temp\items.txt", item.Name+ "\t");            
            foreach(var c in item.Children)
            {
                printitems(c);
            }
        }

        [TestMethod, Ignore ]
        public void ReadEncLong()
        {
            var doc=File.ReadAllBytes(@"c:\temp\EncrDocRead.xlsx");
            var cd = new CompoundDocumentFile(doc);
            var ms = new MemoryStream();
            cd.Write(ms);

            File.WriteAllBytes(@"c:\temp\vba.xlsx", ms.ToArray());
        }
        [TestMethod, Ignore]
        public void ReadVba()
        {
            var p = new ExcelPackage(new FileInfo(@"c:\temp\pricecheck.xlsm"));
            var vba = p.Workbook.VbaProject;            
        }
    }
}
