using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EPPlusTest.FormulaParsing.IntegrationTests.ErrorHandling
{
    /// <summary>
    /// Summary description for SumTests
    /// </summary>
    [TestClass]
    public class SumTests : FormulaErrorHandlingTestBase
    {
        [TestInitialize]
        public void ClassInitialize()
        {
            BaseInitialize();
        }

        [TestCleanup]
        public void ClassCleanup()
        {
            BaseCleanup();
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


        [TestMethod]
        public void SingleCell()
        {
            Assert.AreEqual(3d, Worksheet.Cells["B9"].Value);
        }

        [TestMethod]
        public void MultiCell()
        {
            Assert.AreEqual(40d, Worksheet.Cells["C9"].Value);
        }

        [TestMethod]
        public void Name()
        {
            Assert.AreEqual(10d, Worksheet.Cells["E9"].Value);
        }

        [TestMethod]
        public void ReferenceError()
        {
            Assert.AreEqual("#REF!", Worksheet.Cells["H9"].Value);
        }

        [TestMethod]
        public void NameOnOtherSheet()
        {
            Assert.AreEqual(130d, Worksheet.Cells["I9"].Value);
        }

        [TestMethod]
        public void ArrayInclText()
        {
            Assert.AreEqual(7d, Worksheet.Cells["J9"].Value);
        }

        [TestMethod]
        public void NameError()
        {
            Assert.AreEqual("#NAME?", Worksheet.Cells["L9"].Value);
        }

        [TestMethod]
        public void DivByZeroError()
        {
            Assert.AreEqual("#DIV/0!", Worksheet.Cells["M9"].Value);
        }
    }
}
