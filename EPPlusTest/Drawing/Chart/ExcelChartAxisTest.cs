using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EPPlusTest.Drawing.Chart
{
    [TestClass]
    public class ExcelChartAxisTest
    {
        private ExcelChartAxis axis;
        
        [TestInitialize]
        public void Initialize()
        {
            var xmlDoc = new XmlDocument();
            var xmlNsm = new XmlNamespaceManager(new NameTable());
            xmlNsm.AddNamespace("c", ExcelPackage.schemaChart);
            axis = new ExcelChartAxis(xmlNsm, xmlDoc.CreateElement("axis"));
        }

        [TestMethod]
        public void CrossesAt_SetTo2_Is2()
        {
            axis.CrossesAt = 2;
            Assert.AreEqual(axis.CrossesAt, 2);
        }

        [TestMethod]
        public void CrossesAt_SetTo1EMinus6_Is1EMinus6()
        {
            axis.CrossesAt = 1.2e-6;
            Assert.AreEqual(axis.CrossesAt, 1.2e-6);
        }

        [TestMethod]
        public void MinValue_SetTo2_Is2()
        {
            axis.MinValue = 2;
            Assert.AreEqual(axis.MinValue, 2);
        }

        [TestMethod]
        public void MinValue_SetTo1EMinus6_Is1EMinus6()
        {
            axis.MinValue = 1.2e-6;
            Assert.AreEqual(axis.MinValue, 1.2e-6);
        }

        [TestMethod]
        public void MaxValue_SetTo2_Is2()
        {
            axis.MaxValue = 2;
            Assert.AreEqual(axis.MaxValue, 2);
        }

        [TestMethod]
        public void MaxValue_SetTo1EMinus6_Is1EMinus6()
        {
            axis.MaxValue = 1.2e-6;
            Assert.AreEqual(axis.MaxValue, 1.2e-6);
        }
        [TestMethod] 
        public void Gridlines_Set_IsNotNull()
        { 
            var major = axis.MajorGridlines; 
            Assert.IsTrue(axis.ExistNode("c:majorGridlines")); 
  
            var minor = axis.MinorGridlines; 
            Assert.IsTrue(axis.ExistNode("c:minorGridlines")); 
        } 
  
        [TestMethod] 
        public void Gridlines_Remove_IsNull()
        { 
            var major = axis.MajorGridlines; 
            var minor = axis.MinorGridlines; 
  
            axis.RemoveGridlines(); 
  
            Assert.IsFalse(axis.ExistNode("c:majorGridlines")); 
            Assert.IsFalse(axis.ExistNode("c:minorGridlines")); 
  
            major = axis.MajorGridlines; 
            minor = axis.MinorGridlines; 
  
            axis.RemoveGridlines(true, false); 
  
            Assert.IsFalse(axis.ExistNode("c:majorGridlines")); 
            Assert.IsTrue(axis.ExistNode("c:minorGridlines")); 
  
            major = axis.MajorGridlines; 
            minor = axis.MinorGridlines; 
  
            axis.RemoveGridlines(false, true); 
  
            Assert.IsTrue(axis.ExistNode("c:majorGridlines")); 
            Assert.IsFalse(axis.ExistNode("c:minorGridlines")); 
        } 
    }
}
