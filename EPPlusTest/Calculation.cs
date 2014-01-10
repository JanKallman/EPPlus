using System;
using System.Globalization;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Calculation;
using System.IO;
using System.Diagnostics;

namespace EPPlusTest
{
    [TestClass]
    public class Calculation
    {
        //[TestMethod]
        //public void Calulation()
        //{
        //    var pck = new ExcelPackage(new FileInfo("c:\\temp\\chain.xlsx"));
        //    pck.Workbook.Calculate();
        //    Assert.AreEqual(50D, pck.Workbook.Worksheets[1].Cells["C1"].Value);
        //}
        //[TestMethod]
        //public void Calulation2()
        //{
        //    var pck = new ExcelPackage(new FileInfo("c:\\temp\\chainTest.xlsx"));
        //    pck.Workbook.Calculate();
        //    Assert.AreEqual(1124999960382D, pck.Workbook.Worksheets[1].Cells["C1"].Value);
        //}
        //[TestMethod]
        //public void Calulation3()
        //{
        //    var pck = new ExcelPackage(new FileInfo("c:\\temp\\names.xlsx"));
        //    pck.Workbook.Calculate();
        //    //Assert.AreEqual(1124999960382D, pck.Workbook.Worksheets[1].Cells["C1"].Value);
        //}
        [TestMethod]
        public void Calulation4()
        {
            //C:\Development\epplus formulas\EPPlusTest\Workbooks\FormulaTest.xlsx
            var pck = new ExcelPackage(new FileInfo(@"C:\Development\epplus formulas\EPPlusTest\Workbooks\FormulaTest.xlsx"));
            //var pck = new ExcelPackage(new FileInfo(AppDomain.CurrentDomain.BaseDirectory + "..\\..\\..\\..\\EPPlusTest\\workbooks\\FormulaTest.xlsx"));
            pck.Workbook.Calculate();
            Assert.AreEqual(490D, pck.Workbook.Worksheets[1].Cells["D5"].Value);
        }
        [TestMethod, Ignore]
        public void CalulationValidationExcel()
        {
            //C:\Development\epplus formulas\EPPlusTest\Workbooks\FormulaTest.xlsx
            //C:\temp\EPPlus Testark\Test1.xslx
            var pck = new ExcelPackage(new FileInfo(@"C:\temp\EPPlusTestark\Test1.xlsx"));
            //var pck = new ExcelPackage(new FileInfo(@"C:\temp\Bok2.xlsx"));
            //var ws = pck.Workbook.Worksheets["ValidateFormulas"];
            var ws = pck.Workbook.Worksheets.First();
            var fr = new Dictionary<string, object>();
            foreach (var cell in ws.Cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    fr.Add(cell.Address, cell.Value);
                }
            }
            pck.Workbook.Calculate();
            var nErrors = 0;
            var errors = new List<Tuple<string, object, object>>();
            foreach (var adr in fr.Keys)
            {
                try
                {
                    if (fr[adr] is double && ws.Cells[adr].Value is double)
                    {
                        var d1 = Convert.ToDouble(fr[adr]);
                        var d2 = Convert.ToDouble(ws.Cells[adr].Value);
                        if (Math.Abs(d1 - d2) < 0.0001)
                        {
                            continue;
                        }
                        else
                        {
                            Assert.AreEqual(fr[adr], ws.Cells[adr].Value);
                        }
                    }
                    else
                    {
                        Assert.AreEqual(fr[adr], ws.Cells[adr].Value);
                    }
                }
                catch (Exception e)
                {
                    errors.Add(new Tuple<string, object, object>(adr, fr[adr], ws.Cells[adr].Value));
                    nErrors++;
                }
            }
            
        }

        [TestMethod]
        public void TestOneCell()
        {
            var pck = new ExcelPackage(new FileInfo(@"C:\temp\EPPlusTestark\Test1.xlsx"));
            var ws = pck.Workbook.Worksheets.First(); 
            pck.Workbook.Worksheets.First().Cells["Q6"].Calculate();
            Assert.AreEqual(474378, ws.Cells["Q6"].Value);  
        }

        [TestMethod]
        public void TestPrecedence()
        {
            var pck = new ExcelPackage(new FileInfo(@"C:\temp\EPPlusTestark\Precedence.xlsx"));
            var ws = pck.Workbook.Worksheets.Last();
            pck.Workbook.Calculate();
            Assert.AreEqual(150d, ws.Cells["A1"].Value);
        }
    }
}
