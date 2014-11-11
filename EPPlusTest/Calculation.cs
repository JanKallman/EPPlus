using System;
using System.Globalization;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using System.Diagnostics;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest
{
    [DeploymentItem("Workbooks", "targetFolder")]
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
        public void CalulationTestDatatypes()
        {
            var pck = new ExcelPackage();
            var ws=pck.Workbook.Worksheets.Add("Calc1");
            ws.SetValue("A1", (short)1);
            ws.SetValue("A2", (long)2);
            ws.SetValue("A3", (Single)3);
            ws.SetValue("A4", (double)4);
            ws.SetValue("A5", (Decimal)5);
            ws.SetValue("A6", (byte)6);
            ws.SetValue("A7", null);
            ws.Cells["A10"].Formula = "Sum(A1:A8)";
            ws.Cells["A11"].Formula = "SubTotal(9,A1:A8)";
            ws.Cells["A12"].Formula = "Average(A1:A8)";

            ws.Calculate();
            Assert.AreEqual(21D, ws.Cells["a10"].Value);
            Assert.AreEqual(21D, ws.Cells["a11"].Value);
            Assert.AreEqual(21D/6, ws.Cells["a12"].Value);
        }
        [TestMethod]
        public void CalculateTest()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");

            ws.SetValue("A1",( short)1);
            var v=ws.Calculate("2.5-A1+ABS(-3.0)-SIN(3)");
            Assert.AreEqual(4.3589, Math.Round((double)v, 4));
                        
            ws.Row(1).Hidden = true;
            v = ws.Calculate("subtotal(109,a1:a10)");
            Assert.AreEqual(0D, v);

            v = ws.Calculate("-subtotal(9,a1:a3)");
            Assert.AreEqual(-1D, v);
        }
        [TestMethod]
        public void CalculateTestIsFunctions()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");

            ws.SetValue(1, 1, 1.0D);
            ws.SetFormula(1, 2, "isblank(A1:A5)");
            ws.SetFormula(1, 3, "concatenate(a1,a2,a3)");
            ws.SetFormula(1, 4, "Row()");
            ws.SetFormula(1, 5, "Row(a3)");
            ws.Calculate();
        }
        [Ignore]
        [TestMethod]
        public void Calulation4()
        {
            //C:\Development\epplus formulas\EPPlusTest\Workbooks\FormulaTest.xlsx
            var pck = new ExcelPackage(new FileInfo(@"C:\Development\epplus formulas\EPPlusTest\Workbooks\FormulaTest.xlsx"));
            //var pck = new ExcelPackage(new FileInfo(AppDomain.CurrentDomain.BaseDirectory + "..\\..\\..\\..\\EPPlusTest\\workbooks\\FormulaTest.xlsx"));
            pck.Workbook.Calculate();
            Assert.AreEqual(490D, pck.Workbook.Worksheets[1].Cells["D5"].Value);
        }
        [Ignore]
        [TestMethod]
        public void CalulationValidationExcel()
        {
            var dir = AppDomain.CurrentDomain.BaseDirectory;
            var pck = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "FormulaTest.xlsx")));

            var ws = pck.Workbook.Worksheets["ValidateFormulas"];
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
                catch
                {
                    errors.Add(new Tuple<string, object, object>(adr, fr[adr], ws.Cells[adr].Value));
                    nErrors++;
                }
            }
        }
        [Ignore]
        [TestMethod]
        public void TestOneCell()
        {
            var pck = new ExcelPackage(new FileInfo(@"C:\temp\EPPlusTestark\Test4.xlsm"));
            var ws = pck.Workbook.Worksheets.First(); 
            pck.Workbook.Worksheets["Räntebärande formaterat utland"].Cells["M13"].Calculate();
            Assert.AreEqual(0d, pck.Workbook.Worksheets["Räntebärande formaterat utland"].Cells["M13"].Value);  
        }

        [Ignore]
        [TestMethod]
        public void TestPrecedence()
        {
            var pck = new ExcelPackage(new FileInfo(@"C:\temp\EPPlusTestark\Precedence.xlsx"));
            var ws = pck.Workbook.Worksheets.Last();
            pck.Workbook.Calculate();
            Assert.AreEqual(150d, ws.Cells["A1"].Value);
        }
        [Ignore]
        [TestMethod]
        public void TestDataType()
        {
            var pck = new ExcelPackage(new FileInfo(@"c:\temp\EPPlusTestark\calc_amount.xlsx"));
            var ws = pck.Workbook.Worksheets[1];
            //ws.Names.Add("Name1",ws.Cells["A1"]);
            //ws.Names.Add("Name2", ws.Cells["A2"]);
            ws.Names["PRICE"].Value = 30;
            ws.Names["QUANTITY"].Value = 10;

            ws.Calculate();

            ws.Names["PRICE"].Value = 40;
            ws.Names["QUANTITY"].Value = 20;

            ws.Calculate();
        }
        [TestMethod]
        public void CalcTwiceError()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("CalcTest");
            ws.Names.AddValue("PRICE", 10);
            ws.Names.AddValue("QUANTITY", 11);
            ws.Cells["A1"].Formula="PRICE*QUANTITY";
            ws.Names.AddFormula("AMOUNT", "PRICE*QUANTITY");

            ws.Names["PRICE"].Value = 30;
            ws.Names["QUANTITY"].Value = 10;

            ws.Calculate();
            Assert.AreEqual(300D, ws.Cells["A1"].Value);
            Assert.AreEqual(300D, ws.Names["AMOUNT"].Value);
            ws.Names["PRICE"].Value = 40;
            ws.Names["QUANTITY"].Value = 20;

            ws.Calculate();
            Assert.AreEqual(800D, ws.Cells["A1"].Value);
            Assert.AreEqual(800D, ws.Names["AMOUNT"].Value);
        }
        [TestMethod]
        public void IfError()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("CalcTest");
            ws.Cells["A1"].Value = "test1";
            ws.Cells["A5"].Value = "test2";
            ws.Cells["A2"].Value = "Sant";
            ws.Cells["A3"].Value = "Falskt";
            ws.Cells["A4"].Formula = "if(A1>=A5,true,A3)";
            ws.Cells["B1"].Formula = "isText(a1)";
            ws.Cells["B2"].Formula = "isText(\"Test\")";
            ws.Cells["B3"].Formula = "isText(1)";
            ws.Cells["B4"].Formula = "isText(true)";
            ws.Cells["c1"].Formula = "mid(a1,4,15)";

            ws.Calculate();
        }
        [Ignore]
        [TestMethod]
        public void TestAllWorkbooks()
        {
            StringBuilder sb=new StringBuilder();
            //Add sheets to test in this directory or change it to your testpath.
            string path = @"C:\temp\EPPlusTestark\";
            if(!Directory.Exists(path)) return;

            foreach (var file in Directory.GetFiles(path, "*.xls*"))
            {
                sb.Append(GetOutput(file));
            }

            if (sb.Length > 0)
            {
                File.WriteAllText(string.Format("TestAllWorkooks{0}.txt", DateTime.Now.ToShortDateString() + " " + DateTime.Now.ToShortDateString()), sb.ToString());
                throw(new Exception("Test failed with\r\n\r\n" + sb.ToString()));

            }
        }        
        private string GetOutput(string file)
        {
            using (var pck = new ExcelPackage(new FileInfo(file)))
            {
                var fr = new Dictionary<string, object>();
                foreach (var ws in pck.Workbook.Worksheets)
                {
                    if (!(ws is ExcelChartsheet))
                    {
                        foreach (var cell in ws.Cells)
                        {
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                fr.Add(ws.PositionID.ToString() + "," + cell.Address, cell.Value);
                                ws._values.SetValue(cell.Start.Row, cell.Start.Column, null);
                            }
                        }
                    }
                }

                pck.Workbook.Calculate();
                var nErrors = 0;
                var errors = new List<Tuple<string, object, object>>();
                ExcelWorksheet sheet=null;
                string adr="";
                var fileErr = new System.IO.StreamWriter("c:\\temp\\err.txt");
                foreach (var cell in fr.Keys)
                {
                    try
                    {
                        var spl = cell.Split(',');
                        var ix = int.Parse(spl[0]);
                        sheet = pck.Workbook.Worksheets[ix];
                        adr = spl[1];
                        if (fr[cell] is double && (sheet.Cells[adr].Value is double || sheet.Cells[adr].Value is decimal  || sheet.Cells[adr].Value.GetType().IsPrimitive))
                        {
                            var d1 = Convert.ToDouble(fr[cell]);
                            var d2 = Convert.ToDouble(sheet.Cells[adr].Value);
                            //if (Math.Abs(d1 - d2) < double.Epsilon)
                            if(double.Equals(d1,d2))
                            {
                                continue;
                            }
                            else
                            {
                                //errors.Add(new Tuple<string, object, object>(adr, fr[cell], sheet.Cells[adr].Value));
                                fileErr.WriteLine("Diff cell " + sheet.Name + "!" + adr +"\t" + d1.ToString("R15") + "\t" + d2.ToString("R15"));
                            }
                        }
                        else
                        {
                            if ((fr[cell]??"").ToString() != (sheet.Cells[adr].Value??"").ToString())
                            {
                                fileErr.WriteLine("String?  cell " + sheet.Name + "!" + adr + "\t" + (fr[cell] ?? "").ToString() + "\t" + (sheet.Cells[adr].Value??"").ToString());
                            }
                            //errors.Add(new Tuple<string, object, object>(adr, fr[cell], sheet.Cells[adr].Value));
                        }
                    }
                    catch (Exception e)
                    {                        
                        fileErr.WriteLine("Exception cell " + sheet.Name + "!" + adr + "\t" + fr[cell].ToString() + "\t" + sheet.Cells[adr].Value +  "\t" + e.Message);
                        fileErr.WriteLine("***************************");
                        fileErr.WriteLine(e.ToString());
                        nErrors++;
                    }
                }
                fileErr.Close();
                return nErrors.ToString();
            }
        }
    }
}
