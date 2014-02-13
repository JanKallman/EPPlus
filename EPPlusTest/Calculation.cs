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
        public void Calulation4()
        {
            //C:\Development\epplus formulas\EPPlusTest\Workbooks\FormulaTest.xlsx
            var pck = new ExcelPackage(new FileInfo(@"C:\Development\epplus formulas\EPPlusTest\Workbooks\FormulaTest.xlsx"));
            //var pck = new ExcelPackage(new FileInfo(AppDomain.CurrentDomain.BaseDirectory + "..\\..\\..\\..\\EPPlusTest\\workbooks\\FormulaTest.xlsx"));
            pck.Workbook.Calculate();
            Assert.AreEqual(490D, pck.Workbook.Worksheets[1].Cells["D5"].Value);
        }
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
            pck.Workbook.Worksheets.First().Cells["J966"].Calculate();
            Assert.AreEqual(15.928239987316594, ws.Cells["J966"].Value);  

        }

        [TestMethod]
        public void TestPrecedence()
        {
            var pck = new ExcelPackage(new FileInfo(@"C:\temp\EPPlusTestark\Precedence.xlsx"));
            var ws = pck.Workbook.Worksheets.Last();
            pck.Workbook.Calculate();
            Assert.AreEqual(150d, ws.Cells["A1"].Value);
        }
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
                    foreach (var cell in ws.Cells)
                    {
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            fr.Add(ws.PositionID.ToString()+","+cell.Address, cell.Value);
                            ws._values.SetValue(cell.Start.Row, cell.Start.Column, null);
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
                        nErrors++;
                    }
                }
                fileErr.Close();
                return nErrors.ToString();
            }
        }
    }
}
