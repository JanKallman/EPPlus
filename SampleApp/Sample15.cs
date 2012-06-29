/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 *  Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman                      Added       		        2012-05-01
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Security.Cryptography.X509Certificates;
using System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlusSamples
{
    class Sample15
    {
        public static void VBASample(DirectoryInfo outputDir)
        {
            //Create a macro-enabled workbook from scratch.
            VBASample1(outputDir);
            
            //Open Sample 1 and add code to change the chart to a bubble chart.
            VBASample2(outputDir);

            //Simple battleships game from scratch.
            VBASample3(outputDir);
        }

        private static void VBASample1(DirectoryInfo outputDir)
        {
            ExcelPackage pck = new ExcelPackage();

            //Add a worksheet.
            var ws=pck.Workbook.Worksheets.Add("VBA Sample");
            ws.Drawings.AddShape("VBASampleRect", eShapeStyle.RoundRect);
            
            //Create a vba project             
            pck.Workbook.CreateVBAProject();

            //Now add some code that creates a bubble chart...
            var sb = new StringBuilder();

            sb.AppendLine("Private Sub Workbook_Open()");
            sb.AppendLine("    [VBA Sample].Shapes(\"VBASampleRect\").TextEffect.Text = \"This text is set from VBA!\"");
            sb.AppendLine("End Sub");
            pck.Workbook.CodeModule.Code = sb.ToString();            

            //And Save as xlsm
            pck.SaveAs(new FileInfo(outputDir.FullName + @"\sample15-1.xlsm"));
        }

        private static void VBASample2(DirectoryInfo outputDir)
        {
            //Open Sample 1 again
            ExcelPackage pck = new ExcelPackage(new FileInfo(outputDir.FullName + @"\sample1.xlsx"));
            //Create a vba project             
            pck.Workbook.CreateVBAProject();

            //Now add some code that creates a bubble chart...
            var sb = new StringBuilder();

            sb.AppendLine("Public Sub CreateBubbleChart()");
            sb.AppendLine("Dim co As ChartObject");
            sb.AppendLine("Set co = Inventory.ChartObjects.Add(10, 100, 400, 200)");
            sb.AppendLine("co.Chart.SetSourceData Source:=Range(\"'Inventory'!$B$1:$E$5\")");
            sb.AppendLine("co.Chart.ChartType = xlBubble3DEffect         'Type currently not supported by EPPlus");
            sb.AppendLine("End Sub");

            //Create a new module and set the code
            var module = pck.Workbook.VbaProject.Modules.AddModule("EPPlusGeneratedCode");
            module.Code = sb.ToString();

            //Call the newly created sub from the workbook open event
            pck.Workbook.CodeModule.Code = "Private Sub Workbook_Open()\r\nCreateBubbleChart\r\nEnd Sub";

            //Optionally, Sign the code with your company certificate.
            /*            
            X509Store store = new X509Store(StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            pck.Workbook.VbaProject.Signature.Certificate = store.Certificates[0];
            */

            //And Save as xlsm
            pck.SaveAs(new FileInfo(outputDir.FullName + @"\sample15-2.xlsm"));
        }

        private static void VBASample3(DirectoryInfo outputDir)
        {
            //Now, lets do something a little bit more fun.
            //We are going to create a simple battleships game from scratch.

            ExcelPackage pck = new ExcelPackage();

            //Add a worksheet.
            var ws = pck.Workbook.Worksheets.Add("Battleship");

            ws.View.ShowGridLines = false;
            ws.View.ShowHeaders = false;

            ws.DefaultColWidth = 3;
            ws.DefaultRowHeight = 15;

            int gridSize=10;

            //Create the boards
            var board1 = ws.Cells[2, 2, 2 + gridSize - 1, 2 + gridSize - 1];
            var board2 = ws.Cells[2, 4+gridSize-1, 2 + gridSize-1, 4 + (gridSize-1)*2];
            CreateBoard(board1);
            CreateBoard(board2);

            ws.Select("B2");
            ws.Protection.IsProtected = true;
            ws.Protection.AllowSelectLockedCells = true;

            //Create the VBA Project
            pck.Workbook.CreateVBAProject();
            //Password protect your code
            pck.Workbook.VbaProject.Protection.SetPassword("EPPlus");

            //Add all the code from the textfiles in the Vba-Code sub-folder.
            pck.Workbook.CodeModule.Code = File.ReadAllText("..\\..\\VBA-Code\\ThisWorkbook.txt");
            
            //Add the sheet code
            ws.CodeModule.Code = File.ReadAllText("..\\..\\VBA-Code\\BattleshipSheet.txt");
            var m1=pck.Workbook.VbaProject.Modules.AddModule("Code");
            string code = File.ReadAllText("..\\..\\VBA-Code\\CodeModule.txt");
            
            //Insert your ships on the right board. you can changes these, but don't cheat ;)
            var ships = new string[]{
                "N3:N7",
                "P2:S2",
                "V9:V11",
                "O10:Q10",
                "R11:S11"};

            code = string.Format(code, ships[0],ships[1],ships[2],ships[3],ships[4], board1.Address, board2.Address);  //Ships are injected into the constants in the module
            m1.Code = code;

            //Ships are displayed with a black background
            string shipsaddress = string.Join(",", ships);
            ws.Cells[shipsaddress].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells[shipsaddress].Style.Fill.BackgroundColor.SetColor(Color.Black);

            var m2 = pck.Workbook.VbaProject.Modules.AddModule("ComputerPlay");
            m2.Code = File.ReadAllText("..\\..\\VBA-Code\\ComputerPlayModule.txt");

            var c1 = pck.Workbook.VbaProject.Modules.AddClass("Ship",false);
            c1.Code = File.ReadAllText("..\\..\\VBA-Code\\ShipClass.txt");

            //Add the info text shape.
            var tb = ws.Drawings.AddShape("txtInfo", eShapeStyle.Rect);
            tb.SetPosition(1, 0, 27, 0);
            tb.Fill.Color = Color.LightSlateGray;
            var rt1 = tb.RichText.Add("Battleships");
            rt1.Bold = true;
            tb.RichText.Add("\r\nDouble-click on the left board to make your move. Find and sink all ships to Win!");

            //Set the headers.
            ws.SetValue("B1", "Computer Grid");
            ws.SetValue("M1", "Your Grid");
            ws.Row(1).Style.Font.Size = 18;

            AddChart(ws.Cells["B13"], "chtHitPercent", "Player");
            AddChart(ws.Cells["M13"], "chtComputerHitPercent", "Computer");

            ws.Names.Add("LogStart", ws.Cells["B24"]);
            ws.Cells["B24:X224"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["B25:X224"].Style.Font.Name = "Consolas";
            ws.SetValue("B24", "Log");
            ws.Cells["B24"].Style.Font.Bold = true;
            ws.Cells["B24:X24"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            
            //If you have a valid certificate for code signing you can use this code to set it.
            ///*** Try to find a cert valid for signing... ***/
            //X509Store store = new X509Store(StoreLocation.CurrentUser);
            //store.Open(OpenFlags.ReadOnly);
            //foreach (var cert in store.Certificates)
            //{
            //    if (cert.HasPrivateKey && cert.NotBefore <= DateTime.Today && cert.NotAfter >= DateTime.Today)
            //    {
            //        pck.Workbook.VbaProject.Signature.Certificate = cert;
            //        break;
            //    }
            //}

            pck.SaveAs(new FileInfo(outputDir.FullName + @"\sample15-3.xlsm"));
        }

        private static void AddChart(ExcelRange rng,string name, string prefix)
        {
            var chrt = (ExcelPieChart)rng.Worksheet.Drawings.AddChart(name, eChartType.Pie);
            chrt.SetPosition(rng.Start.Row-1, 0, rng.Start.Column-1, 0);
            chrt.To.Row = rng.Start.Row+9;
            chrt.To.Column = rng.Start.Column + 9;
            chrt.Style = eChartStyle.Style18;
            chrt.DataLabel.ShowPercent = true;

            var serie = chrt.Series.Add(rng.Offset(2, 2, 1, 2), rng.Offset(1, 2, 1, 2));
            serie.Header = "Hits";
            
            chrt.Title.Text = "Hit ratio";
            
            var n1 = rng.Worksheet.Names.Add(prefix + "Misses", rng.Offset(2, 2));
            n1.Value = 0;
            var n2 = rng.Worksheet.Names.Add(prefix + "Hits", rng.Offset(2, 3));
            n2.Value = 0;
            rng.Offset(1, 2).Value = "Misses";
            rng.Offset(1, 3).Value = "Hits";            
        }

        private static void CreateBoard(ExcelRange rng)
        {
            rng.Style.Fill.Gradient.Color1.SetColor(Color.FromArgb(0x80, 0x80, 0XFF));
            rng.Style.Fill.Gradient.Color2.SetColor(Color.FromArgb(0x20, 0x20, 0XFF));
            rng.Style.Fill.Gradient.Type = ExcelFillGradientType.None;
            for (int col = 0; col <= rng.End.Column - rng.Start.Column; col++)
            {
                for (int row = 0; row <= rng.End.Row - rng.Start.Row; row++)
                {
                    if (col % 4 == 0)
                    {
                        rng.Offset(row, col, 1, 1).Style.Fill.Gradient.Degree = 45;
                    }
                    if (col % 4 == 1)
                    {
                        rng.Offset(row, col, 1, 1).Style.Fill.Gradient.Degree = 70;
                    }
                    if (col % 4 == 2)
                    {
                        rng.Offset(row, col, 1, 1).Style.Fill.Gradient.Degree = 110;
                    }
                    else
                    {
                        rng.Offset(row, col, 1, 1).Style.Fill.Gradient.Degree = 135;
                    }
                }
            }
            rng.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            rng.Style.Border.Top.Color.SetColor(Color.Gray);
            rng.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            rng.Style.Border.Right.Color.SetColor(Color.Gray);
            rng.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            rng.Style.Border.Left.Color.SetColor(Color.Gray);
            rng.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            rng.Style.Border.Bottom.Color.SetColor(Color.Gray);

            rng.Offset(0, 0, 1, rng.End.Column - rng.Start.Column+1).Style.Border.Top.Color.SetColor(Color.Black);
            rng.Offset(0, 0, 1, rng.End.Column - rng.Start.Column + 1).Style.Border.Top.Style=ExcelBorderStyle.Medium;
            int rows=rng.End.Row - rng.Start.Row;
            rng.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        }
    }
}
