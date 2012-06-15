/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * See http://epplus.codeplex.com/ for details
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
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		23-MAR-2010
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Text;
namespace EPPlusWebSample
{
    public partial class GetSample : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            switch (Request.QueryString["Sample"])
            {
                case "1":
                    Sample1();
                    break;
                case "2":
                    Sample2();
                    break;
                case "3":
                    Sample3();
                    break;
                case "4":
                    Sample4();
                    break;
                default:
                    Response.Write("<script>javascript:alert('Invalid querystring');</script>");
                    break;

            }
        }

        /// <summary>
        /// Sample 1 
        /// Demonstrates the SaveAs method
        /// </summary>
        private void Sample1()
        {
            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Sample1");

            ws.Cells["A1"].Value = "Sample 1";
            ws.Cells["A1"].Style.Font.Bold = true;
            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(50, 200);
            shape.SetSize(200, 100);
            shape.Text = "Sample 1 saves to the Response.OutputStream";

            pck.SaveAs(Response.OutputStream);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=Sample1.xlsx");
        }
        /// <summary>
        /// Sample 2
        /// Demonstrates the GetAsByteArray method
        /// </summary>
        private void Sample2()
        {
            ExcelPackage pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Sample2");

            ws.Cells["A1"].Value = "Sample 2";
            ws.Cells["A1"].Style.Font.Bold = true;
            var shape = ws.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.SetPosition(50, 200);
            shape.SetSize(200, 100);
            shape.Text = "Sample 2 outputs the sheet using the Response.BinaryWrite method";

            Response.BinaryWrite(pck.GetAsByteArray());
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=Sample2.xlsx");
        }
        /// <summary>
        /// Sample 3
        /// Uses a cached template
        /// </summary>
        private void Sample3()
        {
            if (Application["Sample3Template"] == null) //Check if the template is loaded
            {
                //Here we create the template. 
                //As an alternative the template could be loaded from disk or from a resource.
                ExcelPackage pckTemplate = new ExcelPackage();
                var wsTemplate = pckTemplate.Workbook.Worksheets.Add("Sample3");

                wsTemplate.Cells["A1"].Value = "Sample 3";
                wsTemplate.Cells["A1"].Style.Font.Bold = true;
                var shape = wsTemplate.Drawings.AddShape("Shape1", eShapeStyle.Rect);
                shape.SetPosition(50, 200);
                shape.SetSize(200, 100);
                shape.Text = "Sample 3 uses a template that is stored in the application cashe.";
                pckTemplate.Save();

                Application["Sample3Template"] = pckTemplate.Stream;
            }

            //Open the new package with the template stream.
            //The template stream is copied to the new stream in the constructor
            ExcelPackage pck = new ExcelPackage(new MemoryStream(), Application["Sample3Template"] as Stream);
            var ws = pck.Workbook.Worksheets[1];
            int row = new Random().Next(10) + 10;   //Pick a random row to print the text
            ws.Cells[row,1].Value = "We make a small change here, after the template has been loaded...";
            ws.Cells[row, 1, row, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[row, 1, row, 5].Style.Fill.BackgroundColor.SetColor(Color.LightGoldenrodYellow);

            Response.BinaryWrite(pck.GetAsByteArray());
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("content-disposition", "attachment;  filename=Sample3.xlsx");
        }
        private void Sample4()
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

            Response.BinaryWrite(pck.GetAsByteArray());
            Response.ContentType = "application/vnd.ms-excel.sheet.macroEnabled.12";            //.xlsm files uses a different contenttype than .xlsx
            Response.AddHeader("content-disposition", "attachment;  filename=Sample4.xlsm");

        }

    }
}
