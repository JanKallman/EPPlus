using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Security.Cryptography.X509Certificates;

namespace EPPlusSamples
{
    class Sample15
    {
        public static void VBASample(DirectoryInfo outputDir)
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
            pck.Workbook.CodeModule.Code="Private Sub Workbook_Open()\r\nCreateBubbleChart\r\nEnd Sub";

            //Optionally, Sign the code with your company certificate.
            /*            
            X509Store store = new X509Store(StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            pck.Workbook.VbaProject.Signature.Certificate = store.Certificates[0];
            */

            //Password protect your code
            pck.Workbook.VbaProject.Protection.SetPassword("EPPlus");

            //And Save as xlsm
            pck.SaveAs(new FileInfo(outputDir.FullName + @"\sample15.xlsm"));
        }
    }
}
