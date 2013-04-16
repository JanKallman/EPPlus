using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlusSamples
{
    class Sample16
    {
        public static void RunSample16(DirectoryInfo outputDir)
        {
            using (var package = new ExcelPackage())
            {
                var dataTable = new DataTable("test");
                dataTable.Columns.Add("col1");
                dataTable.Columns.Add("col2");
                dataTable.Columns.Add("col3");
                dataTable.Columns.Add("col4");
                dataTable.Rows.Add("qwe11", "qwe12", "qwe13", "qwe14");
                dataTable.Rows.Add("qwe21", "qwe22", "qwe23", "qwe24");
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(dataTable.TableName);
                worksheet.Cells["A1"].LoadFromDataTable(dataTable, true, TableStyles.None);
                worksheet.Protection.AllowSelectLockedCells = false;
                worksheet.Protection.AllowSelectUnlockedCells = true;
                worksheet.Protection.AllowSort = true;
                worksheet.Protection.AllowFormatColumns = true;
                worksheet.Protection.AllowAutoFilter = true;
                worksheet.Protection.AllowEditObject = true;
                worksheet.Protection.IsProtected = true;
                var r1=worksheet.ProtectedRanges.Add("Range1", new ExcelAddress(1, 1, worksheet.Dimension.End.Row, 4));
                worksheet.ProtectedRanges.Remove(r1);
                var r2 = worksheet.ProtectedRanges.Add("Range2", new ExcelAddress(1, 1, worksheet.Dimension.End.Row, 4));
                r2.SetPassword("EPPlus");

                worksheet.Column(1).Width = 30;
                worksheet.Column(2).Width = 30;
                worksheet.Column(3).Width = 100;
                worksheet.Column(4).Width = 100;
                worksheet.Cells[1, 4, worksheet.Dimension.End.Row, 4].Style.Locked = false;
                worksheet.Cells[1, 3, worksheet.Dimension.End.Row, 4].Style.WrapText = true;

                using (var fs = new FileStream(Path.Combine(outputDir.ToString(), "sample16.xlsx"), FileMode.Create))
                    package.SaveAs(fs);
            }

            using (var fs = new FileStream(Path.Combine(outputDir.ToString(), "sample16.xlsx"), FileMode.Open, FileAccess.Read))
            using (var package = new ExcelPackage(fs))
            {
                foreach (var worksheet1 in package.Workbook.Worksheets)
                {
                    var prCollection = worksheet1.ProtectedRanges;
                    if (prCollection.Count != 1)
                        throw new InvalidOperationException("Expected 1 element");
                }
            }
        }
    }
}
