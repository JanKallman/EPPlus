using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;

namespace GeneratePackage
{
    class Program
    {
        static ExcelPackage package;

        static void Main(string[] args)
        {
            Console.Out.WriteLine("Press enter to Start");
            Console.In.ReadLine();

            for (int ca = 0; ca < 10; ca++)
            {
                if (ca % 10 == 0)
                    Console.Out.WriteLine(ca);
                CreateDeletePackage(4,100000);
            }
            while (!Console.In.ReadLine().Equals("exit", StringComparison.OrdinalIgnoreCase))
            {
                CreateDeletePackage(1,1000);
            }
        }

        private static void CreateDeletePackage(int Sheets, int rows)
        {
            List<object> row = new List<object>();
            row.Add(1);
            row.Add("Some text");
            row.Add(12.0);
            row.Add("Some larger text that has completely no meaning.  How much wood can a wood chuck chuck if a wood chuck could chuck wood.  A wood chuck could chuck as much wood as a wood chuck could chuck wood.");

            FileInfo LocalFullFileName = new FileInfo(Path.GetTempFileName());
            LocalFullFileName.Delete();
            package = new ExcelPackage(LocalFullFileName);

            try
            {
                for (int ca = 0; ca < Sheets; ca++)
                {
                    CreateWorksheet("Sheet" + (ca+1), row, rows);
                }

                package.Save();
            }
            finally
            {
                LocalFullFileName.Refresh();
                if (LocalFullFileName.Exists)
                {
                    LocalFullFileName.Delete();
                }

                package.Dispose();
                package = null;

                GC.Collect();
            }
        }

        private static void CreateWorksheet(string sheetName, List<object> row, int numrows)
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
            for (int ca = 1; ca <= numrows; ca++)
            {
                AddDataRow(worksheet, ca, row);
            }
        }

        public static void AddDataRow(ExcelWorksheet worksheet, int row, IEnumerable<object> values)
        {
            int ca = 0;
            foreach (object v in values)
            {
                ca++;
                using (ExcelRange cell = worksheet.Cells[row, ca])
                {
                    object value = v;
                    cell.Value = value;
                }
            }

            worksheet.Row(row).Height = 10.2;
        }
    }
}
