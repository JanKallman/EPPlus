/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * EPPlus is a fork of the ExcelPackage project
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
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		25-JAN-2010
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing.Imaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace ExcelPackageSamples
{
    /// <summary>
    /// Sample 6 - Reads the filesystem and makes a report.
    /// </summary>               
    class Sample6
    {
        #region "Icon API function"
        [StructLayout(LayoutKind.Sequential)]
        public struct SHFILEINFO
        {
            public IntPtr hIcon;
            public IntPtr iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        };
        public const uint SHGFI_ICON = 0x100;
        public const uint SHGFI_LARGEICON = 0x0;    // 'Large icon
        public const uint SHGFI_SMALLICON = 0x1;    // 'Small icon
        [DllImport("shell32.dll")]
        public static extern IntPtr SHGetFileInfo(string pszPath,
                                    uint dwFileAttributes,
                                    ref SHFILEINFO psfi,
                                    uint cbSizeFileInfo,
                                    uint uFlags);
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = CharSet.Auto)]
        extern static bool DestroyIcon(IntPtr handle);
        #endregion
        public class StatItem : IComparable<StatItem>
        {
            public string Name { get; set; }
            public int Count { get; set; }
            public long Size { get; set; }

            #region IComparable<StatItem> Members

            //Default compare Size
            public int CompareTo(StatItem other)
            {
                return Size < other.Size ? -1 :
                            (Size > other.Size ? 1 : 0);
            }

            #endregion
        }
        static int _maxLevels;

        static Dictionary<string, StatItem> _extStat = new Dictionary<string, StatItem>();
        static List<StatItem> _fileSize = new List<StatItem>();
        /// <summary>
        /// Sample 6 - Reads the filesystem and makes a report.
        /// </summary>
        /// <param name="outputDir">Output directory</param>
        /// <param name="dir">Directory to scan</param>
        /// <param name="depth">How many levels?</param>
        /// <param name="skipIcons">Skip the icons in column A. A lot faster</param>
        public static void RunSample6(DirectoryInfo outputDir, DirectoryInfo dir, int depth, bool skipIcons)
        {
            _maxLevels = depth;

            FileInfo newFile = new FileInfo(outputDir.FullName + @"\sample6.xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new FileInfo(outputDir.FullName + @"\sample6.xlsx");
            }
            
            //Create the workbook
            ExcelPackage pck = new ExcelPackage(newFile);
            //Add the Content sheet
            var ws = pck.Workbook.Worksheets.Add("Content");

            ws.View.ShowGridLines = false;

            ws.Column(1).Width = 2.5;
            ws.Column(2).Width = 60;
            ws.Column(3).Width = 16;
            ws.Column(4).Width = 20;
            ws.Column(5).Width = 20;
            
            //This set the outline for column 4 and 5 and hide them
            ws.Column(4).OutlineLevel = 1;
            ws.Column(4).Collapsed = true;
            ws.Column(5).OutlineLevel = 1;
            ws.Column(5).Collapsed = true;
            ws.OutLineSummaryRight = true;
            
            //Headers
            ws.Cells["B1"].Value = "Name";
            ws.Cells["C1"].Value = "Size";
            ws.Cells["D1"].Value = "Created";
            ws.Cells["E1"].Value = "Last modified";
            ws.Cells["B1:E1"].Style.Font.Bold = true;
            
            //height is 20 pixels 
            double height = 20 * 0.75;
            //Start at row 2;
            int row = 2;

            //Load the directory content to sheet 1
            row = AddDirectory(ws, dir, row, height, 0, skipIcons);
            ws.OutLineSummaryBelow = false;

            //Format columns
            ws.Cells[1, 3, row - 1, 3].Style.Numberformat.Format = "#,##0";
            ws.Cells[1, 4, row - 1, 4].Style.Numberformat.Format = "yyyy-MM-dd hh:mm";
            ws.Cells[1, 5, row - 1, 5].Style.Numberformat.Format = "yyyy-MM-dd hh:mm";

            //Add the textbox
            var shape = ws.Drawings.AddShape("txtDesc", eShapeStyle.Rect);
            shape.SetPosition(0, 5, 5, 5);
            shape.SetSize(400, 200);

            shape.Text = "This example demonstrates how to create various drawing objects like Pictures, Shapes and charts.\n\r\n\rThe first sheet contains all subdirectories and files with an icon, name, size and dates.\n\r\n\rThe second sheet contains statistics about extensions and the top-10 largest files.";
            shape.Fill.Style = eFillStyle.SolidFill;
            shape.Fill.Color = Color.DarkSlateGray;
            shape.Fill.Transparancy = 20;
            shape.Border.Fill.Style = eFillStyle.SolidFill;
            shape.Border.LineStyle = eLineStyle.LongDash;
            shape.Border.Width = 1;
            shape.Border.Fill.Color = Color.Black;
            shape.Border.LineCap = eLineCap.Round;
            shape.TextAnchoring = eTextAnchoringType.Top;
            shape.TextVertical = eTextVerticalType.Horizontal;
            shape.TextAnchoringControl=false;

            //Add graph sheet
            AddGraphs(pck, row, string.Format("Statistics for {0}",dir.FullName));

            //Add a HyperLink to the statistics sheet. 
            var namedStyle = pck.Workbook.Styles.CreateNamedStyle("HyperLink");   //This one is language dependent
            namedStyle.Style.Font.UnderLine = true;
            namedStyle.Style.Font.Color.SetColor(Color.Blue);
            ws.Cells["K12"].Hyperlink = new ExcelHyperLink("Statistics!A1", "Statistics");
            ws.Cells["K12"].StyleName = "HyperLink";

            //Printer settings
            ws.PrinterSettings.FitToPage = true;
            ws.PrinterSettings.FitToWidth = 1;
            ws.PrinterSettings.FitToHeight = 0;

            //Done! save the sheet
            pck.Save();
        }
        /// <summary>
        /// Add the second sheet containg the graphs
        /// </summary>
        /// <param name="pck">Package</param>
        /// <param name="rows"></param>
        /// <param name="header"></param>
        private static void AddGraphs(ExcelPackage pck, int rows, string header)
        {
            var ws = pck.Workbook.Worksheets.Add("Statistics");
            ws.View.ShowGridLines = false;

            ws.Column(1).Width = 40;
            ws.Column(2).Width = 20;

            //Set first the header and format it
            ws.Cells["A1"].Value = header;
            using (ExcelRange r = ws.Cells["A1:N1"])
            {
                r.Merge = true;
                r.Style.Font.SetFromFont(new Font("Arial", 22, FontStyle.Italic));
                r.Style.Font.Color.SetColor(Color.White);
                r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
            }

            //Start with the Extention Size 
            List<StatItem> lst = new List<StatItem>(_extStat.Values);           
            lst.Sort();

            //Add rows
            int row=AddStatRows(ws, lst, 2, "Extensions", "Size");

            //Add the piechart
            var pieChart = ws.Drawings.AddChart("crtExtensionsSize", eChartType.PieExploded3D) as ExcelPieChart;
            //Set top left corner to row 1 column 2
            pieChart.SetPosition(1, 0, 2, 0);
            pieChart.SetSize(400, 400);
            pieChart.Series.Add(ExcelRange.GetAddress(4, 2, row-1, 2), ExcelRange.GetAddress(4, 1, row-1, 1));

            pieChart.Title.Text = "Extension Size";
            //Set datalabels and remove the legend
            pieChart.DataLabel.ShowCategory = true;
            pieChart.DataLabel.ShowPercent = true;
            pieChart.DataLabel.ShowLeaderLines = true;
            pieChart.Legend.Remove();

            //Resort on Count and add the rows
            lst.Sort((first,second) => first.Count < second.Count ? -1 : first.Count > second.Count ? 1 : 0);
            row=AddStatRows(ws, lst, 16, "", "Count");

            //Add the Doughnut chart
            var doughtnutChart = ws.Drawings.AddChart("crtExtensionCount", eChartType.DoughnutExploded) as ExcelDoughnutChart;
            //Set position to row 1 column 7 and 50 pixels offset
            doughtnutChart.SetPosition(1, 0, 7, 50);
            doughtnutChart.SetSize(400, 400);
            doughtnutChart.Series.Add(ExcelRange.GetAddress(17, 2, row - 1, 2), ExcelRange.GetAddress(17, 1, row - 1, 1));

            doughtnutChart.Title.Text = "Extension Count";
            doughtnutChart.DataLabel.ShowPercent = true;
            doughtnutChart.DataLabel.ShowLeaderLines = true;

            //Top-10 filesize
            _fileSize.Sort();
            row=AddStatRows(ws, _fileSize, 29, "Files", "Size");
            var barChart = ws.Drawings.AddChart("crtFiles", eChartType.BarClustered3D) as ExcelBarChart;
            //3d Settings
            barChart.View3D.RotX = 0;
            barChart.View3D.Perspective = 0;

            barChart.SetPosition(22, 0, 2, 0);
            barChart.SetSize(800, 398);
            barChart.Series.Add(ExcelRange.GetAddress(31, 2, row - 1, 2), ExcelRange.GetAddress(31, 1, row - 1, 1));
            barChart.Series[0].Header = "Size";
            barChart.Title.Text = "Top File size";

            //Format the Size and Count column
            ws.Cells["B3:B42"].Style.Numberformat.Format = "#,##0";
            //Set a border around
            ws.Cells["N1:N43"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells["A43:N43"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            //And last the printersettings
            ws.PrinterSettings.Orientation = eOrientation.Landscape;
            ws.PrinterSettings.FitToPage = true;
            ws.PrinterSettings.Scale = 67;
        }
        /// <summary>
        /// Add statistic-rows to the statistics sheet.
        /// </summary>
        /// <param name="ws">Worksheet</param>
        /// <param name="lst">List with statistics</param>
        /// <param name="startRow"></param>
        /// <param name="header">Header text</param>
        /// <param name="propertyName">Size or Count</param>
        /// <returns></returns>
        private static int AddStatRows(ExcelWorksheet ws, List<StatItem> lst, int startRow, string header, string propertyName)
        {
            //Add Headers
            int row = startRow;
            if (header != "")
            {
                ws.Cells[row, 1].Value = header;
                using (ExcelRange r = ws.Cells[row, 1, row, 2])
                {
                    r.Merge = true;
                    r.Style.Font.SetFromFont(new Font("Arial", 16, FontStyle.Italic));
                    r.Style.Font.Color.SetColor(Color.White);
                    r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.CenterContinuous;
                    r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79 , 129, 189));
                }
                row++;
            }
            
            //Header 2
            ws.Cells[row, 1].Value = "Name";
            ws.Cells[row, 2].Value = propertyName;
            using (ExcelRange r = ws.Cells[row, 1, row, 2])
            {
                r.Style.Font.SetFromFont(new Font("Arial", 12, FontStyle.Bold));
                AlterColor(ws, row);
            }

            row++;
            //Add top 10 rows
            for (int i = 0; i < 10; i++)
            {
                if (lst.Count - i > 0)
                {
                    ws.Cells[row, 1].Value = lst[lst.Count - i - 1].Name;
                    if (propertyName == "Size")
                    {
                        ws.Cells[row, 2].Value = lst[lst.Count - i - 1].Size;
                    }
                    else
                    {
                        ws.Cells[row, 2].Value = lst[lst.Count - i - 1].Count;
                    }

                    AlterColor(ws, row);
                    row++;
                }
            }
            
            //If we have more than 10 items, sum...
            long rest = 0;
            for (int i = 0; i < lst.Count - 10; i++)
            {
                if (propertyName == "Size")
                {
                    rest += lst[i].Size;
                }
                else
                {
                    rest += lst[i].Count;
                }
            }
            //... and add an others row
            if (rest > 0)
            {
                ws.Cells[row, 1].Value = "Others";
                ws.Cells[row, 2].Value = rest;
                ws.Cells[row, 1, row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 1, row, 2].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                AlterColor(ws, row);
                row++;
            }

            return row;
        }
        /// <summary>
        /// Just alters the colors in the list
        /// </summary>
        /// <param name="ws">The worksheet</param>
        /// <param name="row">Startrow</param>
        private static void AlterColor(ExcelWorksheet ws, int row)
        {
            using (ExcelRange rowRange = ws.Cells[row, 1, row, 2])
            {
                rowRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                if(row % 2==1)
                {
                    rowRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }
                else
                {
                    rowRange.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                }
            }
        }

        private static int AddDirectory(ExcelWorksheet ws, DirectoryInfo dir, int row, double height, int level, bool skipIcons)
        {
            //Get the icon as a bitmap
            Console.WriteLine("Directory " + dir.Name);
            if (!skipIcons)
            {
                Bitmap icon = GetIcon(dir.FullName);

                ws.Row(row).Height = height;
                //Add the icon as a picture
                if (icon != null)
                {
                    ExcelPicture pic = ws.Drawings.AddPicture("pic" + (row).ToString(), icon);
                    pic.SetPosition((int)20 * (row - 1) + 2, 0);
                }
            }
            ws.Cells[row, 2].Value = dir.Name;
            ws.Cells[row, 4].Value = dir.CreationTime;
            ws.Cells[row, 5].Value = dir.LastAccessTime;

            ws.Cells[row, 2, row, 5].Style.Font.Bold = true;
            //Sets the outline depth
            ws.Row(row).OutlineLevel = level;

            int prevRow = row;
            row++;
            //Add subdirectories
            foreach (DirectoryInfo subDir in dir.GetDirectories())
            {
                if (level < _maxLevels)
                {
                    row = AddDirectory(ws, subDir, row, height, level + 1, skipIcons);
                }                           
            }
            
            //Add files in the directory
            foreach (FileInfo file in dir.GetFiles())
            {
                if (!skipIcons)
                {
                    Bitmap fileIcon = GetIcon(file.FullName);

                    ws.Row(row).Height = height;
                    if (fileIcon != null)
                    {
                        ExcelPicture pic = ws.Drawings.AddPicture("pic" + (row).ToString(), fileIcon);
                        pic.SetPosition((int)20 * (row - 1) + 2, 0);
                    }
                }

                ws.Cells[row, 2].Value = file.Name;
                ws.Cells[row, 3].Value = file.Length;
                ws.Cells[row, 4].Value = file.CreationTime;
                ws.Cells[row, 5].Value = file.LastAccessTime;

                ws.Row(row).OutlineLevel = level+1;

                AddStatistics(file);

                row++;
            }

            //Add a subtotal for the directory
            if (row -1 > prevRow)
            { 
                ws.Cells[prevRow, 3].Formula = string.Format("SUBTOTAL(9, {0})", ExcelCellBase.GetAddress(prevRow + 1, 3, row - 1, 3));
            }
            else
            {
                ws.Cells[prevRow, 3].Value = 0;
            }

            return row;
        }
        /// <summary>
        /// Add statistics to the collections 
        /// </summary>
        /// <param name="file"></param>
        private static void AddStatistics(FileInfo file)
        {
            //Extension
            if (_extStat.ContainsKey(file.Extension))
            {
                _extStat[file.Extension].Count++;
                _extStat[file.Extension].Size+=file.Length;
            }
            else
            {
                string ext = file.Extension.Length > 0 ? file.Extension.Remove(0, 1) : "";
                _extStat.Add(file.Extension, new StatItem() { Name = ext, Count = 1, Size = file.Length });
            }
            
            //File top 10;
            if (_fileSize.Count < 10)
            {
                _fileSize.Add(new StatItem { Name = file.Name, Size = file.Length });
                if (_fileSize.Count == 10)
                {
                    _fileSize.Sort();
                }
            }
            else if(_fileSize[0].Size < file.Length)
            {
                _fileSize.RemoveAt(0);
                _fileSize.Add(new StatItem { Name = file.Name, Size = file.Length });
                _fileSize.Sort();
            }
        }
        /// <summary>
        /// Gets the icon for a file or directory
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns></returns>
        private static Bitmap GetIcon(string FileName)
        {
            try
            {
                SHFILEINFO shinfo = new SHFILEINFO();                

                var ret = SHGetFileInfo(FileName,
                                          0,
                                          ref shinfo,
                                          (uint)Marshal.SizeOf(shinfo),
                                          SHGFI_ICON | SHGFI_SMALLICON);

                if (shinfo.hIcon == IntPtr.Zero) return null;

                Bitmap bmp = Icon.FromHandle(shinfo.hIcon).ToBitmap();
                DestroyIcon(shinfo.hIcon);

                //Fix transparant color 
                Color InvalidColor = Color.FromArgb(0, 0, 0, 0);
                for (int w = 0; w < bmp.PhysicalDimension.Width; w++)
                {
                    for (int h = 0; h < bmp.PhysicalDimension.Height; h++)
                    {
                        if (bmp.GetPixel(w, h) == InvalidColor)
                        {
                            bmp.SetPixel(w, h, Color.Transparent);
                        }
                    }
                }

                return bmp;
            }
            catch
            {
                return null;
            }
        }
    }
}
