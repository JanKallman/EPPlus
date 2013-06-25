using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using OfficeOpenXml;

namespace threadTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            for(int i=0;i<int.Parse(textBox1.Text);i++)
            {
                var t = new Thread(new ThreadStart(GenerateSheet));
                t.Name = "T" + i.ToString();
                t.IsBackground = true;
                t.Start();
            }
        }
        private void GenerateSheet()
        {
            ExcelPackage p = new ExcelPackage();
            var ws=p.Workbook.Worksheets.Add(string.Format("Thread {0}",Thread.CurrentThread.Name));
            for (int row = 1; row < 50000; row++)
            {
                for (int col = 1; col < 60; col++)
                {
                    ws.Cells[row, col].Value = ExcelAddressBase.GetAddress(row,col);
                }                
            }
            p.SaveAs(new System.IO.FileInfo(string.Format("c:\\temp\\threadtest\\{0}.xlsx",Thread.CurrentThread.Name)));
        }
    }
}
