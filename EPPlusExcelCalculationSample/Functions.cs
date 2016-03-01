using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPPlusExcelFormDemo
{
    public partial class frmFunctions : Form
    {
        public frmFunctions()
        {
            InitializeComponent();
        }
        public frmFunctions(List<string> functions)
        {
            InitializeComponent();

            var sb = new StringBuilder();
            foreach (var f in functions)
            {                
                sb.AppendLine(f.ToUpper());
            }
            textBox1.Text = sb.ToString();
        }
    }
}
