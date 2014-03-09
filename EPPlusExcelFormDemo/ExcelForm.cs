using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace EPPlusExcelFormDemo
{
    public partial class ExcelForm : Form
    {
        private ExcelPackage _package;
        private DataGridViewCell _currentCell;
        private Font _inactiveCellFont;
        private Font _activeCellFont;
        private const int NumberOfColumns = 5;

        public ExcelForm()
        {
            InitializeComponent();
            InitializePackage();
            InitPackageToUI();
            this.Closing += (sender, args) => _package.Dispose();
        }

        private void InitializePackage()
        {
            _package = new ExcelPackage(new MemoryStream());
            var ws1 = _package.Workbook.Worksheets.Add("Worksheet1");
            for (var col = 2; col < NumberOfColumns; col++)
            {
                for (var row = 1; row < 9; row++)
                {
                    ws1.Cells[row, col].Value = row*col;
                }
            }
            ws1.Cells[7, 1].Value = "SUM";
            ws1.Cells[7, 2].Formula = "SUM(B1:B6)";
            ws1.Cells[7, 3].Formula = "SUM(C1:C6)";
            ws1.Cells[7, 4].Formula = "SUM(D1:D6)";

            ws1.Cells[8, 1].Value = "STDEV";
            ws1.Cells[8, 2].Formula = "STDEV(B1:B6)";
            ws1.Cells[8, 3].Formula = "STDEV(C1:C6)";
            ws1.Cells[8, 4].Formula = "STDEV(D1:D6)";
            _package.Workbook.Calculate();
        }

        private void InitFonts(DataGridView gridView)
        {
            _activeCellFont = new Font(gridView.Font, FontStyle.Bold);
            _inactiveCellFont = gridView.Font;
        }

        private void InitEvents(DataGridView gridView)
        {
            gridView.CellEnter += DataGrid1OnCellEnter;
            gridView.CellLeave += DataGrid1OnCellLeave; 
        }

        /// <summary>
        /// Binds the EPPlus package (or actually only its first worksheet)
        /// to the DataGridView.
        /// </summary>
        private void InitPackageToUI()
        {
            var ws = _package.Workbook.Worksheets.First();
            var page1 = this.tabControl_Worksheets.Controls[0] as TabPage;
            page1.Text = ws.Name;
            var gridView = GetGrid();
            InitFonts(gridView);
            InitEvents(gridView);
            
            for (var row = 0; row < ws.Dimension.Rows; row++)
            {
                var gridRow = new DataGridViewRow {HeaderCell = {Value = (row + 1).ToString()}};
                for (var col = 0; col < NumberOfColumns; col++)
                {
                    var cell = ws.Cells[row + 1, col + 1];
                    using (var uiCell = new DataGridViewTextBoxCell())
                    {
                        uiCell.Value = cell.Value;
                        gridRow.Cells.Add(uiCell);
                    }
                }
                gridView.Rows.Add(gridRow);
            }
            gridView.Refresh();
        }

        private void BindPackageToUI()
        {
            var dataGrid1 = GetGrid();
            for (var row = 1; row < _package.Workbook.Worksheets.First().Dimension.Rows + 1; row++)
            {
                for (var col = 1; col <= NumberOfColumns; col++)
                {
                    var excelCell = _package.Workbook.Worksheets.First().Cells[row, col];
                    var gridViewCell = dataGrid1.Rows[row - 1].Cells[col - 1];
                    gridViewCell.Value = excelCell.Value;
                }
            }
            dataGrid1.Refresh();
        }

        private object CellValueToObject(string cellVal)
        {
            if (ConvertUtil.IsNumericString(cellVal))
            {
                return double.Parse(cellVal, CultureInfo.InvariantCulture);
            }
            return cellVal;
        }

        private void DataGrid1OnCellLeave(object sender, DataGridViewCellEventArgs e)
        {
            var dataGrid1 = GetGrid();
            var gridViewCell = dataGrid1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            var excelCell = _package.Workbook.Worksheets.First().Cells[e.RowIndex + 1, e.ColumnIndex + 1];
            gridViewCell.Style.ForeColor = Color.Black;
            gridViewCell.Style.BackColor = Color.White;
            gridViewCell.Style.Font = _inactiveCellFont;


            var result = gridViewCell.EditedFormattedValue.ToString();
            
            if (result.StartsWith("="))
            {
                excelCell.Formula = result.Substring(1);
            }
            else if(textBox_fx.Text.StartsWith("="))
            {
                excelCell.Formula = textBox_fx.Text.Substring(1);
            }
            else
            {
                excelCell.Value = CellValueToObject(result);
            }
            _package.Workbook.Calculate();
            //BindPackageToUI();
            //dataGrid1.Refresh();
        }


        private DataGridView GetGrid()
        {
            var page1 = this.tabControl_Worksheets.Controls[0] as TabPage;
            var dataGrid1 = page1.Controls["dataGridView_Ws1"] as DataGridView;
            return dataGrid1;
        }

        private void DataGrid1OnCellEnter(object sender, DataGridViewCellEventArgs e)
        {
            var dataGrid1 = GetGrid();
            dataGrid1.Refresh();
            BindPackageToUI();
            var cell = dataGrid1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            var excelCell = _package.Workbook.Worksheets.First().Cells[e.RowIndex + 1, e.ColumnIndex + 1];
            if (!string.IsNullOrEmpty(excelCell.Formula))
            {
                textBox_fx.Text = "=" + excelCell.Formula;
            }
            else if(excelCell.Value != null)
            {
                textBox_fx.Text = excelCell.Value.ToString();
            }
            cell.Style.ForeColor = Color.Blue;
            cell.Style.BackColor = Color.Gainsboro;
            cell.Style.Font = _activeCellFont;
            _currentCell = cell;
        }

        private void button_Save_Click(object sender, EventArgs e)
        {
            saveFileDialog_SaveExcel.Filter = "Excel files (*.xlsx)|*.xlsx";
            var dialogResult = saveFileDialog_SaveExcel.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                _package.SaveAs(new FileInfo(saveFileDialog_SaveExcel.FileName));
            }
        }

        private void button_ApplyFormula_Click(object sender, EventArgs e)
        {
            var row = _currentCell.RowIndex;
            var col = _currentCell.ColumnIndex;
            var txt = textBox_fx.Text;
            if (txt.StartsWith("="))
            {
                _package.Workbook.Worksheets.First().Cells[row + 1, col + 1].Formula = txt.Substring(1);
            }
            else
            {
                _package.Workbook.Worksheets.First().Cells[row + 1, col + 1].Formula = null;
                _package.Workbook.Worksheets.First().Cells[row + 1, col + 1].Value = CellValueToObject(txt);
            }
            _package.Workbook.Calculate();
            BindPackageToUI();
            this.Refresh();
        }
    }
}
