namespace EPPlusExcelFormDemo
{
    partial class ExcelForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tabControl_Worksheets = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.button_ApplyFormula = new System.Windows.Forms.Button();
            this.dataGridView_Ws1 = new System.Windows.Forms.DataGridView();
            this.A = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.B = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.C = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.D = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.E = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox_fx = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.button_Save = new System.Windows.Forms.Button();
            this.saveFileDialog_SaveExcel = new System.Windows.Forms.SaveFileDialog();
            this.tabControl_Worksheets.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Ws1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl_Worksheets
            // 
            this.tabControl_Worksheets.Controls.Add(this.tabPage1);
            this.tabControl_Worksheets.Controls.Add(this.tabPage2);
            this.tabControl_Worksheets.Location = new System.Drawing.Point(13, 63);
            this.tabControl_Worksheets.Name = "tabControl_Worksheets";
            this.tabControl_Worksheets.SelectedIndex = 0;
            this.tabControl_Worksheets.Size = new System.Drawing.Size(583, 386);
            this.tabControl_Worksheets.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.button_ApplyFormula);
            this.tabPage1.Controls.Add(this.dataGridView_Ws1);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.textBox_fx);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(575, 360);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "tabPage1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // button_ApplyFormula
            // 
            this.button_ApplyFormula.Location = new System.Drawing.Point(456, 14);
            this.button_ApplyFormula.Name = "button_ApplyFormula";
            this.button_ApplyFormula.Size = new System.Drawing.Size(64, 23);
            this.button_ApplyFormula.TabIndex = 3;
            this.button_ApplyFormula.Text = "Apply";
            this.button_ApplyFormula.UseVisualStyleBackColor = true;
            this.button_ApplyFormula.Click += new System.EventHandler(this.button_ApplyFormula_Click);
            // 
            // dataGridView_Ws1
            // 
            this.dataGridView_Ws1.AllowUserToAddRows = false;
            this.dataGridView_Ws1.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView_Ws1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView_Ws1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_Ws1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.A,
            this.B,
            this.C,
            this.D,
            this.E});
            this.dataGridView_Ws1.Location = new System.Drawing.Point(6, 43);
            this.dataGridView_Ws1.MultiSelect = false;
            this.dataGridView_Ws1.Name = "dataGridView_Ws1";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView_Ws1.RowHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView_Ws1.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.dataGridView_Ws1.ShowCellErrors = false;
            this.dataGridView_Ws1.ShowCellToolTips = false;
            this.dataGridView_Ws1.ShowEditingIcon = false;
            this.dataGridView_Ws1.ShowRowErrors = false;
            this.dataGridView_Ws1.Size = new System.Drawing.Size(563, 311);
            this.dataGridView_Ws1.TabIndex = 2;
            this.dataGridView_Ws1.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dataGridView_Ws1_CellBeginEdit);
            this.dataGridView_Ws1.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dataGridView_Ws1_CellValidating);
            // 
            // A
            // 
            this.A.HeaderText = "A";
            this.A.Name = "A";
            // 
            // B
            // 
            this.B.HeaderText = "B";
            this.B.Name = "B";
            // 
            // C
            // 
            this.C.HeaderText = "C";
            this.C.Name = "C";
            // 
            // D
            // 
            this.D.HeaderText = "D";
            this.D.Name = "D";
            // 
            // E
            // 
            this.E.HeaderText = "E";
            this.E.Name = "E";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(73, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(15, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "fx";
            // 
            // textBox_fx
            // 
            this.textBox_fx.Location = new System.Drawing.Point(94, 16);
            this.textBox_fx.Name = "textBox_fx";
            this.textBox_fx.Size = new System.Drawing.Size(356, 20);
            this.textBox_fx.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(575, 360);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // button_Save
            // 
            this.button_Save.Location = new System.Drawing.Point(17, 12);
            this.button_Save.Name = "button_Save";
            this.button_Save.Size = new System.Drawing.Size(95, 23);
            this.button_Save.TabIndex = 1;
            this.button_Save.Text = "Save Excelfile";
            this.button_Save.UseVisualStyleBackColor = true;
            this.button_Save.Click += new System.EventHandler(this.button_Save_Click);
            // 
            // ExcelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(619, 461);
            this.Controls.Add(this.button_Save);
            this.Controls.Add(this.tabControl_Worksheets);
            this.Name = "ExcelForm";
            this.Text = "EPPlus Excel demo";
            this.tabControl_Worksheets.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Ws1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl_Worksheets;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.DataGridView dataGridView_Ws1;
        private System.Windows.Forms.DataGridViewTextBoxColumn A;
        private System.Windows.Forms.DataGridViewTextBoxColumn B;
        private System.Windows.Forms.DataGridViewTextBoxColumn C;
        private System.Windows.Forms.DataGridViewTextBoxColumn D;
        private System.Windows.Forms.DataGridViewTextBoxColumn E;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_fx;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button button_Save;
        private System.Windows.Forms.Button button_ApplyFormula;
        private System.Windows.Forms.SaveFileDialog saveFileDialog_SaveExcel;
    }
}

