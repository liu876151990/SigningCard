﻿namespace SigningCard
{
    partial class SigningCardForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SigningCardForm));
            this.dataGridViewHoliday = new System.Windows.Forms.DataGridView();
            this.日 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.一 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.二 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.三 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.四 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.五 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.六 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.buttonImport = new System.Windows.Forms.Button();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.openFileDialogPunchRecord = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.buttonClipBoard = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewHoliday)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewHoliday
            // 
            this.dataGridViewHoliday.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewHoliday.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.日,
            this.一,
            this.二,
            this.三,
            this.四,
            this.五,
            this.六});
            this.dataGridViewHoliday.Location = new System.Drawing.Point(12, 12);
            this.dataGridViewHoliday.Name = "dataGridViewHoliday";
            this.dataGridViewHoliday.RowHeadersVisible = false;
            this.dataGridViewHoliday.RowTemplate.Height = 23;
            this.dataGridViewHoliday.Size = new System.Drawing.Size(297, 168);
            this.dataGridViewHoliday.TabIndex = 1;
            this.dataGridViewHoliday.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridViewHoliday_CellContentClick);
            this.dataGridViewHoliday.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridViewHoliday_CellContentDoubleClick);
            this.dataGridViewHoliday.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DataGridViewHoliday_CellMouseDoubleClick);
            // 
            // 日
            // 
            this.日.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.日.HeaderText = "日";
            this.日.Name = "日";
            this.日.Width = 42;
            // 
            // 一
            // 
            this.一.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.一.HeaderText = "一";
            this.一.Name = "一";
            this.一.Width = 42;
            // 
            // 二
            // 
            this.二.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.二.HeaderText = "二";
            this.二.Name = "二";
            this.二.Width = 42;
            // 
            // 三
            // 
            this.三.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.三.HeaderText = "三";
            this.三.Name = "三";
            this.三.Width = 42;
            // 
            // 四
            // 
            this.四.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.四.HeaderText = "四";
            this.四.Name = "四";
            this.四.Width = 42;
            // 
            // 五
            // 
            this.五.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.五.HeaderText = "五";
            this.五.Name = "五";
            this.五.Width = 42;
            // 
            // 六
            // 
            this.六.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.六.HeaderText = "六";
            this.六.Name = "六";
            this.六.Width = 42;
            // 
            // buttonImport
            // 
            this.buttonImport.Location = new System.Drawing.Point(219, 185);
            this.buttonImport.Name = "buttonImport";
            this.buttonImport.Size = new System.Drawing.Size(78, 26);
            this.buttonImport.TabIndex = 2;
            this.buttonImport.Text = "导入打卡";
            this.buttonImport.UseVisualStyleBackColor = true;
            this.buttonImport.Click += new System.EventHandler(this.ButtonImport_Click);
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(55, 186);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(111, 21);
            this.dateTimePicker1.TabIndex = 3;
            this.dateTimePicker1.ValueChanged += new System.EventHandler(this.DateTimePicker1_ValueChanged);
            // 
            // openFileDialogPunchRecord
            // 
            this.openFileDialogPunchRecord.FileName = "openFileDialogPunchRecord";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 192);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 5;
            this.label1.Text = "日期";
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2});
            this.dataGridView1.Location = new System.Drawing.Point(384, 10);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(447, 377);
            this.dataGridView1.TabIndex = 6;
            // 
            // Column1
            // 
            this.Column1.DataPropertyName = "Reason1";
            this.Column1.HeaderText = "签卡事由";
            this.Column1.Name = "Column1";
            this.Column1.Width = 200;
            // 
            // Column2
            // 
            this.Column2.DataPropertyName = "Reason2";
            this.Column2.HeaderText = "加班事由";
            this.Column2.Name = "Column2";
            this.Column2.Width = 200;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(681, 393);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(150, 27);
            this.button1.TabIndex = 7;
            this.button1.Text = "保存";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(226, 266);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(72, 16);
            this.checkBox1.TabIndex = 8;
            this.checkBox1.Text = "允许连班";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // buttonClipBoard
            // 
            this.buttonClipBoard.Location = new System.Drawing.Point(219, 219);
            this.buttonClipBoard.Name = "buttonClipBoard";
            this.buttonClipBoard.Size = new System.Drawing.Size(75, 23);
            this.buttonClipBoard.TabIndex = 9;
            this.buttonClipBoard.Text = "从剪贴板";
            this.buttonClipBoard.UseVisualStyleBackColor = true;
            this.buttonClipBoard.Click += new System.EventHandler(this.buttonClipBoard_Click);
            // 
            // SigningCardForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(843, 424);
            this.Controls.Add(this.buttonClipBoard);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.buttonImport);
            this.Controls.Add(this.dataGridViewHoliday);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SigningCardForm";
            this.Text = "签卡工具V1.4.3";
            this.Load += new System.EventHandler(this.SigningCardForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewHoliday)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridViewHoliday;
        private System.Windows.Forms.Button buttonImport;
        private System.Windows.Forms.DataGridViewTextBoxColumn 日;
        private System.Windows.Forms.DataGridViewTextBoxColumn 一;
        private System.Windows.Forms.DataGridViewTextBoxColumn 二;
        private System.Windows.Forms.DataGridViewTextBoxColumn 三;
        private System.Windows.Forms.DataGridViewTextBoxColumn 四;
        private System.Windows.Forms.DataGridViewTextBoxColumn 五;
        private System.Windows.Forms.DataGridViewTextBoxColumn 六;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.OpenFileDialog openFileDialogPunchRecord;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.Button buttonClipBoard;
    }
}

