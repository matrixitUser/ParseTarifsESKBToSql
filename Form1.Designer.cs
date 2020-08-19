namespace ParseEcxellTarifsToSQl
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.btnStart = new System.Windows.Forms.Button();
            this.Calendar = new System.Windows.Forms.MonthCalendar();
            this.tableProcess1 = new System.Windows.Forms.DataGridView();
            this.ColProcess = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColProc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColExe = new System.Windows.Forms.DataGridViewImageColumn();
            this.tableProcess2 = new System.Windows.Forms.DataGridView();
            this.colProcess2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColProc2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColExe2 = new System.Windows.Forms.DataGridViewImageColumn();
            this.tableProcess3 = new System.Windows.Forms.DataGridView();
            this.ColProcess3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColProc3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColExe3 = new System.Windows.Forms.DataGridViewImageColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.LabFinish = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btnOkFile = new System.Windows.Forms.Button();
            this.btnOkDate = new System.Windows.Forms.Button();
            this.btnOkWrite = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.tableProcess1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tableProcess2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.tableProcess3)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.btnOpenFile.Location = new System.Drawing.Point(668, 222);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(109, 33);
            this.btnOpenFile.TabIndex = 0;
            this.btnOpenFile.Text = "Открыть файлы";
            this.btnOpenFile.UseVisualStyleBackColor = false;
            this.btnOpenFile.Click += new System.EventHandler(this.BtnOpenFile_Click);
            // 
            // btnStart
            // 
            this.btnStart.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnStart.Location = new System.Drawing.Point(783, 222);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(109, 33);
            this.btnStart.TabIndex = 1;
            this.btnStart.Text = "Начать";
            this.btnStart.UseVisualStyleBackColor = false;
            this.btnStart.Click += new System.EventHandler(this.BtnStart_Click);
            // 
            // Calendar
            // 
            this.Calendar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.Calendar.Location = new System.Drawing.Point(728, 18);
            this.Calendar.Name = "Calendar";
            this.Calendar.TabIndex = 9;
            this.Calendar.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.Calendar_DateSelected);
            // 
            // tableProcess1
            // 
            this.tableProcess1.AllowUserToAddRows = false;
            this.tableProcess1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tableProcess1.ColumnHeadersVisible = false;
            this.tableProcess1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColProcess,
            this.ColProc,
            this.ColExe});
            this.tableProcess1.Enabled = false;
            this.tableProcess1.Location = new System.Drawing.Point(16, 31);
            this.tableProcess1.Name = "tableProcess1";
            this.tableProcess1.RowHeadersVisible = false;
            this.tableProcess1.Size = new System.Drawing.Size(198, 135);
            this.tableProcess1.TabIndex = 11;
            // 
            // ColProcess
            // 
            this.ColProcess.HeaderText = "Процесс";
            this.ColProcess.Name = "ColProcess";
            this.ColProcess.Width = 120;
            // 
            // ColProc
            // 
            this.ColProc.HeaderText = "ColProc";
            this.ColProc.Name = "ColProc";
            this.ColProc.Width = 45;
            // 
            // ColExe
            // 
            this.ColExe.HeaderText = "ColExe";
            this.ColExe.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Zoom;
            this.ColExe.Name = "ColExe";
            this.ColExe.Width = 30;
            // 
            // tableProcess2
            // 
            this.tableProcess2.AllowUserToAddRows = false;
            this.tableProcess2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tableProcess2.ColumnHeadersVisible = false;
            this.tableProcess2.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colProcess2,
            this.ColProc2,
            this.ColExe2});
            this.tableProcess2.Enabled = false;
            this.tableProcess2.Location = new System.Drawing.Point(231, 31);
            this.tableProcess2.Name = "tableProcess2";
            this.tableProcess2.RowHeadersVisible = false;
            this.tableProcess2.Size = new System.Drawing.Size(198, 92);
            this.tableProcess2.TabIndex = 13;
            // 
            // colProcess2
            // 
            this.colProcess2.HeaderText = "Процесс";
            this.colProcess2.Name = "colProcess2";
            this.colProcess2.Width = 120;
            // 
            // ColProc2
            // 
            this.ColProc2.HeaderText = "ColProc";
            this.ColProc2.Name = "ColProc2";
            this.ColProc2.Width = 45;
            // 
            // ColExe2
            // 
            this.ColExe2.HeaderText = "ColExe";
            this.ColExe2.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Zoom;
            this.ColExe2.Name = "ColExe2";
            this.ColExe2.Width = 30;
            // 
            // tableProcess3
            // 
            this.tableProcess3.AllowUserToAddRows = false;
            this.tableProcess3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tableProcess3.ColumnHeadersVisible = false;
            this.tableProcess3.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColProcess3,
            this.ColProc3,
            this.ColExe3});
            this.tableProcess3.Enabled = false;
            this.tableProcess3.Location = new System.Drawing.Point(449, 31);
            this.tableProcess3.Name = "tableProcess3";
            this.tableProcess3.RowHeadersVisible = false;
            this.tableProcess3.Size = new System.Drawing.Size(198, 92);
            this.tableProcess3.TabIndex = 14;
            // 
            // ColProcess3
            // 
            this.ColProcess3.HeaderText = "Процесс";
            this.ColProcess3.Name = "ColProcess3";
            this.ColProcess3.Width = 120;
            // 
            // ColProc3
            // 
            this.ColProc3.HeaderText = "ColProc";
            this.ColProc3.Name = "ColProc3";
            this.ColProc3.Width = 45;
            // 
            // ColExe3
            // 
            this.ColExe3.HeaderText = "ColExe";
            this.ColExe3.ImageLayout = System.Windows.Forms.DataGridViewImageCellLayout.Zoom;
            this.ColExe3.Name = "ColExe3";
            this.ColExe3.Width = 30;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(35, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(156, 19);
            this.label1.TabIndex = 15;
            this.label1.Text = "Мощность до 670 кВт";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(230, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(202, 19);
            this.label2.TabIndex = 16;
            this.label2.Text = "Мощность от 670 до 10 МВт";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(462, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(177, 19);
            this.label3.TabIndex = 17;
            this.label3.Text = "Мощность более 10 МВт";
            // 
            // LabFinish
            // 
            this.LabFinish.AutoSize = true;
            this.LabFinish.Font = new System.Drawing.Font("Times New Roman", 64F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.LabFinish.Location = new System.Drawing.Point(305, 137);
            this.LabFinish.Name = "LabFinish";
            this.LabFinish.Size = new System.Drawing.Size(318, 98);
            this.LabFinish.TabIndex = 18;
            this.LabFinish.Text = "Готово!";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.Location = new System.Drawing.Point(12, 179);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(115, 19);
            this.label4.TabIndex = 19;
            this.label4.Text = "Выбрать файлы";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label5.Location = new System.Drawing.Point(12, 209);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(100, 19);
            this.label5.TabIndex = 20;
            this.label5.Text = "Выбрать дату";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label6.Location = new System.Drawing.Point(7, 256);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(163, 19);
            this.label6.TabIndex = 21;
            this.label6.Text = "Запись в Базу Данных ";
            // 
            // btnOkFile
            // 
            this.btnOkFile.BackColor = System.Drawing.Color.White;
            this.btnOkFile.Location = new System.Drawing.Point(139, 179);
            this.btnOkFile.Name = "btnOkFile";
            this.btnOkFile.Size = new System.Drawing.Size(31, 23);
            this.btnOkFile.TabIndex = 22;
            this.btnOkFile.UseVisualStyleBackColor = false;
            // 
            // btnOkDate
            // 
            this.btnOkDate.Location = new System.Drawing.Point(139, 208);
            this.btnOkDate.Name = "btnOkDate";
            this.btnOkDate.Size = new System.Drawing.Size(31, 23);
            this.btnOkDate.TabIndex = 23;
            this.btnOkDate.UseVisualStyleBackColor = true;
            // 
            // btnOkWrite
            // 
            this.btnOkWrite.Location = new System.Drawing.Point(176, 255);
            this.btnOkWrite.Name = "btnOkWrite";
            this.btnOkWrite.Size = new System.Drawing.Size(31, 23);
            this.btnOkWrite.TabIndex = 24;
            this.btnOkWrite.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(901, 309);
            this.Controls.Add(this.btnOkWrite);
            this.Controls.Add(this.btnOkDate);
            this.Controls.Add(this.btnOkFile);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.LabFinish);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tableProcess3);
            this.Controls.Add(this.tableProcess2);
            this.Controls.Add(this.tableProcess1);
            this.Controls.Add(this.Calendar);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.btnOpenFile);
            this.Name = "Form1";
            this.Text = "ParseTarifToSql";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.tableProcess1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tableProcess2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.tableProcess3)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.MonthCalendar Calendar;
        private System.Windows.Forms.DataGridView tableProcess1;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColProcess;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColProc;
        private System.Windows.Forms.DataGridViewImageColumn ColExe;
        private System.Windows.Forms.DataGridView tableProcess2;
        private System.Windows.Forms.DataGridView tableProcess3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label LabFinish;
        private System.Windows.Forms.DataGridViewTextBoxColumn colProcess2;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColProc2;
        private System.Windows.Forms.DataGridViewImageColumn ColExe2;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColProcess3;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColProc3;
        private System.Windows.Forms.DataGridViewImageColumn ColExe3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnOkFile;
        private System.Windows.Forms.Button btnOkDate;
        private System.Windows.Forms.Button btnOkWrite;
    }
}

