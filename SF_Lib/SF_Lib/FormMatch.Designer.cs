namespace winform_SF_Lib
    {
    partial class FormMatch
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
            this.components = new System.ComponentModel.Container();
            this.dataGridView_SF = new System.Windows.Forms.DataGridView();
            this.ColumnSf_MemberID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.First_Name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Trans_Date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Agent_nameFull = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_SF_Product_Name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_SF_Payroll_date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridView_Carrier = new System.Windows.Forms.DataGridView();
            this.Column_Member_ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnPhone = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_Product_Name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_Termination_Date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnPayroll_date_ca = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewLinked = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_SP_Phone = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnMember_ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_Enrollment = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnCarrName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_Phone = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.unbindToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox_SF = new System.Windows.Forms.TextBox();
            this.textBox_Carrier = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.btn_ExportReady = new System.Windows.Forms.Button();
            this.backgroundWorkerFillExcel = new System.ComponentModel.BackgroundWorker();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_SF)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Carrier)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewLinked)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView_SF
            // 
            this.dataGridView_SF.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView_SF.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_SF.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnSf_MemberID,
            this.First_Name,
            this.Trans_Date,
            this.Agent_nameFull,
            this.Column_SF_Product_Name,
            this.Column_SF_Payroll_date});
            this.dataGridView_SF.Location = new System.Drawing.Point(3, 3);
            this.dataGridView_SF.MultiSelect = false;
            this.dataGridView_SF.Name = "dataGridView_SF";
            this.dataGridView_SF.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_SF.Size = new System.Drawing.Size(338, 317);
            this.dataGridView_SF.TabIndex = 4;
            // 
            // ColumnSf_MemberID
            // 
            this.ColumnSf_MemberID.DataPropertyName = "Sf_MemberID";
            this.ColumnSf_MemberID.HeaderText = "Sf Policy";
            this.ColumnSf_MemberID.Name = "ColumnSf_MemberID";
            this.ColumnSf_MemberID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // First_Name
            // 
            this.First_Name.DataPropertyName = "First_Name";
            this.First_Name.FillWeight = 10F;
            this.First_Name.HeaderText = "Name";
            this.First_Name.Name = "First_Name";
            this.First_Name.ReadOnly = true;
            this.First_Name.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Trans_Date
            // 
            this.Trans_Date.DataPropertyName = "Application_Date";
            this.Trans_Date.FillWeight = 10F;
            this.Trans_Date.HeaderText = "Enrollment";
            this.Trans_Date.Name = "Trans_Date";
            this.Trans_Date.ReadOnly = true;
            this.Trans_Date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Agent_nameFull
            // 
            this.Agent_nameFull.DataPropertyName = "Phone";
            this.Agent_nameFull.FillWeight = 10F;
            this.Agent_nameFull.HeaderText = "Phone";
            this.Agent_nameFull.Name = "Agent_nameFull";
            this.Agent_nameFull.ReadOnly = true;
            this.Agent_nameFull.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Column_SF_Product_Name
            // 
            this.Column_SF_Product_Name.DataPropertyName = "Product_Name";
            this.Column_SF_Product_Name.HeaderText = "Product Name";
            this.Column_SF_Product_Name.Name = "Column_SF_Product_Name";
            this.Column_SF_Product_Name.ReadOnly = true;
            this.Column_SF_Product_Name.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Column_SF_Payroll_date
            // 
            this.Column_SF_Payroll_date.DataPropertyName = "Payroll_date";
            this.Column_SF_Payroll_date.HeaderText = "Payroll date";
            this.Column_SF_Payroll_date.Name = "Column_SF_Payroll_date";
            this.Column_SF_Payroll_date.ReadOnly = true;
            this.Column_SF_Payroll_date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dataGridView_Carrier
            // 
            this.dataGridView_Carrier.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView_Carrier.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_Carrier.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column_Member_ID,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.ColumnPhone,
            this.Column_Product_Name,
            this.Column_Termination_Date,
            this.ColumnPayroll_date_ca});
            this.dataGridView_Carrier.Location = new System.Drawing.Point(430, 3);
            this.dataGridView_Carrier.MultiSelect = false;
            this.dataGridView_Carrier.Name = "dataGridView_Carrier";
            this.dataGridView_Carrier.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_Carrier.Size = new System.Drawing.Size(339, 317);
            this.dataGridView_Carrier.TabIndex = 6;
            // 
            // Column_Member_ID
            // 
            this.Column_Member_ID.DataPropertyName = "Member_ID";
            this.Column_Member_ID.HeaderText = "Policy No.";
            this.Column_Member_ID.Name = "Column_Member_ID";
            this.Column_Member_ID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "First_Name";
            this.dataGridViewTextBoxColumn2.FillWeight = 10F;
            this.dataGridViewTextBoxColumn2.HeaderText = "Name";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            this.dataGridViewTextBoxColumn2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "Application_Date";
            this.dataGridViewTextBoxColumn3.FillWeight = 10F;
            this.dataGridViewTextBoxColumn3.HeaderText = "Enrollment";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.ReadOnly = true;
            this.dataGridViewTextBoxColumn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnPhone
            // 
            this.ColumnPhone.DataPropertyName = "Phone";
            this.ColumnPhone.HeaderText = "Phone";
            this.ColumnPhone.Name = "ColumnPhone";
            this.ColumnPhone.ReadOnly = true;
            this.ColumnPhone.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Column_Product_Name
            // 
            this.Column_Product_Name.DataPropertyName = "Product_Name";
            this.Column_Product_Name.HeaderText = "Product Name";
            this.Column_Product_Name.Name = "Column_Product_Name";
            this.Column_Product_Name.ReadOnly = true;
            this.Column_Product_Name.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Column_Termination_Date
            // 
            this.Column_Termination_Date.DataPropertyName = "Termination_Date";
            this.Column_Termination_Date.HeaderText = "Term Date";
            this.Column_Termination_Date.Name = "Column_Termination_Date";
            this.Column_Termination_Date.ReadOnly = true;
            this.Column_Termination_Date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnPayroll_date_ca
            // 
            this.ColumnPayroll_date_ca.DataPropertyName = "Payroll_date";
            this.ColumnPayroll_date_ca.HeaderText = "Payroll Date";
            this.ColumnPayroll_date_ca.Name = "ColumnPayroll_date_ca";
            this.ColumnPayroll_date_ca.ReadOnly = true;
            this.ColumnPayroll_date_ca.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dataGridViewLinked
            // 
            this.dataGridViewLinked.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewLinked.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewLinked.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.Column_SP_Phone,
            this.dataGridViewTextBoxColumn6,
            this.ColumnMember_ID,
            this.Column_Enrollment,
            this.ColumnCarrName,
            this.Column_Phone});
            this.dataGridViewLinked.ContextMenuStrip = this.contextMenuStrip1;
            this.dataGridViewLinked.Location = new System.Drawing.Point(6, 399);
            this.dataGridViewLinked.MultiSelect = false;
            this.dataGridViewLinked.Name = "dataGridViewLinked";
            this.dataGridViewLinked.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewLinked.Size = new System.Drawing.Size(772, 125);
            this.dataGridViewLinked.TabIndex = 7;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "Sf_MemberID";
            this.dataGridViewTextBoxColumn4.HeaderText = "Sf Policy";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.ReadOnly = true;
            this.dataGridViewTextBoxColumn4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "SF_First_Name";
            this.dataGridViewTextBoxColumn5.FillWeight = 10F;
            this.dataGridViewTextBoxColumn5.HeaderText = "SF_Name";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.ReadOnly = true;
            this.dataGridViewTextBoxColumn5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Column_SP_Phone
            // 
            this.Column_SP_Phone.DataPropertyName = "SF_phone";
            this.Column_SP_Phone.HeaderText = "SF_Phone";
            this.Column_SP_Phone.Name = "Column_SP_Phone";
            this.Column_SP_Phone.ReadOnly = true;
            this.Column_SP_Phone.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.DataPropertyName = "SF_Application_Date";
            this.dataGridViewTextBoxColumn6.FillWeight = 10F;
            this.dataGridViewTextBoxColumn6.HeaderText = "SF_Enrollment";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            this.dataGridViewTextBoxColumn6.ReadOnly = true;
            this.dataGridViewTextBoxColumn6.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnMember_ID
            // 
            this.ColumnMember_ID.DataPropertyName = "Carrier_Policynumber";
            this.ColumnMember_ID.HeaderText = "Policy No.";
            this.ColumnMember_ID.Name = "ColumnMember_ID";
            this.ColumnMember_ID.ReadOnly = true;
            this.ColumnMember_ID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Column_Enrollment
            // 
            this.Column_Enrollment.DataPropertyName = "Carrier_Application_Date";
            this.Column_Enrollment.HeaderText = "Enrollment";
            this.Column_Enrollment.Name = "Column_Enrollment";
            this.Column_Enrollment.ReadOnly = true;
            this.Column_Enrollment.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnCarrName
            // 
            this.ColumnCarrName.DataPropertyName = "Carrier_FirstName";
            this.ColumnCarrName.HeaderText = "Name";
            this.ColumnCarrName.Name = "ColumnCarrName";
            this.ColumnCarrName.ReadOnly = true;
            this.ColumnCarrName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Column_Phone
            // 
            this.Column_Phone.DataPropertyName = "Carrier_Phone";
            this.Column_Phone.HeaderText = "Phone";
            this.Column_Phone.Name = "Column_Phone";
            this.Column_Phone.ReadOnly = true;
            this.Column_Phone.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.unbindToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(112, 26);
            // 
            // unbindToolStripMenuItem
            // 
            this.unbindToolStripMenuItem.Name = "unbindToolStripMenuItem";
            this.unbindToolStripMenuItem.Size = new System.Drawing.Size(111, 22);
            this.unbindToolStripMenuItem.Text = "UnLink";
            this.unbindToolStripMenuItem.Click += new System.EventHandler(this.unbindToolStripMenuItem_Click);
            // 
            // button1
            // 
            this.button1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.button1.Location = new System.Drawing.Point(2, 119);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 5;
            this.button1.Text = "Link";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox_SF
            // 
            this.textBox_SF.Location = new System.Drawing.Point(125, 31);
            this.textBox_SF.Name = "textBox_SF";
            this.textBox_SF.Size = new System.Drawing.Size(222, 20);
            this.textBox_SF.TabIndex = 1;
            this.textBox_SF.TextChanged += new System.EventHandler(this.textBox_SF_TextChanged);
            this.textBox_SF.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox_SF_KeyPress);
            // 
            // textBox_Carrier
            // 
            this.textBox_Carrier.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_Carrier.Location = new System.Drawing.Point(542, 30);
            this.textBox_Carrier.Name = "textBox_Carrier";
            this.textBox_Carrier.Size = new System.Drawing.Size(233, 20);
            this.textBox_Carrier.TabIndex = 2;
            this.textBox_Carrier.TextChanged += new System.EventHandler(this.textBox_Carrier_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(116, 13);
            this.label1.TabIndex = 20;
            this.label1.Text = "Filter Salesforce values";
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(433, 34);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(105, 13);
            this.label2.TabIndex = 21;
            this.label2.Text = "Filter Provider values";
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 83F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.dataGridView_SF, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.panel1, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.dataGridView_Carrier, 2, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(6, 56);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(772, 323);
            this.tableLayoutPanel1.TabIndex = 3;
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.Controls.Add(this.button1);
            this.panel1.Location = new System.Drawing.Point(347, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(77, 317);
            this.panel1.TabIndex = 16;
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Location = new System.Drawing.Point(703, 530);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 22;
            this.button2.Text = "Close";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btn_ExportReady
            // 
            this.btn_ExportReady.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_ExportReady.Location = new System.Drawing.Point(623, 530);
            this.btn_ExportReady.Name = "btn_ExportReady";
            this.btn_ExportReady.Size = new System.Drawing.Size(74, 23);
            this.btn_ExportReady.TabIndex = 23;
            this.btn_ExportReady.Text = "Export Excel";
            this.btn_ExportReady.UseVisualStyleBackColor = true;
            this.btn_ExportReady.Click += new System.EventHandler(this.btn_ExportReady_Click);
            // 
            // backgroundWorkerFillExcel
            // 
            this.backgroundWorkerFillExcel.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerFillExcel_DoWork);
            this.backgroundWorkerFillExcel.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerFillExcel_RunWorkerCompleted);
            // 
            // FormMatch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.button2;
            this.ClientSize = new System.Drawing.Size(784, 562);
            this.Controls.Add(this.btn_ExportReady);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.textBox_SF);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox_Carrier);
            this.Controls.Add(this.dataGridViewLinked);
            this.Name = "FormMatch";
            this.Text = "Link";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.FormMatch_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_SF)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_Carrier)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewLinked)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

            }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView_SF;
        private System.Windows.Forms.DataGridView dataGridView_Carrier;
        private System.Windows.Forms.DataGridView dataGridViewLinked;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox_SF;
        private System.Windows.Forms.TextBox textBox_Carrier;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem unbindToolStripMenuItem;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_SP_Phone;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnMember_ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Enrollment;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnCarrName;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Phone;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnSf_MemberID;
        private System.Windows.Forms.DataGridViewTextBoxColumn First_Name;
        private System.Windows.Forms.DataGridViewTextBoxColumn Trans_Date;
        private System.Windows.Forms.DataGridViewTextBoxColumn Agent_nameFull;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_SF_Product_Name;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_SF_Payroll_date;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Member_ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnPhone;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Product_Name;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Termination_Date;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnPayroll_date_ca;
        private System.Windows.Forms.Button btn_ExportReady;
        private System.ComponentModel.BackgroundWorker backgroundWorkerFillExcel;
        }
    }