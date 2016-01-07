namespace SF_Lib
    {
    partial class Form_Manual_Insert
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
            this.dataGridView_HII = new System.Windows.Forms.DataGridView();
            this.ColumnSf_MemberID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OpportunityLine_id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Agent_nameFull = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Agent_Status_c = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnAgent_Commision = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnTrans_Type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnPayment_date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1 = new System.Windows.Forms.Panel();
            this.chb_new = new System.Windows.Forms.CheckBox();
            this.cb_PayrollDate = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.btn_delete = new System.Windows.Forms.Button();
            this.btn_Save = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.tb_Amount = new System.Windows.Forms.TextBox();
            this.tb_policynumber = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cb_commissiontype = new System.Windows.Forms.ComboBox();
            this.lb_Agent = new System.Windows.Forms.Label();
            this.cb_agent = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_HII)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView_HII
            // 
            this.dataGridView_HII.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView_HII.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_HII.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnSf_MemberID,
            this.OpportunityLine_id,
            this.Agent_nameFull,
            this.Agent_Status_c,
            this.ColumnAgent_Commision,
            this.ColumnTrans_Type,
            this.ColumnPayment_date});
            this.dataGridView_HII.Location = new System.Drawing.Point(1, 1);
            this.dataGridView_HII.MultiSelect = false;
            this.dataGridView_HII.Name = "dataGridView_HII";
            this.dataGridView_HII.ReadOnly = true;
            this.dataGridView_HII.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_HII.Size = new System.Drawing.Size(744, 382);
            this.dataGridView_HII.TabIndex = 100;
            this.dataGridView_HII.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_HII_CellClick);
            this.dataGridView_HII.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_HII_CellEnter);
            // 
            // ColumnSf_MemberID
            // 
            this.ColumnSf_MemberID.DataPropertyName = "Sf_MemberID";
            this.ColumnSf_MemberID.HeaderText = "Sf Policy";
            this.ColumnSf_MemberID.Name = "ColumnSf_MemberID";
            this.ColumnSf_MemberID.ReadOnly = true;
            this.ColumnSf_MemberID.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // OpportunityLine_id
            // 
            this.OpportunityLine_id.DataPropertyName = "OpportunityLine_id";
            this.OpportunityLine_id.FillWeight = 10F;
            this.OpportunityLine_id.HeaderText = "OppLine id";
            this.OpportunityLine_id.Name = "OpportunityLine_id";
            this.OpportunityLine_id.ReadOnly = true;
            this.OpportunityLine_id.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Agent_nameFull
            // 
            this.Agent_nameFull.DataPropertyName = "Agent_Fullname";
            this.Agent_nameFull.FillWeight = 10F;
            this.Agent_nameFull.HeaderText = "Agent full name";
            this.Agent_nameFull.Name = "Agent_nameFull";
            this.Agent_nameFull.ReadOnly = true;
            this.Agent_nameFull.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Agent_Status_c
            // 
            this.Agent_Status_c.DataPropertyName = "Agent_Status";
            this.Agent_Status_c.FillWeight = 10F;
            this.Agent_Status_c.HeaderText = "Agent Status";
            this.Agent_Status_c.Name = "Agent_Status_c";
            this.Agent_Status_c.ReadOnly = true;
            this.Agent_Status_c.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnAgent_Commision
            // 
            this.ColumnAgent_Commision.DataPropertyName = "Agent_Commision";
            this.ColumnAgent_Commision.FillWeight = 10F;
            this.ColumnAgent_Commision.HeaderText = "Comm";
            this.ColumnAgent_Commision.Name = "ColumnAgent_Commision";
            this.ColumnAgent_Commision.ReadOnly = true;
            this.ColumnAgent_Commision.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnTrans_Type
            // 
            this.ColumnTrans_Type.DataPropertyName = "Payroll_Type";
            this.ColumnTrans_Type.HeaderText = "Trans Type";
            this.ColumnTrans_Type.Name = "ColumnTrans_Type";
            this.ColumnTrans_Type.ReadOnly = true;
            this.ColumnTrans_Type.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // ColumnPayment_date
            // 
            this.ColumnPayment_date.DataPropertyName = "Payroll_date";
            this.ColumnPayment_date.HeaderText = "Payment date";
            this.ColumnPayment_date.Name = "ColumnPayment_date";
            this.ColumnPayment_date.ReadOnly = true;
            this.ColumnPayment_date.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.chb_new);
            this.panel1.Controls.Add(this.cb_PayrollDate);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.btn_delete);
            this.panel1.Controls.Add(this.btn_Save);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.tb_Amount);
            this.panel1.Controls.Add(this.tb_policynumber);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.cb_commissiontype);
            this.panel1.Controls.Add(this.lb_Agent);
            this.panel1.Controls.Add(this.cb_agent);
            this.panel1.Location = new System.Drawing.Point(1, 389);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(744, 170);
            this.panel1.TabIndex = 2;
            // 
            // chb_new
            // 
            this.chb_new.AutoSize = true;
            this.chb_new.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chb_new.Location = new System.Drawing.Point(111, 9);
            this.chb_new.Name = "chb_new";
            this.chb_new.Size = new System.Drawing.Size(75, 17);
            this.chb_new.TabIndex = 17;
            this.chb_new.Text = "Add new";
            this.chb_new.UseVisualStyleBackColor = true;
            this.chb_new.CheckedChanged += new System.EventHandler(this.chb_new_CheckedChanged);
            // 
            // cb_PayrollDate
            // 
            this.cb_PayrollDate.DisplayMember = "Payment_Date__c";
            this.cb_PayrollDate.FormattingEnabled = true;
            this.cb_PayrollDate.Location = new System.Drawing.Point(111, 91);
            this.cb_PayrollDate.Name = "cb_PayrollDate";
            this.cb_PayrollDate.Size = new System.Drawing.Size(208, 21);
            this.cb_PayrollDate.TabIndex = 3;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(12, 92);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(97, 17);
            this.label7.TabIndex = 16;
            this.label7.Text = "Payroll Date";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(665, 89);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 10;
            this.button2.Text = "Close";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btn_delete
            // 
            this.btn_delete.Location = new System.Drawing.Point(665, 63);
            this.btn_delete.Name = "btn_delete";
            this.btn_delete.Size = new System.Drawing.Size(75, 23);
            this.btn_delete.TabIndex = 9;
            this.btn_delete.Text = "Delete";
            this.btn_delete.UseVisualStyleBackColor = true;
            this.btn_delete.Click += new System.EventHandler(this.btn_delete_Click);
            // 
            // btn_Save
            // 
            this.btn_Save.Location = new System.Drawing.Point(665, 36);
            this.btn_Save.Name = "btn_Save";
            this.btn_Save.Size = new System.Drawing.Size(75, 23);
            this.btn_Save.TabIndex = 8;
            this.btn_Save.Text = "Save";
            this.btn_Save.UseVisualStyleBackColor = true;
            this.btn_Save.Click += new System.EventHandler(this.btn_Save_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(325, 66);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(93, 17);
            this.label4.TabIndex = 13;
            this.label4.Text = "Commission";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(325, 39);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(112, 17);
            this.label3.TabIndex = 12;
            this.label3.Text = "Policy Number";
            // 
            // tb_Amount
            // 
            this.tb_Amount.Location = new System.Drawing.Point(451, 65);
            this.tb_Amount.Name = "tb_Amount";
            this.tb_Amount.Size = new System.Drawing.Size(208, 20);
            this.tb_Amount.TabIndex = 5;
            this.tb_Amount.Text = "0.00";
            this.tb_Amount.TextChanged += new System.EventHandler(this.validateTextDouble);
            // 
            // tb_policynumber
            // 
            this.tb_policynumber.Location = new System.Drawing.Point(451, 38);
            this.tb_policynumber.Name = "tb_policynumber";
            this.tb_policynumber.Size = new System.Drawing.Size(208, 20);
            this.tb_policynumber.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 68);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(94, 17);
            this.label2.TabIndex = 9;
            this.label2.Text = "Payroll type";
            // 
            // cb_commissiontype
            // 
            this.cb_commissiontype.FormattingEnabled = true;
            this.cb_commissiontype.Items.AddRange(new object[] {
            "Adjusment",
            "Commission",
            "Chargeback",
            "Terminated"});
            this.cb_commissiontype.Location = new System.Drawing.Point(111, 65);
            this.cb_commissiontype.Name = "cb_commissiontype";
            this.cb_commissiontype.Size = new System.Drawing.Size(208, 21);
            this.cb_commissiontype.TabIndex = 2;
            // 
            // lb_Agent
            // 
            this.lb_Agent.AutoSize = true;
            this.lb_Agent.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_Agent.Location = new System.Drawing.Point(12, 39);
            this.lb_Agent.Name = "lb_Agent";
            this.lb_Agent.Size = new System.Drawing.Size(50, 17);
            this.lb_Agent.TabIndex = 5;
            this.lb_Agent.Text = "Agent";
            // 
            // cb_agent
            // 
            this.cb_agent.DisplayMember = "Username__c";
            this.cb_agent.FormattingEnabled = true;
            this.cb_agent.Location = new System.Drawing.Point(111, 38);
            this.cb_agent.Name = "cb_agent";
            this.cb_agent.Size = new System.Drawing.Size(208, 21);
            this.cb_agent.TabIndex = 1;
            // 
            // Form_Manual_Insert
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(748, 562);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dataGridView_HII);
            this.Name = "Form_Manual_Insert";
            this.Text = "Manual Insertion";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form_Manual_Insert_FormClosed);
            this.Load += new System.EventHandler(this.Form_Manual_Insert_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_HII)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

            }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView_HII;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lb_Agent;
        private System.Windows.Forms.ComboBox cb_agent;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tb_Amount;
        private System.Windows.Forms.TextBox tb_policynumber;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cb_commissiontype;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btn_delete;
        private System.Windows.Forms.Button btn_Save;
        private System.Windows.Forms.ComboBox cb_PayrollDate;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.CheckBox chb_new;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnSf_MemberID;
        private System.Windows.Forms.DataGridViewTextBoxColumn OpportunityLine_id;
        private System.Windows.Forms.DataGridViewTextBoxColumn Agent_nameFull;
        private System.Windows.Forms.DataGridViewTextBoxColumn Agent_Status_c;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnAgent_Commision;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnTrans_Type;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnPayment_date;
        }
    }