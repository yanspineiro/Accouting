namespace winform_SF_Lib
    {
    partial class ReCalculatePolicyForm
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
            this.First_Name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Product_Name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Agent_nameFull = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewupdate = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.backgroundWorker_Search = new System.ComponentModel.BackgroundWorker();
            this.backgroundWorker_update = new System.ComponentModel.BackgroundWorker();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_HII)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewupdate)).BeginInit();
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
            this.First_Name,
            this.Product_Name,
            this.Agent_nameFull});
            this.dataGridView_HII.Location = new System.Drawing.Point(13, 27);
            this.dataGridView_HII.Name = "dataGridView_HII";
            this.dataGridView_HII.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView_HII.Size = new System.Drawing.Size(449, 490);
            this.dataGridView_HII.TabIndex = 14;
            // 
            // ColumnSf_MemberID
            // 
            this.ColumnSf_MemberID.DataPropertyName = "Sf_MemberID";
            this.ColumnSf_MemberID.HeaderText = "Policy";
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
            // Product_Name
            // 
            this.Product_Name.DataPropertyName = "Product_Name";
            this.Product_Name.FillWeight = 10F;
            this.Product_Name.HeaderText = "Product Name";
            this.Product_Name.Name = "Product_Name";
            this.Product_Name.ReadOnly = true;
            this.Product_Name.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
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
            // dataGridViewupdate
            // 
            this.dataGridViewupdate.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewupdate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewupdate.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4});
            this.dataGridViewupdate.Location = new System.Drawing.Point(493, 27);
            this.dataGridViewupdate.Name = "dataGridViewupdate";
            this.dataGridViewupdate.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewupdate.Size = new System.Drawing.Size(454, 490);
            this.dataGridViewupdate.TabIndex = 15;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "Sf_MemberID";
            this.dataGridViewTextBoxColumn1.HeaderText = "Policy";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
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
            this.dataGridViewTextBoxColumn3.DataPropertyName = "Product_Name";
            this.dataGridViewTextBoxColumn3.FillWeight = 10F;
            this.dataGridViewTextBoxColumn3.HeaderText = "Product Name";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.ReadOnly = true;
            this.dataGridViewTextBoxColumn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "Agent_Fullname";
            this.dataGridViewTextBoxColumn4.FillWeight = 10F;
            this.dataGridViewTextBoxColumn4.HeaderText = "Agent full name";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.ReadOnly = true;
            this.dataGridViewTextBoxColumn4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(735, 531);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 16;
            this.button1.Text = "Update";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(826, 531);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 17;
            this.button2.Text = "Close";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(643, 531);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 18;
            this.button3.Text = "Search";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // backgroundWorker_Search
            // 
            this.backgroundWorker_Search.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_Search_DoWork);
            this.backgroundWorker_Search.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_Search_RunWorkerCompleted);
            // 
            // ReCalculatePolicyForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(959, 566);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridViewupdate);
            this.Controls.Add(this.dataGridView_HII);
            this.Name = "ReCalculatePolicyForm";
            this.Text = "Search Change in Policy number";
            this.Load += new System.EventHandler(this.ReCalculatePolicyForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_HII)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewupdate)).EndInit();
            this.ResumeLayout(false);

            }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView_HII;
        private System.Windows.Forms.DataGridView dataGridViewupdate;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnSf_MemberID;
        private System.Windows.Forms.DataGridViewTextBoxColumn First_Name;
        private System.Windows.Forms.DataGridViewTextBoxColumn Product_Name;
        private System.Windows.Forms.DataGridViewTextBoxColumn Agent_nameFull;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.Button button3;
        private System.ComponentModel.BackgroundWorker backgroundWorker_Search;
        private System.ComponentModel.BackgroundWorker backgroundWorker_update;
        }
    }