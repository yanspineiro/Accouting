namespace winform_SF_Lib
    {
    partial class Form_PrintReport
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
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.comboBox_Agent = new System.Windows.Forms.ComboBox();
            this.btn_Print = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.cb_Profile = new System.Windows.Forms.ComboBox();
            this.btn_close = new System.Windows.Forms.Button();
            this.cb_payment_period = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.backgroundWorkerReport = new System.ComponentModel.BackgroundWorker();
            this.backgroundWorkerCalendar = new System.ComponentModel.BackgroundWorker();
            this.backgroundWorkerGetALL = new System.ComponentModel.BackgroundWorker();
            this.backgroundWorkerSalesboard = new System.ComponentModel.BackgroundWorker();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.comboBox_Agent);
            this.groupBox2.Controls.Add(this.btn_Print);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.cb_Profile);
            this.groupBox2.Controls.Add(this.btn_close);
            this.groupBox2.Controls.Add(this.cb_payment_period);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(6, -2);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox2.Size = new System.Drawing.Size(693, 251);
            this.groupBox2.TabIndex = 19;
            this.groupBox2.TabStop = false;
            // 
            // button2
            // 
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Location = new System.Drawing.Point(532, 115);
            this.button2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(112, 35);
            this.button2.TabIndex = 26;
            this.button2.Text = "Salesboard";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button1.Location = new System.Drawing.Point(532, 71);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(112, 35);
            this.button1.TabIndex = 25;
            this.button1.Text = "All Data";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(16, 125);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(76, 25);
            this.label7.TabIndex = 24;
            this.label7.Text = "Agent:";
            // 
            // comboBox_Agent
            // 
            this.comboBox_Agent.DisplayMember = "Payment_Date__c";
            this.comboBox_Agent.FormattingEnabled = true;
            this.comboBox_Agent.Location = new System.Drawing.Point(216, 123);
            this.comboBox_Agent.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.comboBox_Agent.Name = "comboBox_Agent";
            this.comboBox_Agent.Size = new System.Drawing.Size(306, 28);
            this.comboBox_Agent.TabIndex = 23;
            // 
            // btn_Print
            // 
            this.btn_Print.Location = new System.Drawing.Point(532, 29);
            this.btn_Print.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btn_Print.Name = "btn_Print";
            this.btn_Print.Size = new System.Drawing.Size(112, 35);
            this.btn_Print.TabIndex = 22;
            this.btn_Print.Text = "Get Report";
            this.btn_Print.UseVisualStyleBackColor = true;
            this.btn_Print.Click += new System.EventHandler(this.btn_Print_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(16, 75);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(171, 25);
            this.label3.TabIndex = 7;
            this.label3.Text = "Payment Period:";
            // 
            // cb_Profile
            // 
            this.cb_Profile.FormattingEnabled = true;
            this.cb_Profile.Location = new System.Drawing.Point(216, 32);
            this.cb_Profile.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cb_Profile.Name = "cb_Profile";
            this.cb_Profile.Size = new System.Drawing.Size(306, 28);
            this.cb_Profile.TabIndex = 0;
            this.cb_Profile.SelectedIndexChanged += new System.EventHandler(this.cb_Companny_SelectedIndexChanged);
            // 
            // btn_close
            // 
            this.btn_close.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btn_close.Location = new System.Drawing.Point(532, 160);
            this.btn_close.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btn_close.Name = "btn_close";
            this.btn_close.Size = new System.Drawing.Size(112, 35);
            this.btn_close.TabIndex = 12;
            this.btn_close.Text = "Close";
            this.btn_close.UseVisualStyleBackColor = true;
            this.btn_close.Click += new System.EventHandler(this.btn_close_Click);
            // 
            // cb_payment_period
            // 
            this.cb_payment_period.DisplayMember = "Payment_Date__c";
            this.cb_payment_period.FormattingEnabled = true;
            this.cb_payment_period.Location = new System.Drawing.Point(216, 74);
            this.cb_payment_period.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cb_payment_period.Name = "cb_payment_period";
            this.cb_payment_period.Size = new System.Drawing.Size(306, 28);
            this.cb_payment_period.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(16, 34);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 25);
            this.label1.TabIndex = 1;
            this.label1.Text = "Profile";
            // 
            // backgroundWorkerReport
            // 
            this.backgroundWorkerReport.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerReport_DoWork);
            this.backgroundWorkerReport.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerReport_RunWorkerCompleted);
            // 
            // backgroundWorkerCalendar
            // 
            this.backgroundWorkerCalendar.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerCalendar_DoWork);
            this.backgroundWorkerCalendar.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerCalendar_RunWorkerCompleted);
            // 
            // backgroundWorkerGetALL
            // 
            this.backgroundWorkerGetALL.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerGetALL_DoWork);
            this.backgroundWorkerGetALL.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerGetALL_RunWorkerCompleted);
            // 
            // backgroundWorkerSalesboard
            // 
            this.backgroundWorkerSalesboard.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerSalesboard_DoWork);
            this.backgroundWorkerSalesboard.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerSalesboard_RunWorkerCompleted);
            // 
            // Form_PrintReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(694, 232);
            this.Controls.Add(this.groupBox2);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(716, 287);
            this.MinimumSize = new System.Drawing.Size(716, 287);
            this.Name = "Form_PrintReport";
            this.Text = "Generate Sheet";
            this.Load += new System.EventHandler(this.Form_PrintReport_Load);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

            }

        #endregion

        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox comboBox_Agent;
        private System.Windows.Forms.Button btn_Print;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cb_Profile;
        private System.Windows.Forms.Button btn_close;
        private System.Windows.Forms.ComboBox cb_payment_period;
        private System.Windows.Forms.Label label1;
        private System.ComponentModel.BackgroundWorker backgroundWorkerReport;
        private System.ComponentModel.BackgroundWorker backgroundWorkerCalendar;
        private System.Windows.Forms.Button button1;
        private System.ComponentModel.BackgroundWorker backgroundWorkerGetALL;
        private System.Windows.Forms.Button button2;
        private System.ComponentModel.BackgroundWorker backgroundWorkerSalesboard;
        }
    }