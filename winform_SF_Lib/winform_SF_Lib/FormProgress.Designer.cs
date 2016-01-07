namespace SF_Lib
{
    partial class FormProgress
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
            this.panel_progress = new System.Windows.Forms.Panel();
            this.Label_infoFile = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.panel_progress.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel_progress
            // 
            this.panel_progress.Controls.Add(this.Label_infoFile);
            this.panel_progress.Controls.Add(this.label2);
            this.panel_progress.Controls.Add(this.progressBar1);
            this.panel_progress.Location = new System.Drawing.Point(-85, 4);
            this.panel_progress.Name = "panel_progress";
            this.panel_progress.Size = new System.Drawing.Size(474, 70);
            this.panel_progress.TabIndex = 8;
            // 
            // Label_infoFile
            // 
            this.Label_infoFile.AutoSize = true;
            this.Label_infoFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Label_infoFile.Location = new System.Drawing.Point(217, 50);
            this.Label_infoFile.Name = "Label_infoFile";
            this.Label_infoFile.Size = new System.Drawing.Size(102, 13);
            this.Label_infoFile.TabIndex = 8;
            this.Label_infoFile.Text = "Task in Progress";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(22, 43);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Reading  ";
            // 
            // progressBar1
            // 
            this.progressBar1.BackColor = System.Drawing.Color.White;
            this.progressBar1.Location = new System.Drawing.Point(90, 19);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(381, 23);
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar1.TabIndex = 6;
            // 
            // FormProgress
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(390, 81);
            this.ControlBox = false;
            this.Controls.Add(this.panel_progress);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(406, 97);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(406, 97);
            this.Name = "FormProgress";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.FormProgress_Load);
            this.panel_progress.ResumeLayout(false);
            this.panel_progress.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel_progress;
        private System.Windows.Forms.Label Label_infoFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar progressBar1;

    }
}