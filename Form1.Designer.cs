namespace TestExecution
{
    partial class TestExecutionForm
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
            this.buttonUpdateExcel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonUpdateXml = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonUpdateExcel
            // 
            this.buttonUpdateExcel.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.buttonUpdateExcel.Location = new System.Drawing.Point(13, 246);
            this.buttonUpdateExcel.Margin = new System.Windows.Forms.Padding(4);
            this.buttonUpdateExcel.Name = "buttonUpdateExcel";
            this.buttonUpdateExcel.Size = new System.Drawing.Size(180, 62);
            this.buttonUpdateExcel.TabIndex = 0;
            this.buttonUpdateExcel.Text = "Upload Excel";
            this.buttonUpdateExcel.UseVisualStyleBackColor = true;
            this.buttonUpdateExcel.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 28.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label1.Location = new System.Drawing.Point(74, 60);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(347, 54);
            this.label1.TabIndex = 1;
            this.label1.Text = "Test Execution";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // buttonUpdateXml
            // 
            this.buttonUpdateXml.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.buttonUpdateXml.Location = new System.Drawing.Point(297, 246);
            this.buttonUpdateXml.Margin = new System.Windows.Forms.Padding(4);
            this.buttonUpdateXml.Name = "buttonUpdateXml";
            this.buttonUpdateXml.Size = new System.Drawing.Size(180, 62);
            this.buttonUpdateXml.TabIndex = 2;
            this.buttonUpdateXml.Text = "Update Xml";
            this.buttonUpdateXml.UseVisualStyleBackColor = true;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(14, 141);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(464, 23);
            this.progressBar1.TabIndex = 3;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.ForeColor = System.Drawing.Color.ForestGreen;
            this.lblStatus.Location = new System.Drawing.Point(167, 177);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(144, 20);
            this.lblStatus.TabIndex = 4;
            this.lblStatus.Text = "Processing...0%";
            // 
            // TestExecutionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Info;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.ClientSize = new System.Drawing.Size(490, 364);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.buttonUpdateXml);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonUpdateExcel);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "TestExecutionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "TestExecutionForm";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonUpdateExcel;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonUpdateXml;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblStatus;
    }
}

