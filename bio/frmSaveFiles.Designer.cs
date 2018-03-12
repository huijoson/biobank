namespace BioBank
{
    partial class frmSaveFiles
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
            this.txtPath = new System.Windows.Forms.TextBox();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.btnPath = new System.Windows.Forms.Button();
            this.btnExlExp = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(12, 12);
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(288, 22);
            this.txtPath.TabIndex = 0;
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(12, 43);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(288, 22);
            this.txtFileName.TabIndex = 1;
            // 
            // btnPath
            // 
            this.btnPath.Location = new System.Drawing.Point(307, 12);
            this.btnPath.Name = "btnPath";
            this.btnPath.Size = new System.Drawing.Size(75, 23);
            this.btnPath.TabIndex = 2;
            this.btnPath.Text = "瀏覽資料夾";
            this.btnPath.UseVisualStyleBackColor = true;
            this.btnPath.Click += new System.EventHandler(this.btnPath_Click);
            // 
            // btnExlExp
            // 
            this.btnExlExp.Location = new System.Drawing.Point(307, 42);
            this.btnExlExp.Name = "btnExlExp";
            this.btnExlExp.Size = new System.Drawing.Size(75, 23);
            this.btnExlExp.TabIndex = 3;
            this.btnExlExp.Text = "存檔";
            this.btnExlExp.UseVisualStyleBackColor = true;
            this.btnExlExp.Click += new System.EventHandler(this.btnExlExp_Click);
            // 
            // frmSaveFiles
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(391, 76);
            this.Controls.Add(this.btnExlExp);
            this.Controls.Add(this.btnPath);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.txtPath);
            this.Name = "frmSaveFiles";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmSaveFiles";
            this.Load += new System.EventHandler(this.frmSaveFiles_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Button btnPath;
        private System.Windows.Forms.Button btnExlExp;
    }
}