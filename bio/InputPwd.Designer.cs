﻿namespace BioBank
{
    partial class frmVerPwd
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
            this.lblVerPwd = new System.Windows.Forms.Label();
            this.txtPWD = new System.Windows.Forms.TextBox();
            this.lblRemind = new System.Windows.Forms.Label();
            this.btnVer = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblID = new System.Windows.Forms.Label();
            this.txtID = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // lblVerPwd
            // 
            this.lblVerPwd.BackColor = System.Drawing.Color.DimGray;
            this.lblVerPwd.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVerPwd.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblVerPwd.ForeColor = System.Drawing.Color.White;
            this.lblVerPwd.Location = new System.Drawing.Point(12, 107);
            this.lblVerPwd.Name = "lblVerPwd";
            this.lblVerPwd.Size = new System.Drawing.Size(86, 29);
            this.lblVerPwd.TabIndex = 0;
            this.lblVerPwd.Text = "密碼確認:";
            this.lblVerPwd.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // txtPWD
            // 
            this.txtPWD.Font = new System.Drawing.Font("微軟正黑體", 12F);
            this.txtPWD.Location = new System.Drawing.Point(97, 107);
            this.txtPWD.Name = "txtPWD";
            this.txtPWD.Size = new System.Drawing.Size(169, 29);
            this.txtPWD.TabIndex = 3;
            this.txtPWD.UseSystemPasswordChar = true;
            this.txtPWD.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPWD_KeyPress);
            // 
            // lblRemind
            // 
            this.lblRemind.AutoSize = true;
            this.lblRemind.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblRemind.ForeColor = System.Drawing.Color.Black;
            this.lblRemind.Location = new System.Drawing.Point(12, 38);
            this.lblRemind.Name = "lblRemind";
            this.lblRemind.Size = new System.Drawing.Size(0, 21);
            this.lblRemind.TabIndex = 4;
            // 
            // btnVer
            // 
            this.btnVer.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnVer.Location = new System.Drawing.Point(97, 175);
            this.btnVer.Name = "btnVer";
            this.btnVer.Size = new System.Drawing.Size(74, 30);
            this.btnVer.TabIndex = 5;
            this.btnVer.Text = "確認";
            this.btnVer.UseVisualStyleBackColor = true;
            this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnCancel.Location = new System.Drawing.Point(187, 175);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(79, 30);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblID
            // 
            this.lblID.BackColor = System.Drawing.Color.DimGray;
            this.lblID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblID.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblID.ForeColor = System.Drawing.Color.White;
            this.lblID.Location = new System.Drawing.Point(10, 38);
            this.lblID.Name = "lblID";
            this.lblID.Size = new System.Drawing.Size(86, 29);
            this.lblID.TabIndex = 7;
            this.lblID.Text = "ID :";
            this.lblID.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.lblID.Visible = false;
            // 
            // txtID
            // 
            this.txtID.Font = new System.Drawing.Font("微軟正黑體", 12F);
            this.txtID.Location = new System.Drawing.Point(97, 38);
            this.txtID.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtID.Name = "txtID";
            this.txtID.Size = new System.Drawing.Size(169, 29);
            this.txtID.TabIndex = 8;
            this.txtID.Visible = false;
            // 
            // frmVerPwd
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.RosyBrown;
            this.ClientSize = new System.Drawing.Size(314, 217);
            this.ControlBox = false;
            this.Controls.Add(this.txtID);
            this.Controls.Add(this.lblID);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnVer);
            this.Controls.Add(this.lblRemind);
            this.Controls.Add(this.txtPWD);
            this.Controls.Add(this.lblVerPwd);
            this.Name = "frmVerPwd";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "密碼確認";
            this.Load += new System.EventHandler(this.frmVerPwd_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblVerPwd;
        private System.Windows.Forms.TextBox txtPWD;
        private System.Windows.Forms.Label lblRemind;
        private System.Windows.Forms.Button btnVer;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblID;
        private System.Windows.Forms.TextBox txtID;
    }
}