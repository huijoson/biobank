namespace BioBank
{
    partial class LogIn
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnLogin = new System.Windows.Forms.Button();
            this.txtPWD = new System.Windows.Forms.TextBox();
            this.txtID = new System.Windows.Forms.TextBox();
            this.lnklblModPwd = new System.Windows.Forms.LinkLabel();
            this.txtNewPwdVer = new System.Windows.Forms.TextBox();
            this.lblNewPwd = new System.Windows.Forms.Label();
            this.lblNewPwdVer = new System.Windows.Forms.Label();
            this.btnSavePwd = new System.Windows.Forms.Button();
            this.lblAlarmId = new System.Windows.Forms.Label();
            this.lblAlarmPwd = new System.Windows.Forms.Label();
            this.lblAlarm2 = new System.Windows.Forms.Label();
            this.lblAlarm = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtNewPwd = new System.Windows.Forms.TextBox();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.SteelBlue;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(14, 88);
            this.label1.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 28);
            this.label1.TabIndex = 0;
            this.label1.Text = "密碼";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.SteelBlue;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.label2.ForeColor = System.Drawing.Color.White;
            this.label2.Location = new System.Drawing.Point(14, 21);
            this.label2.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 29);
            this.label2.TabIndex = 0;
            this.label2.Text = "帳號";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(115, 68);
            this.label3.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(220, 37);
            this.label3.TabIndex = 0;
            this.label3.Text = "生物資料庫系統";
            // 
            // btnLogin
            // 
            this.btnLogin.BackColor = System.Drawing.Color.White;
            this.btnLogin.Location = new System.Drawing.Point(114, 149);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(75, 30);
            this.btnLogin.TabIndex = 3;
            this.btnLogin.Text = "登入";
            this.btnLogin.UseVisualStyleBackColor = true;
            this.btnLogin.Click += new System.EventHandler(this.buttonLogIn_Click);
            // 
            // txtPWD
            // 
            this.txtPWD.Location = new System.Drawing.Point(114, 87);
            this.txtPWD.Name = "txtPWD";
            this.txtPWD.Size = new System.Drawing.Size(179, 29);
            this.txtPWD.TabIndex = 2;
            this.txtPWD.UseSystemPasswordChar = true;
            this.txtPWD.TextChanged += new System.EventHandler(this.txtPWD_TextChanged);
            this.txtPWD.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPWD_KeyPress);
            // 
            // txtID
            // 
            this.txtID.Location = new System.Drawing.Point(114, 21);
            this.txtID.Name = "txtID";
            this.txtID.Size = new System.Drawing.Size(179, 29);
            this.txtID.TabIndex = 1;
            this.txtID.TextChanged += new System.EventHandler(this.txtID_TextChanged);
            // 
            // lnklblModPwd
            // 
            this.lnklblModPwd.AutoSize = true;
            this.lnklblModPwd.Location = new System.Drawing.Point(25, 154);
            this.lnklblModPwd.Name = "lnklblModPwd";
            this.lnklblModPwd.Size = new System.Drawing.Size(73, 20);
            this.lnklblModPwd.TabIndex = 3;
            this.lnklblModPwd.TabStop = true;
            this.lnklblModPwd.Text = "修改密碼";
            this.lnklblModPwd.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnklblModPwd_LinkClicked);
            // 
            // txtNewPwdVer
            // 
            this.txtNewPwdVer.Font = new System.Drawing.Font("微軟正黑體", 12F);
            this.txtNewPwdVer.Location = new System.Drawing.Point(114, 221);
            this.txtNewPwdVer.Name = "txtNewPwdVer";
            this.txtNewPwdVer.Size = new System.Drawing.Size(179, 29);
            this.txtNewPwdVer.TabIndex = 4;
            this.txtNewPwdVer.UseSystemPasswordChar = true;
            this.txtNewPwdVer.Visible = false;
            // 
            // lblNewPwd
            // 
            this.lblNewPwd.BackColor = System.Drawing.Color.SteelBlue;
            this.lblNewPwd.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblNewPwd.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold);
            this.lblNewPwd.ForeColor = System.Drawing.Color.White;
            this.lblNewPwd.Location = new System.Drawing.Point(14, 150);
            this.lblNewPwd.Name = "lblNewPwd";
            this.lblNewPwd.Size = new System.Drawing.Size(100, 28);
            this.lblNewPwd.TabIndex = 7;
            this.lblNewPwd.Text = "新密碼";
            this.lblNewPwd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblNewPwd.Visible = false;
            // 
            // lblNewPwdVer
            // 
            this.lblNewPwdVer.BackColor = System.Drawing.Color.SteelBlue;
            this.lblNewPwdVer.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblNewPwdVer.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblNewPwdVer.ForeColor = System.Drawing.Color.White;
            this.lblNewPwdVer.Location = new System.Drawing.Point(14, 222);
            this.lblNewPwdVer.Name = "lblNewPwdVer";
            this.lblNewPwdVer.Size = new System.Drawing.Size(100, 28);
            this.lblNewPwdVer.TabIndex = 8;
            this.lblNewPwdVer.Text = "新密碼確認";
            this.lblNewPwdVer.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblNewPwdVer.Visible = false;
            // 
            // btnSavePwd
            // 
            this.btnSavePwd.BackColor = System.Drawing.Color.DeepSkyBlue;
            this.btnSavePwd.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnSavePwd.Location = new System.Drawing.Point(114, 283);
            this.btnSavePwd.Name = "btnSavePwd";
            this.btnSavePwd.Size = new System.Drawing.Size(75, 35);
            this.btnSavePwd.TabIndex = 5;
            this.btnSavePwd.Text = "確認";
            this.btnSavePwd.UseVisualStyleBackColor = true;
            this.btnSavePwd.Visible = false;
            this.btnSavePwd.Click += new System.EventHandler(this.btnSavePwd_Click);
            // 
            // lblAlarmId
            // 
            this.lblAlarmId.AutoSize = true;
            this.lblAlarmId.Font = new System.Drawing.Font("微軟正黑體", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblAlarmId.ForeColor = System.Drawing.Color.Red;
            this.lblAlarmId.Location = new System.Drawing.Point(376, 157);
            this.lblAlarmId.MaximumSize = new System.Drawing.Size(90, 0);
            this.lblAlarmId.Name = "lblAlarmId";
            this.lblAlarmId.Size = new System.Drawing.Size(51, 15);
            this.lblAlarmId.TabIndex = 12;
            this.lblAlarmId.Text = "無此帳號";
            this.lblAlarmId.Visible = false;
            // 
            // lblAlarmPwd
            // 
            this.lblAlarmPwd.AutoSize = true;
            this.lblAlarmPwd.Font = new System.Drawing.Font("微軟正黑體", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblAlarmPwd.ForeColor = System.Drawing.Color.Red;
            this.lblAlarmPwd.Location = new System.Drawing.Point(377, 222);
            this.lblAlarmPwd.MaximumSize = new System.Drawing.Size(90, 0);
            this.lblAlarmPwd.Name = "lblAlarmPwd";
            this.lblAlarmPwd.Size = new System.Drawing.Size(51, 15);
            this.lblAlarmPwd.TabIndex = 13;
            this.lblAlarmPwd.Text = "密碼錯誤";
            this.lblAlarmPwd.Visible = false;
            // 
            // lblAlarm2
            // 
            this.lblAlarm2.AutoSize = true;
            this.lblAlarm2.Font = new System.Drawing.Font("微軟正黑體", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblAlarm2.ForeColor = System.Drawing.Color.Red;
            this.lblAlarm2.Location = new System.Drawing.Point(377, 346);
            this.lblAlarm2.MaximumSize = new System.Drawing.Size(90, 0);
            this.lblAlarm2.Name = "lblAlarm2";
            this.lblAlarm2.Size = new System.Drawing.Size(90, 45);
            this.lblAlarm2.TabIndex = 11;
            this.lblAlarm2.Text = "密碼長度需大於8且含有英文及數字";
            this.lblAlarm2.Visible = false;
            // 
            // lblAlarm
            // 
            this.lblAlarm.AutoSize = true;
            this.lblAlarm.Font = new System.Drawing.Font("微軟正黑體", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lblAlarm.ForeColor = System.Drawing.Color.Red;
            this.lblAlarm.Location = new System.Drawing.Point(377, 273);
            this.lblAlarm.MaximumSize = new System.Drawing.Size(90, 0);
            this.lblAlarm.Name = "lblAlarm";
            this.lblAlarm.Size = new System.Drawing.Size(90, 45);
            this.lblAlarm.TabIndex = 10;
            this.lblAlarm.Text = "密碼長度需大於8且含有英文及數字";
            this.lblAlarm.Visible = false;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.txtNewPwd);
            this.panel1.Controls.Add(this.btnExit);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.btnSavePwd);
            this.panel1.Controls.Add(this.lblNewPwdVer);
            this.panel1.Controls.Add(this.lblNewPwd);
            this.panel1.Controls.Add(this.txtNewPwdVer);
            this.panel1.Controls.Add(this.lnklblModPwd);
            this.panel1.Controls.Add(this.txtID);
            this.panel1.Controls.Add(this.txtPWD);
            this.panel1.Controls.Add(this.btnLogin);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(78, 125);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(302, 331);
            this.panel1.TabIndex = 14;
            // 
            // txtNewPwd
            // 
            this.txtNewPwd.Font = new System.Drawing.Font("微軟正黑體", 12F);
            this.txtNewPwd.Location = new System.Drawing.Point(114, 149);
            this.txtNewPwd.Name = "txtNewPwd";
            this.txtNewPwd.Size = new System.Drawing.Size(179, 29);
            this.txtNewPwd.TabIndex = 11;
            this.txtNewPwd.UseSystemPasswordChar = true;
            this.txtNewPwd.Visible = false;
            // 
            // btnExit
            // 
            this.btnExit.BackColor = System.Drawing.Color.White;
            this.btnExit.Location = new System.Drawing.Point(218, 150);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 28);
            this.btnExit.TabIndex = 10;
            this.btnExit.Text = "離開";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.Color.DeepSkyBlue;
            this.btnCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnCancel.Location = new System.Drawing.Point(218, 283);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 35);
            this.btnCancel.TabIndex = 9;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // LogIn
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightBlue;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(478, 482);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.lblAlarmPwd);
            this.Controls.Add(this.lblAlarmId);
            this.Controls.Add(this.lblAlarm2);
            this.Controls.Add(this.lblAlarm);
            this.Controls.Add(this.label3);
            this.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.Name = "LogIn";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "生物資料庫";
            this.Load += new System.EventHandler(this.LogIn_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnLogin;
        private System.Windows.Forms.TextBox txtPWD;
        private System.Windows.Forms.TextBox txtID;
        private System.Windows.Forms.LinkLabel lnklblModPwd;
        private System.Windows.Forms.TextBox txtNewPwdVer;
        private System.Windows.Forms.Label lblNewPwd;
        private System.Windows.Forms.Label lblNewPwdVer;
        private System.Windows.Forms.Button btnSavePwd;
        private System.Windows.Forms.Label lblAlarmId;
        private System.Windows.Forms.Label lblAlarmPwd;
        private System.Windows.Forms.Label lblAlarm2;
        private System.Windows.Forms.Label lblAlarm;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.TextBox txtNewPwd;
        private System.Windows.Forms.Button btnExit;
    }
}

