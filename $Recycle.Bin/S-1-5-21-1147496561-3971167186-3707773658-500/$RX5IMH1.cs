﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Diagnostics;

namespace BioBank
{
    public partial class LogIn : Form
    {
        
        public LogIn()
        {
            InitializeComponent();
            txtID.Validating += new CancelEventHandler(txtID_Validating);
            txtPWD.Validating += new CancelEventHandler(txtPWD_Validating);
            txtNewPwd.Validating += new CancelEventHandler(txtNewPwd_Validating);
            txtNewPwdVer.Validating += new CancelEventHandler(txtNewPwdVer_Validating);
        }
        public static string GetMD5(string original)
        {
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            byte[] b = md5.ComputeHash(Encoding.UTF8.GetBytes(original));
            return BitConverter.ToString(b).Replace("-", string.Empty);
        }

        private void LoginSuccess(string sType,string sID, string sName)
        {
            string ID = "";
            ID = sID;
            Name = sName;
            BioBank NewFrm = new BioBank();
            NewFrm.Text = "【" + sType + "】 ID: " + sID + " Name: " + sName;
            NewFrm.Show();
            this.Hide();
        }

        private void buttonLogIn_Click(object sender, EventArgs e)
        {
            string sID;
            string sName;
            string sPWD;
            string sSQL;
            string sCorrectPwd;
            string sType = "";
            sID = ""; sPWD = ""; sSQL = ""; sName = ""; sCorrectPwd = "";

            try
            {
                sID = txtID.Text;
                sPWD = txtPWD.Text;

                /*1.check Administrator 中是否有帳號*/
                //using (SqlConnection sCon = new SqlConnection(ClsShareFunc.DB_SECConnection()))
                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                {
                    sCon.Open();
                    sSQL = " select * from BioAdministratorKeyTbl (nolock) where chUserID = '" + sID + "' ";
                    SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                    SqlDataReader sRead = sCmd.ExecuteReader();

                    if (sRead.HasRows == true)
                    {
                        while (sRead.Read())
                        {
                            sCorrectPwd = ClsShareFunc.gfunCheck(sRead["chAdministratorKey"]);
                            sName = ClsShareFunc.gfunCheck(sRead["chUserName"]);
                            sType = ClsShareFunc.gfunCheck(sRead["chBioEmpFlag"]);
                        }
                        sRead.Close();

                        if (sCorrectPwd == GetMD5(sPWD))
                            LoginSuccess("Administrator (" + sType+")", sID, sName);
                    }
                    else /*2.Administrator中沒有就去Common中查*/
                    {
                        string sSQL2 = "";
                        string sEnable = "";
                        //using (SqlConnection sCon2 = new SqlConnection(ClsShareFunc.DB_SECConnection()))
                        using (SqlConnection sCon2 = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                        {
                            sCon2.Open();
                            sSQL2 = " select * from BioCommonLoginTbl (nolock) where chUserID = '" + sID + "' ";
                            SqlCommand sCmd2 = new SqlCommand(sSQL2, sCon2);
                            SqlDataReader sRead2 = sCmd2.ExecuteReader();
                            if (sRead2.HasRows == true)
                            {
                                while (sRead2.Read())
                                {
                                    sCorrectPwd = ClsShareFunc.gfunCheck(sRead2["chPassword"]);
                                    sName = ClsShareFunc.gfunCheck(sRead2["chUserName"]);
                                    sEnable = ClsShareFunc.gfunCheck(sRead2["chEnableFlag"]);
                                    sType = ClsShareFunc.gfunCheck(sRead2["chBioEmpFlag"]);
                                }
                                sRead2.Close();

                                /*enable = 'Y' -> 可使用 enable = 'N' -> 不可使用*/
                                if (sEnable == "Y")
                                {
                                    if (sCorrectPwd == GetMD5(sPWD))
                                        LoginSuccess("Common (" + sType + ")", sID, sName);
                                }
                                else
                                {
                                    MessageBox.Show("此帳號無使用權限!");
                                    return;//exit function
                                }
                            }
                            else/*Administrator和Common中皆無此帳號*/
                            {
                                MessageBox.Show("查無此帳號!");
                                return;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("登入(buttonLogIn_Click) : " + ex.Message.ToString());
                return;
            }

        }

        private void lnklblModPwd_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //btnLogin.Top = 478;209, 472
            //btnLogin.Left = 270;
            lnklblModPwd.Visible = false;
            btnLogin.Visible = false;
            lblNewPwd.Visible = true;
            lblNewPwdVer.Visible = true;
            txtNewPwd.Visible = true;
            txtNewPwdVer.Visible = true;
            btnSavePwd.Visible = true;
            btnCancel.Visible = true;

            txtNewPwd.Text = "";
            txtNewPwdVer.Text = "";
        }

        private void btnSavePwd_Click(object sender, EventArgs e)
        {
            string ID = txtID.Text;
            string sPwd = "";
            string sPwdVer = "";
            sPwd = txtNewPwd.Text;
            sPwdVer = txtNewPwdVer.Text;

            /*1.帳號不為空白*/
            if (ID != "")
            {
                    /*3.密碼輸入相同*/
                    if (sPwd == sPwdVer)
                    {
                        if (VerAction("修改") == false)
                            return;

                            /*4.帳號存在Administrator db*/
                            if (ClsShareFunc.CheckInDb(ClsShareFunc.DbAdmin(), ID, "modify") == true)
                            {
                                /*5.更新密碼*/
                                //using (SqlConnection updateCon = new SqlConnection(ClsShareFunc.DB_SECConnection()))
                                using (SqlConnection updateCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                                {
                                    updateCon.Open();
                                    SqlCommand updateCmd = new SqlCommand("update BioAdministratorKeyTbl " +
                                        "set chAdministratorKey = '" + GetMD5(sPwdVer) + "' where chUserId = '" + ID + "' ", updateCon);
                                    updateCmd.ExecuteNonQuery();

                                    MessageBox.Show("密碼修改成功!請重新登入。");
                                    InitFrm();
                                    updateCon.Close();
                                    updateCon.Dispose();                               
                                }
                            }
                            else
                            {
                                /*4.帳號存在 Common db*/
                                if (ClsShareFunc.CheckInDb(ClsShareFunc.DbCom(), ID, "modify") == true)
                                {
                                    /*5.更新密碼*/
                                    //using (SqlConnection updateCon = new SqlConnection(ClsShareFunc.DB_SECConnection()))
                                    using (SqlConnection updateCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                                    {
                                        updateCon.Open();
                                        SqlCommand updateCmd = new SqlCommand("update BioCommonLoginTbl " +
                                            "set chPassword = '" + GetMD5(sPwdVer) + "' where chUserId = '" + ID + "' ", updateCon);
                                        updateCmd.ExecuteNonQuery();

                                        MessageBox.Show("密碼修改成功!請重新登入。");
                                        InitFrm();
                                        updateCon.Close();
                                        updateCon.Dispose();
                                    }
                                }
                                else
                                    MessageBox.Show("查無此帳號!");
                            }
                        }
                        else
                            MessageBox.Show("密碼不一致。請重新輸入!");
            }
            else
                MessageBox.Show("請先登入!");
            txtNewPwd.Text = "";
            txtNewPwdVer.Text = "";
        }

        private void InitFrm()
        {
            lnklblModPwd.Visible = true;
            btnLogin.Visible = true;
            btnSavePwd.Visible = false;
            lblNewPwd.Visible = false;
            lblNewPwdVer.Visible = false;
            txtNewPwd.Visible = false;
            txtNewPwdVer.Visible = false;
            btnCancel.Visible = false;

            txtID.Text = "";
            txtPWD.Text = "";
        }

        /*跳出確認視窗*/
        private Boolean VerAction(string Str)
        {
            Boolean Check = false;
            string message = "確定" + Str + "?";
            string caption = Str;
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result;

            result = MessageBox.Show(this, message, caption, buttons);
            if (result == DialogResult.Yes)
            { 
                //this.Close();
                Check = true;
            }
            else
                Check = false;

            return Check;
        }

        /*驗證txtPwd*/
        private void txtPWD_Validating(object sender, CancelEventArgs e)
        {
            lblAlarmPwd.Visible = false;
            txtNewPwd.Enabled = true;
            txtNewPwdVer.Enabled = true;
            btnSavePwd.Enabled = true;

            string sPwd = "";
            string sId = "";
            sId = txtID.Text;
            sPwd = txtPWD.Text;

            if (ClsShareFunc.CheckInDb(ClsShareFunc.DbAdmin(), sId, "modify") == false)
            {
                if (ClsShareFunc.CheckInDb(ClsShareFunc.DbCom(), sId, "modify") == true)
                {
                    //Common有
                    string sSQL = "";
                    //using (SqlConnection sCon = new SqlConnection(ClsShareFunc.DB_SECConnection()))
                    using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                    {
                        sCon.Open();
                        sSQL = "select * from BioCommonLoginTbl (nolock) where chUserId = '" + sId + "' ";
                        SqlCommand sCmd2 = new SqlCommand(sSQL, sCon);
                        SqlDataReader sRead2 = sCmd2.ExecuteReader();
                        if (sRead2.HasRows == true)
                        {
                            while (sRead2.Read())
                            {
                                sPwd = ClsShareFunc.gfunCheck(sRead2["chPassword"]).ToString().Trim();
                            }
                        }
                        sRead2.Close();
                        sCon.Dispose();
                    }

                    if (sPwd != GetMD5(txtPWD.Text))
                    {
                        //密碼錯誤
                        lblAlarmPwd.Visible = true;
                        txtNewPwd.Enabled = false;
                        txtNewPwdVer.Enabled = false;
                        btnSavePwd.Enabled = false;
                    }
                }
            }
            else
            {
                //Administrator有
                string sSQL2 = "";
                //using (SqlConnection sCon = new SqlConnection(ClsShareFunc.DB_SECConnection()))
                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                {
                    sCon.Open();
                    sSQL2 = "select * from BioAdministratorKeyTbl (nolock) where chUserId = '" + sId + "' ";
                    SqlCommand sCmd = new SqlCommand(sSQL2, sCon);
                    SqlDataReader sRead = sCmd.ExecuteReader();
                    if (sRead.HasRows == true)
                    {
                        while (sRead.Read())
                        {
                            sPwd = ClsShareFunc.gfunCheck(sRead["chAdministratorKey"]).ToString().Trim();
                        }
                    }
                    sRead.Close();
                }

                if (sPwd != GetMD5(txtPWD.Text))
                {
                    //密碼錯誤
                    lblAlarmPwd.Visible = true;
                    txtNewPwd.Enabled = false;
                    txtNewPwdVer.Enabled = false;
                    btnSavePwd.Enabled = false;
                }
            }
        }

        /*驗證txtNewPwd*/
        private void txtNewPwd_Validating(object sender, CancelEventArgs e)
        {
            lblAlarm.Visible = false;
            btnSavePwd.Enabled = true;

            string sPwd = "";
            sPwd = txtNewPwd.Text;

                /*.驗證正確性*/
                if (ClsShareFunc.gfunCheckPwd(sPwd) == false)
                {
                    lblAlarm.Visible = true;
                    txtNewPwd.Text = "";
                    btnSavePwd.Enabled = false;
                }
                else
                {
                    lblAlarm.Visible = false;
                    //btnSavePwd.Enabled = true;
                }

        }

        /*驗證txtNewPwdVer*/
        private void txtNewPwdVer_Validating(object sender, CancelEventArgs e)
        {
            string sPwdVer = "";
            sPwdVer = txtNewPwdVer.Text;

                /*.驗證正確性*/
                if (sPwdVer != "" && ClsShareFunc.gfunCheckPwd(sPwdVer) == false)
                {
                    lblAlarm2.Visible = true;
                    txtNewPwdVer.Text = "";
                    btnSavePwd.Enabled = false;
                }
                else
                {
                    lblAlarm2.Visible = false;
                    btnSavePwd.Enabled = true;
                }
        }

        /*驗證帳號是否存在*/
        private void txtID_Validating(object sender, CancelEventArgs e)
        {
            lblAlarmId.Visible = false;
            txtPWD.Enabled = true;
            txtNewPwd.Enabled = true;
            txtNewPwdVer.Enabled = true;
            btnSavePwd.Enabled = true;

            string sId = "";
            sId = txtID.Text;

            if (sId != "")
            {
                if (ClsShareFunc.CheckInDb(ClsShareFunc.DbAdmin(), sId, "modify") == false)
                {
                    if (ClsShareFunc.CheckInDb(ClsShareFunc.DbCom(), sId, "modify") == false)
                    {
                        //兩個db皆無
                        lblAlarmId.Visible = true;
                        txtPWD.Enabled = false;
                        txtNewPwd.Enabled = false;
                        txtNewPwdVer.Enabled = false;
                        btnSavePwd.Enabled = false;
                    }
                }
            }

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            InitFrm();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
            Environment.Exit(Environment.ExitCode);

            InitializeComponent();
        }

        private void btnRestore_Click(object sender, EventArgs e)
        {
            Process myProcess = new Process();
            myProcess.StartInfo.FileName = @"C:\Batch2\RestoreBio.bat";
            myProcess.StartInfo.UseShellExecute = false;
            myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            // Wait for the sort process to write the sorted text lines.
            myProcess.Start();
            myProcess.WaitForExit();
            MessageBox.Show("還原完成!");
        }

        private void txtID_TextChanged(object sender, EventArgs e)
        {

        }

        private void LogIn_Load(object sender, EventArgs e)
        {
            using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
            {
                try
                {
                    sCon.Open();
                    SqlCommand updateCmd = new SqlCommand("alter user biorest with login = biorest", sCon);
                    updateCmd.ExecuteNonQuery();

                    updateCmd = new SqlCommand("use DB_BIO alter user biobank with login = biobank", sCon);
                    updateCmd.ExecuteNonQuery();
                    updateCmd = new SqlCommand("use DB_BIO alter user biorest with login = biorest", sCon);
                    updateCmd.ExecuteNonQuery();                  
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Login Form Load's Error : " + ex.Message.ToString());
                    return;
                }



            }
        }



    }
}
