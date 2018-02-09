using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace BioBank
{
    public partial class frmVerPwd : Form
    {
        public frmVerPwd()
        {
            InitializeComponent();
        }

        public Boolean PassVerPwd;
        public string pEntrySource = "";

        private void btnVer_Click(object sender, EventArgs e)
        {
            string sSQL = "";
            string sPwd = "";
            string sID = "";
            string sCorrectPwd = "";
            PassVerPwd = false;

            sPwd = txtPWD.Text;
            sID = ClsShareFunc.sUserId;


            if (sPwd == "")
            {
                MessageBox.Show("請輸入密碼!");
                return;
            }
           
                //switch (ClsShareFunc.sLoginIdentity)
            if (pEntrySource == "Function6" || pEntrySource == "Function7" || pEntrySource == "Function10")
                {
                    if (ClsShareFunc.sLoginIdentity != "Administrator")
                    {                       
                        MessageBox.Show("非【生物、資訊主管-Administrator】權限，無法進入！", "Administrator Only!!!");
                        this.Close();
                    }
                    else
                    {
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
                                }
                                sRead.Close();

                                if (sCorrectPwd == ClsShareFunc.GetMD5(sPwd))
                                {
                                    PassVerPwd = true;
                                    this.Close();
                                }
                                else
                                {
                                    PassVerPwd = false;
                                    MessageBox.Show("密碼錯誤，請重新輸入！");
                                    txtPWD.Text = "";
                                }
                            }
                        }  
                    }
                }
                if (pEntrySource == "Function8")
                {
                    if (ClsShareFunc.sLoginIdentity != "Common")
                    {
                        MessageBox.Show("需先以一般行政同仁權限進入【再輔以生物、資訊主管-Administrator 權限進入】！", "行政同仁 First!!!");
                        return;
                    }
                    else
                    {
                        if (txtID.Text.Trim() == "" || txtPWD.Text.Trim() == "")
                        {
                            MessageBox.Show("ID 及 PWD不可為空白!");
                            return;
                        }
                        //using (SqlConnection sCon = new SqlConnection(ClsShareFunc.DB_SECConnection()))
                        using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                        {
                            sCon.Open();
                            sSQL = " select * from BioAdministratorKeyTbl (nolock) where chUserID = '" + txtID.Text.Trim()   + "' ";
                            SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                            SqlDataReader sRead = sCmd.ExecuteReader();

                            if (sRead.HasRows == true)
                            {
                                while (sRead.Read())
                                {
                                    sCorrectPwd = ClsShareFunc.gfunCheck(sRead["chAdministratorKey"]);
                                }
                                sRead.Close();

                                if (sCorrectPwd == ClsShareFunc.GetMD5(sPwd))
                                {
                                    PassVerPwd = true;
                                    BioBank.pFunction8_AdminID = txtID.Text.Trim();
                                    this.Close();
                                }
                                else
                                {
                                    PassVerPwd = false;
                                    MessageBox.Show("密碼錯誤，請重新輸入！");
                                    txtPWD.Text = "";
                                }
                            }
                            else
                            {
                                PassVerPwd = false;
                                MessageBox.Show("ID 或 密碼錯誤，請重新輸入！");
                                txtPWD.Text = "";
                            }

                        }
                    }

                }
        }

        private void frmVerPwd_Load(object sender, EventArgs e)
        {
            PassVerPwd = false;
            if  (pEntrySource == "Function8")
            {
                lblID.Visible = true;
                txtID.Visible = true;
            }
            else
            {
                lblID.Visible = false;
                txtID.Visible = false;
            lblRemind.Text = "請輸入使用者ID: " + ClsShareFunc.sUserId + "的密碼";
            }
            //Function 6 & 7 : Administrator only
            if (pEntrySource == "Function6" || pEntrySource == "Function7" || pEntrySource == "Function10")
            {
                if (ClsShareFunc.sLoginIdentity != "Administrator")
                {
                    MessageBox.Show("非【生物、資訊主管-Administrator】權限，無法進入！", "Administrator Only!!!");
                    this.Close();
                }
            }
            //特殊權限:需先以一般行政同仁權限進入
            if (pEntrySource == "Function8")
            {
                if (ClsShareFunc.sLoginIdentity != "Common")
                {
                    MessageBox.Show("需先以一般行政同仁權限進入【再輔以生物、資訊主管-Administrator權限進入】", "行政同仁 First!!!");
                    this.Close();
                }
            }

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            PassVerPwd = false;
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
