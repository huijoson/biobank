using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Remotion;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;
using System.Data.SqlClient;
using System.Security.Cryptography;

namespace BioBank
{
    class ClsShareFunc
    {
        public static string sUserId;

        public static string sChkID;

        public static string sUserName;

        public static string sLoginIdentity; //主管or一般成員

        public static string sLoginDepartment;//所屬部門(生物or資訊室)

        public static string DbAdmin(){ return "BioAdministratorKeyTbl"; }

        public static string DbCom(){return "BioCommonLoginTbl";}

        
/*============密碼驗證: (1)長度大於8  (2)英數字============*/
        public  static Boolean gfunCheckPwd(string Pwd)
        {
            Boolean Check = false;
            char  [] Character;
            int UpperSum = 0;
            int LowerSum = 0;
            int NumSum = 0;

            if (Pwd.Length >= 8)
            {
                Character = new char[Pwd.Length];

                for (int i = 0; i < Pwd.Length; i++)
                {
                    Character[i] = Convert.ToChar(Pwd.Substring(i, 1));
                    if (Char.IsUpper(Character[i]) == true)
                        UpperSum += 1;
                    if (Char.IsLower(Character[i]) == true)
                        LowerSum += 1;
                    if (Char.IsNumber(Character[i]) == true)
                        NumSum += 1;
                }

                if ((UpperSum > 0 || LowerSum > 0) && NumSum > 0)
                    Check = true;
            }

            return Check;
        }

/*===================check 帳號是否存在db=====================*/
        public static Boolean CheckInDb(string db, string sId, string action)
        {
            Boolean Check = false;

            ////string sEnableFlg = "";
            string sSQL = "";

            //傳進來的第一個參數db,其實是Table name; 都固定去讀DB_SEC
            using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
            {
                sCon.Open();
                sSQL = " select * from " + db + " (nolock) where chUserID = '" + sId + "' ";
                SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                SqlDataReader sRead = sCmd.ExecuteReader();
                if (sRead.HasRows)
                {
                    switch (db)
                    {
                        case "BioCommonLoginTbl":
                            ////while (sRead.Read())
                            ////{
                            ////    sEnableFlg = gfunCheck(sRead["chEnableFlag"]).Trim();
                            ////}

                            if (action == "insert") //&& sEnableFlg == "Y")
                            {
                                MessageBox.Show("此帳號已存在於一般人員帳號!");
                                Check = true;
                                
                            }
                             break;

                        case "BioAdministratorKeyTbl":
                             if (action == "insert")
                             {
                                 MessageBox.Show("此帳號已存在於主管帳號!");
                             }
                             break;
                       };
                    Check = true;
                            //gfunClearTextBox();
                      
                    }sRead.Close();
            }
            return Check;
        }

/*====================check data 不為null=======================*/
        public static string gfunCheck(object sStr)
        {
            string sCheckStr = "";
            if (sStr is DBNull)
                sCheckStr = "";
            else
                sCheckStr = sStr.ToString().Trim();

            return sCheckStr;
        }

        public static string GetMD5(string original)
        {
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            byte[] b = md5.ComputeHash(Encoding.UTF8.GetBytes(original));
            return BitConverter.ToString(b).Replace("-", string.Empty);
        }

        public static string ChangeDateFormat(string sDateTime)
        {
            string sDate = "";
            sDate = (Convert.ToInt32(sDateTime.Substring(0, 4)) - 1911).ToString() + sDateTime.Substring(4, 4).ToString();
            return sDate;
        }

        /*===================inset  Enent Log =====================*/
        public static Boolean insEvenLogt(string eEventNo, string eClerkName, string eLabNo, string eMRNo, string eOtherValue)
        {
           try
            {

                string sSQL = "";

                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                {
                    sCon.Open();
                    sSQL = " insert into BioEventLogTbl (chEventDateTime, chEventNo, chClerkName, chLabNo, chMRNo, chOtherValue) values (dbo.GetDateToDate16(getdate()),'" + eEventNo + "','" + eClerkName + "','" + eLabNo + "','" + eMRNo + "','" + eOtherValue + "')";
                    SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                    sCmd.ExecuteNonQuery();
                }
                 return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Insert Event Log Error: " + ex.Message.ToString());
                return false;
            }
        }

        /*== 隱藏身分證中間幾碼 ==*/

        public static string replaceID(string id, int start, int len)
        {
            string tempStr = "";
            try
            {
                if (!id.Equals(""))
                {
                    tempStr = id.Substring(0,start);
                    for (int i = start; i <= len ; i++)
                    {
                        tempStr += "*";
                    }
                    tempStr += id.Substring((start + len), (id.Length-start-len));
                }
                return tempStr;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
	}
}
