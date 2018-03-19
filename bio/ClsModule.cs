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
using System.Runtime.InteropServices;
using System.Reflection;

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

        public static DataGridView nowDGV;
        
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

        /*------- 隱藏身分證中間幾碼  start ------*/

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

        /*------- 隱藏身分證中間幾碼  end ------*/

        /* start 匯出 Excel */
        private static void ReleaseExcelCOM(Excel.Worksheet sheet = null, Excel.Workbook workbook = null, Excel.Application app = null)
        {
            if (sheet != null)
                Marshal.ReleaseComObject(sheet);
            if (workbook != null)
                Marshal.ReleaseComObject(workbook);
            if (app != null)
                Marshal.ReleaseComObject(app);
            sheet = null;
            workbook = null;
            app = null;
            GC.Collect();
        }

        public static void OutPutExcel(DataGridView dgv, string pathName, string fileName)
        {
            //確認datagridview的Name
            string dgvName = dgv.Name.ToString();

            //引用EXCEL Application類別
            Excel.Application myExcel = null;

            //引用活頁簿類別
            Excel.Workbook myBook = null;

            //引用工作表類別
            Excel.Worksheet mySheet = null;

            //引用範圍類別
            Excel.Range myRange = null;

            //開啟一個新的應用程式
            myExcel = new Excel.Application();

            //暫存檔案路徑
            string tmpPath = pathName;

            //FolderBrowserDialog dlg = new FolderBrowserDialog();

            //設定EXCEL檔案路徑
            try
            {

                //if (dlg.ShowDialog() == DialogResult.OK)
                //{
                //    tmpPath = dlg.SelectedPath;
                //}
                // 儲存路徑
                string path = tmpPath;
                // 新增Excel物件
                myExcel = new Microsoft.Office.Interop.Excel.Application();
                // 新增workbook
                myBook = myExcel.Application.Workbooks.Add(true);

                //停用警告訊息
                myExcel.DisplayAlerts = false;

                //讓活頁簿可以看見
                myExcel.Visible = false;

                //引用第一個活頁簿
                myBook = myExcel.Workbooks[1];

                //設定活頁簿為焦點
                myBook.Activate();

                //引用一個工作表
                mySheet = (Excel.Worksheet)myBook.Worksheets[1];

                mySheet.Cells.Clear();

                //設定工作表焦點
                mySheet.Activate();

                switch (dgvName)
                {
                    case "dgvSearchData":
                        //生成Header
                        for (int i = 1; i < dgv.ColumnCount; i++)
                        {
                            mySheet.Cells[1, i] = dgv.Columns[i].HeaderText;
                        }

                        //迴圈加入內容
                        for (int i = 0; i < dgv.RowCount; i++)
                        {
                            for (int j = 1; j < dgv.ColumnCount; j++)
                            {
                                if (dgv[j, i].ValueType == typeof(string))
                                {
                                    mySheet.Cells[i + 2, j] = "'" + dgv[j, i].Value.ToString();
                                }
                                else
                                {
                                    mySheet.Cells[i + 2, j] = dgv[j, i].Value.ToString();
                                }
                            }
                        }

                        //設定EXCEL範圍
                        myRange = mySheet.Range[mySheet.Cells[1, 1], mySheet.Cells[dgv.Rows.Count + 1,

                        dgv.Columns.Count - 1]];

                        break;

                    case "dgvOutRecord":
                        //生成Header
                        for (int i = 0; i < dgv.Columns.GetColumnCount(DataGridViewElementStates.Visible); i++)
                        {
                            mySheet.Cells[1, i+1] = dgv.Columns[i].HeaderText;
                        }

                        //迴圈加入內容
                        for (int i = 0; i < dgv.Rows.GetRowCount(DataGridViewElementStates.Visible); i++)
                        {
                            for (int j = 0; j < dgv.Columns.GetColumnCount(DataGridViewElementStates.Visible); j++)
                            {
                                if (dgv[j, i].ValueType == typeof(string))
                                {
                                    mySheet.Cells[i + 2, j + 1] = "'" + dgv[j, i].Value.ToString();
                                }
                                else
                                {
                                    mySheet.Cells[i + 2, j + 1] = dgv[j, i].Value.ToString();
                                }
                            }
                        }

                        //設定EXCEL範圍
                        myRange = mySheet.Range[mySheet.Cells[1, 1], mySheet.Cells[dgv.Rows.GetRowCount(DataGridViewElementStates.Visible) + 1,

                        dgv.Columns.GetColumnCount(DataGridViewElementStates.Visible)]];
                        break;
                }

                //設定儲存格框線
                myRange.Borders.Weight = Excel.XlBorderWeight.xlThin;

                //column自動對齊
                myRange.EntireColumn.AutoFit();

                //row自動對齊
                myRange.EntireRow.AutoFit();

                if (path.EndsWith("\\"))
                {
                    myBook.SaveAs(path + fileName + ".xlsx");
                }
                else
                {
                    myBook.SaveAs(path + "\\" + fileName + ".xlsx");
                }
                //ReleaseExcelCOM(mySheet, myBook, myExcel);
                MessageBox.Show("匯出成功!");
            }

            catch (Exception ex)
            {
                //throw ex;
                throw new Exception(ex.Message.ToString());
            }
            finally
            {
                //釋放資源
                System.Runtime.InteropServices.Marshal.ReleaseComObject(myRange);
                myRange = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(mySheet);
                mySheet = null;
                myBook.Close(false, Missing.Value, Missing.Value);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(myBook);
                myBook = null;
                myExcel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);
                myExcel = null;
            }
        }
        /* End 匯出 Excel */

        /* 若後面有特殊字元的話就清除 */
        internal static string clearLastMark(string p)
        {
            if (p.Length > 0)
            {
                int starPts = p.IndexOf("*");
                if (starPts >= 0)
                {
                    return StrLeft(p, starPts);
                }
                else
                {
                    return p;
                }
            }
            else
            {
                return p;
            }
        }

        /* 字串處理用 */
        public static string StrLeft(string s, int length)
        {
            return s.Substring(0, length);
        }

        public static string StrRight(string s, int length)
        {
            return s.Substring(s.Length - length);
        }

        public static string StrMid(string s, int start, int length)
        {
            return s.Substring(start, length);
        }
    }
}
