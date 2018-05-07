using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using LinqToExcel;
using Remotion;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Data.OleDb;
using System.Net;//dns
using System.Diagnostics;// 開啟bat
using System.Net.NetworkInformation; // Ping
using BioBank_Conn;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
using System.Collections;

namespace BioBank
{

    public partial class BioBank : Form
    {
        int ComColumn = 30;//共通欄位長度
        Boolean bolClear;
        Boolean bolKeyPass;//是否需再打密碼(true:是/false:否)
        string sExcelName;
        string printNum = "";
        private ContextMenuStrip menu = new ContextMenuStrip();

        string[] StorageRecordColumns = { "檢體管號碼", "舊檢體位置","新檢體位置", "性別", "檢體採集當時年齡", "檢體種類", "檢體採集日期", "檢體採集部位",
                                          "保存方式", "檢體離體時刻","檢體處理時刻","離體後環境", "離體後時間", "收案小組", "罹病部位", "診斷名稱1", "診斷名稱2", "診斷名稱3",
                                          "檔案登錄人", "研究計劃同意書", "同意書編號", "截止日期", "變更範圍", "退出、停止變更、死亡", "變更備註","出庫人","出庫日期","使用者(申請人)","計畫編號","出庫備註","入庫日期"};
        //string[] StorageRecordColumns = { "收案小組", "檢體種類", "保存方式", "罹病部位", "診斷名稱" };
        public static string pFunction8_AdminID = "";

        public BioBank()
        {
            InitializeComponent();
            dgvShowMsg.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            menu.Items.Add("複製");
            menu.ItemClicked += new ToolStripItemClickedEventHandler(contexMenuuu_ItemClicked);
        }

        private void frmBioBank_Load(object sender, EventArgs e)
        {
            sExcelName = "";
            ExcelDt = null;
            bolClear = true;
            bolKeyPass = true;

            /*1.取出title字串,判斷登入者身分*/
            string txt = ""; ClsShareFunc.sUserId = ""; ClsShareFunc.sLoginIdentity = ""; ClsShareFunc.sLoginDepartment = "";
            txt = this.Text.ToString();
            string[] splitTxt = txt.Split(new Char[] { ' ' });
            ClsShareFunc.sLoginIdentity = splitTxt[0].Replace("【", "").Trim();
            ClsShareFunc.sLoginDepartment = splitTxt[1].Replace("(", "").Replace(")", "").Replace("】", "").Trim();
            ClsShareFunc.sUserId = splitTxt[3].Trim();
            ClsShareFunc.sUserName = splitTxt[5].Trim();
            //txtSDate.Text = ClsShareFunc.getTodayDate();
            //txtEDate.Text = ClsShareFunc.getTodayDate();

            /*2.主管維護只有主管才能使用
            switch (ClsShareFunc.sLoginIdentity)
            {
                case "Common":
                    pnlKey.Enabled = false;
                    pnlAdmin.Enabled = false;
                    gbID2LReqNo.Enabled = false;
                    gbLReqNo2All.Enabled = false;
                    break;
                default:
                    pnlKey.Enabled = true;
                    pnlAdmin.Enabled = true;
                    gbID2LReqNo.Enabled = true;
                    gbLReqNo2All.Enabled = true;
                    break;
            };*/

            /*3.資訊室維護只有資訊室人員才能使用
            switch (ClsShareFunc.sLoginDepartment)
            {
                case "B": // B -> 生物資料庫成員 ; M -> 資訊室成員
                    pnlInform.Enabled = false;
                    break;
                default:
                    pnlInform.Enabled = true;
                    break;
            };*/

            LoadCase(cboTeamNo);
            LoadPart(cboAdoptPortion);
        }

        /* 預設檢體部位combobox的值*/
        private void LoadPart(ComboBox comboPart)
        {
            string strSQL = "";
            try
            {
                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                {
                    
                    sCon.Open();
                    strSQL = "SELECT DISTINCT chAdoptPortion from dbo.BioPerMasterTbl";
                    SqlCommand sCmd = new SqlCommand(strSQL, sCon);
                    SqlDataReader sRead = sCmd.ExecuteReader();
                    comboPart.Items.Clear();
                    comboPart.Items.Add("");
                    if (sRead.HasRows)
                    {
                        while (sRead.Read())
                        {
                            comboPart.Items.Add(ClsShareFunc.gfunCheck(sRead["chAdoptPortion"]).ToString().Trim());
                        }
                    }
                    sRead.Close();
                    sRead.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("LoadCase: " + ex.Message.ToString());
            }
        }

        /*cbo載入收案小組*/
        private void LoadCase(ComboBox cboTeamNo)
        {
            string sSQL = "";

            try
            {
                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                {
                    sCon.Open();

                    sSQL = "select sTeam = chCaseNo+'-'+chCaseName  from dbo.BioCaseBasicTbl;";
                    SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                    SqlDataReader sRead = sCmd.ExecuteReader();
                    if (sRead.HasRows)
                    {
                        while (sRead.Read())
                        {
                            cboTeamNo.Items.Add(ClsShareFunc.gfunCheck(sRead["sTeam"]).ToString().Trim());
                        }
                    }
                    sRead.Close();
                    sRead.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("LoadCase: " + ex.Message.ToString());
            }
        }

        public bool IsDate(string strDate)
        {
            //判斷是否為日期
            try
            {
                DateTime.Parse(strDate);
                return true;
            }
            catch
            {
                return false;
            }
        }
        //西元年轉民國(年月日時分秒)
        /*private string ChangeDateTime(string date)
        {
            //西元轉民國
            string[] tempDate = date.Split('/', ' ', ':');
            string mm = "", dd = "", yy = "", hh = "", mo = "", ss = "";
            yy = (Convert.ToInt32(tempDate[0]) - 1911).ToString().PadLeft(3, '0');
            mm = tempDate[1].PadLeft(2, '0');
            dd = tempDate[2].PadLeft(2, '0');
            hh = tempDate[4].PadLeft(2, '0');
            if (tempDate[3] == "下午")
                hh = (Convert.ToInt32(hh) + 12).ToString();
            mo = tempDate[5].PadLeft(2, '0');
            ss = tempDate[6].PadLeft(2, '0');
            return yy + mm + dd + hh + mo + ss;
        }*/
        private string sStr(string sDate)
        {
            string Str; Str = "";
            switch (sDate.Length)
            {
                case 1:
                    Str = "0" + sDate;
                    break;
                case 2:
                    Str = sDate;
                    break;
            };
            return Str;
        }

        private string ChangeDateTime(string date)
        {
            string sYear, sMonth, sDday, sHour, sMiniute, sSceond;
            sYear = ""; sMonth = ""; sDday = ""; sHour = ""; sMiniute = ""; sSceond = "";
            string[] tempDate = date.Split(' ');
            string[] temp1 = tempDate[0].Split('/');
            if (tempDate.Length == 1 && temp1[0] != "")
            {
                sYear = (Convert.ToInt32(temp1[0]) - 1911).ToString();
                if (sYear.Length == 2)
                    sYear = "0" + sYear;
                sMonth = sStr(temp1[1]);
                sDday = sStr(temp1[2]);
            }
            else
            {
                string[] temp2 = tempDate[2].Split(':');
                sYear = (Convert.ToInt32(temp1[0]) - 1911).ToString();
                if (sYear.Length == 2)
                    sYear = "0" + sYear;
                sMonth = sStr(temp1[1]);
                sDday = sStr(temp1[2]);
                sHour = temp2[0];
                sMiniute = temp2[1];
                sSceond = temp2[2];
                if (tempDate[1] == "下午")
                    sHour = (Convert.ToInt32(sHour) + 12).ToString();
            }
            return sYear + sMonth + sDday + sHour + sMiniute + sSceond;
        }

        //回傳Server資料庫民國年月日(7碼:1041231)
        private string GetTime()
        {
            string time;
            //撈資料庫現在時間並將西元轉民國
            using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
            {
                conn.Open();
                SqlCommand gatTime = new SqlCommand("Select sdate=[dbo].GetDate7(GetDate())", conn);
                SqlDataReader drGatTime = gatTime.ExecuteReader();
                drGatTime.Read();
                time = drGatTime["sdate"].ToString();
            }
            return time;
        }
        //回傳Server資料庫民國年月日(7碼:1041231)
        private string GetTime13()
        {
            string time;
            //撈資料庫現在時間並將西元轉民國
            using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
            {
                conn.Open();
                SqlCommand gatTime = new SqlCommand("Select sdate=[dbo].GetDate13(GetDate())", conn);
                SqlDataReader drGatTime = gatTime.ExecuteReader();
                drGatTime.Read();
                time = drGatTime["sdate"].ToString();
            }
            return time;
        }
        //比較日期
        private int DateSubtrac(string date, string Endday)
        {
            //比較時間
            int age = 0;
            string yy = ((Convert.ToInt32(Endday.Substring(0, 3)) + 1911)).ToString();
            string yy2 = ((Convert.ToInt32(date.Substring(0, 3)) + 1911)).ToString();
            string mmdd = Endday.Substring(3, 4);
            string mmdd2 = date.Substring(3, 4);

            if (Convert.ToInt32(mmdd) >= Convert.ToInt32(mmdd2))
                age = Convert.ToInt32(yy) - Convert.ToInt32(yy2);
            else
                age = Convert.ToInt32(yy) - Convert.ToInt32(yy2) - 1;
            return age;
        }
        //加密
        private string Sen_AES(string Data, string LabNo, string inComeDate)
        {
            if (LabNo.Trim() == "" || inComeDate.Trim() == "")
                return "";

            string str_AES = "";
            string key = "";
            //using (SqlConnection conn = new SqlConnection(ClsShareFunc.DB_SECConnection()))
            using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
            {
                conn.Open();
                string strSQL = "select chMasterKey from dbo.BioMasterKeyTbl where chYear='" + inComeDate + "'";
                SqlCommand find_key = new SqlCommand(strSQL, conn);
                SqlDataReader _Key = find_key.ExecuteReader();
                if (_Key.HasRows == true)
                {
                    _Key.Read();
                    key = _Key["chMasterKey"].ToString();
                }

                else
                    return "";
                _Key.Close();
            }
            str_AES = AES.AESencrypt(Data, LabNo + key);
            return str_AES;
        }
        //解密
        private string dec_AES(string Data, string LabNo, string inComeDate)
        {
            string str_AES = "";
            string key = "";
            //using (SqlConnection conn = new SqlConnection(ClsShareFunc.DB_SECConnection()))
            using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
            {
                conn.Open();
                string strSQL = "select chMasterKey from dbo.BioMasterKeyTbl where chYear='" + inComeDate + "'";
                SqlCommand find_key = new SqlCommand(strSQL, conn);
                SqlDataReader _Key = find_key.ExecuteReader();
                if (_Key.HasRows == true)
                {
                    _Key.Read();
                    key = _Key["chMasterKey"].ToString();
                }
                else
                    return "";
            }
            str_AES = AES.AESdecrypt(Data, LabNo + key);
            return str_AES;
        }

        DataTable ExcelDt;

        /*======================瀏覽檔案======================*/
        private void buttonBrowser_Click(object sender, EventArgs e)
        {
            dbPrintMsg.Text = "警告訊息";
            dgvShowMsg.Rows.Clear();
            InitImportPage();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx";
            //openFileDialog1.Filter = "xls files (*.*)|*.*";
            openFileDialog1.Title = "Select a xls File";
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    sExcelName = openFileDialog1.SafeFileName;
                    DataTable dtNew = new DataTable();
                    dtNew = ExcelToDataTable(openFileDialog1.FileName, "共同欄位$");
                    // 4. fill to datagridview
                    dgvShowExcel.Columns.Clear();
                    dgvShowExcel.DataSource = dtNew;
                    textBoxFilePath.Text = openFileDialog1.FileName;
                    buttonPass.Visible = false;

                    //5.檢查format
                    ExcelDt = dtNew;
                      CheckWarning(dtNew);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("buttonBrowser_Click: " + ex.Message);
                }
            }

            /*
            buttonPrint.Visible = false;
           buttonPrintLabNo.Visible = false;
           OpenFileDialog dialog = new OpenFileDialog();
           dialog.Title = "Select file";
           dialog.InitialDirectory = ".\\";
           dialog.Filter = "xls files (*.xls*,*.xlsx)|*.xls*;*.xlsx";
           if (dialog.ShowDialog() == DialogResult.OK)
           {
               //將excel存入datatable
               DataTable dt = ExcelToDataTable(dialog.FileName, "共同欄位");
               ExcelDt = dt;
               textBoxFilePath.Text = dialog.FileName;
               dgvShowExcel.Columns.Clear();
               buttonPass.Visible = false;
               dgvShowExcel.Height = 208;
               dgvShowExcel.DataSource = dt;
               //檢查資料
               CheckWarning(dt);

           }*/

        }

        private DataTable ExcelToDataTable(string filePath, string sheetName)
        {
            string sCon;
            sCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " + filePath + ";Extended Properties= \"Excel 8.0;HDR=yes;IMEX=1\"";
            //sCon = "provider=Microsoft.ACE.OLEDB.12.0;data source=" + filePath + "'Extended Properties='Excel 12.0;HDR=No'";

            string sQry = "SELECT * From [" + sheetName + "]";
            DataTable dt = new DataTable();
            OleDbCommand oleCmd;
            OleDbDataAdapter oda = new OleDbDataAdapter();
            DataTable dtNew = new DataTable();

            using (OleDbConnection oleCon = new OleDbConnection(sCon))
            {
                //3.寫入DataTable
                oleCon.Open();
                DataTable dtExcel;
                dtExcel = oleCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                oleCmd = new OleDbCommand(sQry, oleCon);
                oda = new OleDbDataAdapter(oleCmd);
                oda.Fill(dt);

                //4.新增一個dtNew存放string格式
                for (int j = 0; j < dt.Columns.Count; j++)
                    dtNew.Columns.Add(ClsShareFunc.clearLastMark(dt.Columns[j].ColumnName));
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string[] tmpStr = new string[dt.Columns.Count];
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        string tmpHeader = "";
                        tmpHeader = ClsShareFunc.clearLastMark(dt.Columns[j].ColumnName);

                        switch (tmpHeader)
                        {
                            //初始日期格式
                            case "出生日期":
                            case "檢體採集日期":
                            case "研究計劃簽署日期":
                            case "同意書簽署日期":
                                string[] tmpDate = new string[2];
                                tmpDate = dt.Rows[i][j].ToString().Split(' ');
                                tmpStr[j] = tmpDate[0];
                                break;

                            //初始年齡格式(取整數)
                            case "檢體採集當時年齡":
                            case "年齡":
                                string[] tmpValue = new string[1];
                                tmpValue = dt.Rows[i][j].ToString().Split('.');
                                tmpStr[j] = tmpValue[0];
                                break;

                            default:
                                tmpStr[j] = dt.Rows[i][j].ToString();
                                break;
                        };
                    }
                    dtNew.Rows.Add(tmpStr);
                }
            }
            return dtNew;

            /*DataTable dt = new DataTable();
            ExcelQueryFactory excel = new ExcelQueryFactory(filePath);
            //將excel的row取出來
            IQueryable query = from row in excel.Worksheet(sheetName) select row;

            var columnName = excel.GetColumnNames(sheetName);

            //建立欄位名稱
            foreach (var col in columnName)
            {
                dt.Columns.Add(col.ToString());
            }

            //寫入資料到資料列
            foreach (Row item in query)
            {
                dt.NewRow();
                object[] cell = new object[columnName.Count()];
                int idx = 0;
                foreach (var col in columnName)
                {
                    cell[idx] = item[col].Value;
                    idx++;
                }
                dt.Rows.Add(cell);
            }
            return dt;*/
        }

        string[] CheckColumn = { @"個案碼", "檢體管號碼", "病歷號", "性別", "出生日期", "檢體種類",
                                   "檢體採集日期", "檢體採集部位", "保存方式",  "罹病部位", "診斷名稱1", "檔案登錄人", "同意書編號", "同意書簽署日期" };
        string[] changeDate = { "出生日期", "檢體採集日期", "研究計劃簽署日期", "同意書簽署日期", "檢體離體時刻", "檢體處理時刻" };
        string ErrorString = "";
        //檢查資料
        private bool checkRepeat(DataRow dataRow)
        {
            string LabPieNo = dataRow["檢體管號碼"].ToString();
            string MRNo = dataRow["病歷號"].ToString();
            //using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
            using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
            {
                conn.Open();

                string strSQL = @"select A.chLabNo,B.chInComeDate,A.chMRNo,A.chLabPieNo 
                         from [DB_SEC].[dbo].BioPerMappingTbl as A,[DB_BIO].dbo.BioPerMasterTbl as B 
                         where A.chLabNo COLLATE Chinese_Taiwan_Stroke_BIN=B.chLabNo  COLLATE Chinese_Taiwan_Stroke_BIN";

                SqlDataAdapter find_BioPerMappingTbl = new SqlDataAdapter(strSQL, conn);
                DataTable Data = new DataTable();
                find_BioPerMappingTbl.Fill(Data);
                if (Data.Rows.Count > 0)
                {
                    for (int i = 0; i < Data.Rows.Count; i++)
                    {
                        string _LabPieNo = dec_AES(Data.Rows[i]["chLabPieNo"].ToString().Trim(), Data.Rows[i]["chLabNo"].ToString().Trim(), Data.Rows[i]["chInComeDate"].ToString().Trim().Substring(0, 3));
                        string _MRNo = dec_AES(Data.Rows[i]["chMRNo"].ToString().Trim(), Data.Rows[i]["chLabNo"].ToString().Trim(), Data.Rows[i]["chInComeDate"].ToString().Trim().Substring(0, 3));
                        if (_LabPieNo == "" || _MRNo == "")
                        {
                            MessageBox.Show("解密時出現錯誤, 請通知資訊人員!");
                            return false;
                        }

                        if (_LabPieNo == LabPieNo && _MRNo == MRNo)
                            return false;
                    }
                }
            }
            return true;
        }

        private void CheckWarning(DataTable dt)
        {
            //warning flag
            Boolean warningFlag = false;
            StringBuilder sbWarning = new StringBuilder();

            //判斷column數是否相同
            using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
            {
                conn.Open();
                string strSQL = "select count(*) count from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME ='BioPerSlaver" + cboTeamNo.Text.Substring(0, 1) + "Tbl'";
                SqlCommand Schema = new SqlCommand(strSQL, conn);
                SqlDataReader SchemaCount = Schema.ExecuteReader();
                SchemaCount.Read();
                if (Convert.ToInt32(SchemaCount["count"].ToString()) + ComColumn > dt.Columns.Count)
                {
                    MessageBox.Show("資料庫欄位" + (Convert.ToInt32(SchemaCount["count"].ToString()) + ComColumn) + " 預匯入excel檔欄位" + dt.Columns.Count + "，請重新確認資料!");
                    return;
                }
            }
            dgvShowMsg.Columns.Clear();
            dgvShowMsg.Columns.Add("row", "第幾列");
            dgvShowMsg.Columns.Add("msg", "錯誤訊息");
            string Warning = "";
            string bodyExWarning = "";
            string SlaverTbl = "BioPerSlaver" + cboTeamNo.Text.Substring(0, 1) + "Tbl";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (checkRepeat(dt.Rows[i]) == false)
                {
                    warningFlag = true;

                    bodyExWarning = "檢體管號碼重複:" + dt.Rows[i][1];
                    dgvShowMsg.Rows.Add("第" + Convert.ToInt32(i+1) + "列", bodyExWarning);
                }
                //病歷號是否為身分證 可pass
                if (dt.Rows[i]["病歷號"].ToString().Length < 10)
                    Warning += "病歷號不是身分證" + Environment.NewLine;
                if (dt.Rows[i]["收案小組"].ToString() != cboTeamNo.Text.Substring(2, cboTeamNo.Text.Length - 2))
                    Warning += "所選收案小組與收案小組資料不同!\n";
                if (Warning.Length > 0)
                    dgvShowMsg.Rows.Add("第" + (i + 2) + "列", Warning);
                Warning = "";
            }
            if (dgvShowMsg.Rows.Count > 0)
            {
                dbPrintMsg.Text = "警告訊息";
                dgvShowMsg.Visible = true;
                buttonPass.Visible = true;
                buttonPrint.Visible = true;
            }
            else
                CheckFile(dt);
            if (warningFlag == true)
            {
                MessageBox.Show("檢體管號重複，請重新確認資料");
            }
        }

        private void InitImportPage()
        {
            gbDataImport.Height = 237;
            dgvShowExcel.Height = 198;
            dbPrintMsg.Visible = true;

            dgvShowExcel.DataSource = null;
        }

        private void CheckFile(DataTable dt)
        {
            dgvShowMsg.Columns.Clear();
            dgvShowMsg.Columns.Add("row", "第幾列");
            dgvShowMsg.Columns.Add("msg", "錯誤訊息");
            buttonPass.Visible = false;
            for (int i = 0; i < dt.Rows.Count; i++)
            {

                //相同檢體管號碼
                for (int j = i + 1; j < dt.Rows.Count - 1; j++)
                {
                    if (dt.Rows[i]["檢體管號碼"].ToString() == dt.Rows[j]["檢體管號碼"].ToString())
                        ErrorString += "檢體館號碼與第" + (j + 1) + "個檢體管號碼相同!!\n";
                }
                //如有研究同意書，簽屬日期則不可空白
                if (dt.Rows[i]["研究計劃同意書"].ToString().Trim() != "")
                {
                    if (dt.Rows[i]["研究計劃簽署日期"].ToString().Trim() == "")
                        ErrorString += "研究計劃簽署日期不可空白!!\n";
                }
                //檢查不可空白
                for (int j = 0; j < CheckColumn.Length; j++)
                {
                    if (dt.Rows[i][CheckColumn[j]].ToString().Trim() == "")
                        ErrorString += CheckColumn[j] + " 不可空白!!\n";
                }
                //檢查時間格式
                for (int k = 0; k < changeDate.Length; k++)
                {
                    if (dt.Rows[i][changeDate[k]].ToString().Trim() != "")
                    {
                        //判斷日期格式
                        if (IsDate(dt.Rows[i][changeDate[k]].ToString().Trim()) == false)
                            ErrorString += changeDate[k] + "日期格試錯誤!\n";
                        else
                        {
                            // if (changeDate[k].ToString() == "檢體離體時刻" || changeDate[k].ToString() == "檢體處理時刻")
                            dt.Rows[i][changeDate[k]] = ChangeDateTime(dt.Rows[i][changeDate[k]].ToString());
                            //  else
                            //    dt.Rows[i][changeDate[k]] = ChangeDateTime(dt.Rows[i][changeDate[k]].ToString()).Substring(0, 7);
                        }
                    }
                }

                //計算檢體採集當時年齡
                //dt.Rows[i]["檢體採集當時年齡"] = DateSubtrac(dt.Rows[i]["出生日期"].ToString(), dt.Rows[i]["檢體採集日期"].ToString()).ToString().Substring(0,2);

                if (ErrorString.Length > 0)
                    dgvShowMsg.Rows.Add("第" + (i + 1) + "列", ErrorString);
                ErrorString = "";
            }
            if (dgvShowMsg.Rows.Count > 0)
            {
                dbPrintMsg.Text = "錯誤訊息";
                buttonSaveToDB.Visible = false;
                buttonPrintLabNo.Visible = false;
                dbPrintMsg.Visible = true;
                buttonClear.Visible = false;
                buttonPrint.Visible = true;
                cboTeamNo.Enabled = true;
                MessageBox.Show("資料有誤，無法倒入資料庫");
                //將光碟退出
                EjectMedia.EjectMedia.Eject(@"\\.\E:");
            }
            else
            {
                cboTeamNo.Enabled = false;
                dgvShowExcel.DataSource = null;
                dbPrintMsg.Visible = false;
                gbDataImport.Height = 445;
                dgvShowExcel.Height = 400;

                dgvShowExcel.Columns.Add("筆數", "筆數");
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (i == 3)
                    {
                        dgvShowExcel.Columns.Add("新檢體位置", "新檢體位置");
                        dgvShowExcel.Columns[i].ReadOnly = false;
                    }
                    dgvShowExcel.Columns.Add(dt.Columns[i].ColumnName, dt.Columns[i].ColumnName);
                    if (i != 4)
                        dgvShowExcel.Columns[i].ReadOnly = true;
                }
                dgvShowExcel.RowCount = dt.Rows.Count;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    
                    dgvShowExcel.Rows[i].HeaderCell.Value = (i + 2).ToString();
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        if (j == 0)
                        {
                            dgvShowExcel.Rows[i].Cells[0].Value = i + 1;
                        }
                        if (j < 3)
                            dgvShowExcel.Rows[i].Cells[j + 1].Value = dt.Rows[i][j].ToString().Trim();
                        else
                            dgvShowExcel.Rows[i].Cells[j + 2].Value = dt.Rows[i][j].ToString().Trim();
                    }
                }
                buttonPrint.Visible = false;
                buttonSaveToDB.Visible = true;
                buttonClear.Visible = true;
            }
        }
        //匯入按鈕
        private void buttonSaveToDB_Click(object sender, EventArgs e)
        {
            //insert Event Log: 5. --匯入Excel (start)--
            ClsShareFunc.insEvenLogt("5", ClsShareFunc.sUserName, "", "", "匯入Excel (start)--" + sExcelName);
            ArrayList posList = new ArrayList();
            Dictionary<string, string> dicPos = new Dictionary<string, string>();
            Dictionary<string, string> dicSpacePos = new Dictionary<string, string>();
            Boolean checkFlag = false;
            InitImportPage();
            dbPrintMsg.Text = "列印訊息";

            //檢查使用者輸入有沒有重複的檢體
            dgvShowMsg.Rows.Clear();

            for (int i = 0; i < dgvShowExcel.Rows.Count; i++)
            {

                if (dgvShowExcel.Rows[i].Cells["新檢體位置"].Value == null)
                {
                    dicSpacePos.Add((i + 1).ToString(), "");
                    checkFlag = true;
                }
                else
                {
                    dicPos.Add((i + 1).ToString(), dgvShowExcel.Rows[i].Cells["新檢體位置"].Value.ToString());
                }
            }
            //檢查重複值
            if (checkFlag != true)
            {
                var duplicateValues = dicPos.GroupBy(x => x.Value).Where(x => x.Count() > 1);
                if (duplicateValues.Count() > 0)
                {
                    try
                    {
                        foreach (KeyValuePair<string, string> item in dicPos)
                        {
                            foreach (var item2 in duplicateValues)
                            {
                                if (item.Value == item2.Key)
                                {
                                    dgvShowMsg.Rows.Add("第" + item.Key + "列", "檢體位置重複: " +
                                    dgvShowExcel.Rows[Convert.ToInt32(item.Key) - 1].Cells["檢體管號碼"].Value.ToString() + " : " +
                                    dgvShowExcel.Rows[Convert.ToInt32(item.Key) - 1].Cells["新檢體位置"].Value.ToString());
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("檢查目前輸入的檢體位置有無重複的錯誤: " + ex.Message);
                    }
                }
                posList = checkPos(dicPos);
                if (posList.Count > 0)
                {
                    try
                    {
                        for (int i = 0; i <= posList.Count; i++)
                        {
                            dgvShowMsg.Rows.Add("第" + posList[i] + "列", "檢體位置重複: " +
                                        dgvShowExcel.Rows[Convert.ToInt32(posList[i]) - 1].Cells["檢體管號碼"].Value.ToString() + " : " +
                                        dgvShowExcel.Rows[Convert.ToInt32(posList[i]) - 1].Cells["新檢體位置"].Value.ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("檢查資料庫與目前輸入的檢體位置的錯誤: " + ex.Message);
                    }
                }
                if (dgvShowMsg.Rows.Count > 0)
                {
                    MessageBox.Show("共 " + dgvShowMsg.Rows.Count + " 筆檢體位置重複，檢體位置不可重複!");
                    return;
                }
            }
            else
            {
                foreach (KeyValuePair<string, string> item in dicSpacePos)
                {
                    dgvShowMsg.Rows.Add("第" + item.Key + "列", "檢體位置空白: " +
                                dgvShowExcel.Rows[Convert.ToInt32(item.Key) - 1].Cells["檢體管號碼"].Value.ToString() + " : " +
                                "");
                }
                MessageBox.Show("共 " + dgvShowMsg.Rows.Count + " 筆檢體位置空白，檢體位置不可空白!");
                return;
            }

            DialogResult myResult = MessageBox.Show("共有" + ExcelDt.Rows.Count + "筆資料，確定要倒入資料庫?", "資料格式正確", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (myResult == DialogResult.Yes)
            {
                //insert Event Log: 5-1. --匯入Excel  (Press Yes)-- 
                ClsShareFunc.insEvenLogt("5-1", ClsShareFunc.sUserName, "", "", "匯入Excel  (Press Yes)--" + "收案組別:" + cboTeamNo.Text.Substring(0, 1) + sExcelName);

                //CopyExcel(dt);
                DataTable tempDt = ExcelToDataTable(textBoxFilePath.Text, "共同欄位$");
                //var excelFile = new ExcelQueryFactory(textBoxFilePath.Text);
                //確認資料筆數是否相同
                if (tempDt.Rows.Count == ExcelDt.Rows.Count)
                {
                    string FileName = @"C:\\BioBankLog\" + sExcelName + "_" + GetTime13() + ".xls";
                    if (File.Exists(FileName))
                        File.Delete(FileName);
                    File.Copy(textBoxFilePath.Text, FileName);
                    cboTeamNo.Enabled = false;
                    if (SaveToDB() == false)
                    {
                        //insert Event Log: 5-2. --匯入Excel (加密出現錯誤)-- 
                        ClsShareFunc.insEvenLogt("5-2", ClsShareFunc.sUserName, "", "", "匯入Excel (加密出現錯誤)--" + "收案組別:" + cboTeamNo.Text.Substring(0, 1) + sExcelName);
                        MessageBox.Show("加密時出現錯誤,立即停上作業,通知資訊人員處理!");
                    }
                    else
                    {
                        //insert Event Log: 5-3. --匯入Excel (成功)-- 
                        ClsShareFunc.insEvenLogt("5-3", ClsShareFunc.sUserName, "", "", "匯入Excel (成功)--" + "收案組別:" + cboTeamNo.Text.Substring(0, 1) + sExcelName);
                    }
                }
                else
                {
                    //insert Event Log: 5-21. --匯入Excel (錯誤:欲存入資料不符)-- 
                    ClsShareFunc.insEvenLogt("5-21", ClsShareFunc.sUserName, "", "", "匯入Excel (錯誤:欲存入資料不符)--" + "收案組別:" + cboTeamNo.Text.Substring(0, 1) + sExcelName);
                    MessageBox.Show("存入資料與確認資料不符，請重新操作一次!");
                    textBoxFilePath.Text = "";
                    cboTeamNo.Enabled = true;
                }
                LoadPart(cboAdoptPortion);
            }
            else if (myResult == DialogResult.No)
            {
                cboTeamNo.Enabled = true;
                textBoxFilePath.Text = "";
            }
        }
        //-------------------------確認檢體位置有沒有重複
        private ArrayList checkPos(Dictionary<string, string> dicPos)
        {
            string strSQL = "select distinct chNewLabPositon from [DB_BIO].[dbo].[BioPerMasterTbl]";
            ArrayList alRepeatList = new ArrayList();
            using (SqlConnection conn1 = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
            {
                conn1.Open();
                SqlCommand sCmd = new SqlCommand(strSQL, conn1);
                SqlDataReader sRead = sCmd.ExecuteReader();

                if (sRead.HasRows)
                {
                    while (sRead.Read())
                    {
                        foreach (KeyValuePair<string, string> item in dicPos)
                        {
                            if (item.Value == ClsShareFunc.gfunCheck(sRead["chNewLabPositon"].ToString()))
                            {
                                alRepeatList.Add(item.Key);
                            }
                        }
                    }
                }

                return alRepeatList;
            }
        }
        //--------------------
        string[] secTabl = { "個案碼", "檢體管號碼", "病歷號", "出生日期", "研究計劃同意書",  "同意書編號", "姓名",   
                           "chPerCaseNo", "chLabPieNo", "chMRNo", "chBirthday", "chPlanAgree", "chAgreeNo", "chMRName"};
        string[] comTabl = { "檢體位置","新檢體位置", "性別", "檢體採集當時年齡", "檢體採集日期", "檢體採集部位", "保存方式",  "檢體離體時刻","檢體處理時刻", "離體後環境", "離體後時間", "罹病部位", "診斷名稱1", "診斷名稱2", "診斷名稱3", "檔案登錄人","研究計劃簽署日期","同意書簽署日期", "截止日期", "變更範圍", "退出、停止變更、死亡", "備註",
                           "chOldLabPosition","chNewLabPositon", "chSex", "intAge", "chLabAdoptDate", "chAdoptPortion", "chStoreageMethod", "chLabLeaveBodyDatetime","chLabDealDatetime" , "chLabLeaveBodyEnvir", "chLabLeaveBodyHour", "chSickPortion", "chDiagName1", "chDiagName2", "chDiagName3", "chClerkName", "chPlanAgreeDate", "chAgreeNoDate", "chUseLimitYear", "chChangeRange", "chStatus", "chNote"};
        //存入資料庫
        private bool SaveToDB()
        {
            dgvShowMsg.Columns.Clear();
            dgvShowMsg.Columns.Add("新檢體管號碼", "新檢體管號碼");
            dgvShowMsg.Columns.Add("舊檢體管號碼", "舊檢體管號碼");
            dgvShowMsg.Columns.Add("新檢體位置", "新檢體位置");
            dgvShowMsg.Columns.Add("舊檢體位置", "舊檢體位置");
            string NewNumber = "";
            string SlaverTbl = "BioPerSlaver" + cboTeamNo.Text.Substring(0, 1) + "Tbl";
            DataTable SecDt = new DataTable();
            DataTable BiologyDt = new DataTable();
            DataTable SlaverDt = new DataTable();
            //using (SqlConnection conn = new SqlConnection(ClsShareFunc.DB_SECConnection()))
            using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
            {
                conn.Open();
                //將sec資料庫的格式取出
                SqlDataAdapter SecDtAdapter = new SqlDataAdapter("select * from BioPerMappingTbl where 2=1", conn);
                SecDtAdapter.Fill(SecDt);

                using (SqlConnection conn1 = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                {
                    conn1.Open();
                    //將bio資料庫的格式取出
                    SqlDataAdapter BiologyDtAdapter = new SqlDataAdapter("select * from BioPerMasterTbl where 2=1", conn1);
                    BiologyDtAdapter.Fill(BiologyDt);
                    SqlDataAdapter SlaverDtAdapter = new SqlDataAdapter("select * from " + SlaverTbl + " where 2=1", conn1);
                    SlaverDtAdapter.Fill(SlaverDt);
                    string NowDate = ChangeDateTime(DateTime.Now.ToString()).Substring(0, 11);
                    for (int i = 0; i < dgvShowExcel.Rows.Count; i++)
                    {
                        SecDt.Rows.Add();
                        BiologyDt.Rows.Add();
                        SlaverDt.Rows.Add();

                        //新檢體號碼
                        NewNumber = RandomNumber(dgvShowExcel.Rows[i].Cells["檢體種類"].Value.ToString(), i);
                        dgvShowMsg.Rows.Add(NewNumber, dgvShowExcel.Rows[i].Cells["檢體管號碼"].Value, dgvShowExcel.Rows[i].Cells["新檢體位置"].Value, dgvShowExcel.Rows[i].Cells["檢體位置"].Value);
                        SecDt.Rows[i]["chLabNo"] = NewNumber;
                        BiologyDt.Rows[i]["chLabNo"] = NewNumber;
                        BiologyDt.Rows[i]["chInComeDate"] = NowDate;

                        //是否是sec
                        Boolean SecFlag = false;

                        //共通欄位
                        for (int j = 0; j < ComColumn; j++)
                        {
                            SecFlag = false;
                            for (int k = 0; k < (secTabl.Length) / 2; k++)
                            {
                                //需存入SEC的資料
                                if (dgvShowExcel.Columns[j].Name == secTabl[k])
                                {
                                    //SecDt.Rows[i][secTabl[k + ((secTabl.Length) / 2)]] = Sen_AES(dgvShowExcel.Rows[i].Cells[j].Value.ToString(), NewNumber, GetTime().Substring(0, 3));
                                    SecDt.Rows[i][secTabl[k + ((secTabl.Length) / 2)]] = Sen_AES(dgvShowExcel.Rows[i].Cells[j].Value.ToString(), NewNumber, NowDate.Substring(0, 3));
                                    //Sen_AES :AES加密時傳入, 欲加密9欄位值, 新檢體號碼, 所屬的年度
                                    if (SecDt.Rows[i][secTabl[k + ((secTabl.Length) / 2)]] == "")
                                    {
                                        //加密時出現錯誤, Return False; 由上一層的button_click show出message:「立即停上作業,通知資訊人員」
                                        return false;
                                    }
                                    SecFlag = true;
                                    break;
                                }
                            }
                            //共通欄位
                            if (dgvShowExcel.Columns[j].Name == "檢體種類")
                                BiologyDt.Rows[i]["chLabType"] = dgvShowExcel.Rows[i].Cells[j].Value.ToString().Substring(0, 1);
                            else if (dgvShowExcel.Columns[j].Name == "收案小組")
                            {
                                string[] temp = cboTeamNo.Text.Trim().Split('-');
                                BiologyDt.Rows[i]["chSubStock"] = temp[1];
                            }
                            else
                            {
                                for (int x = 0; x < (comTabl.Length) / 2; x++)
                                {
                                    if (SecFlag == true)
                                        break;
                                    if (dgvShowExcel.Columns[j].Name == comTabl[x])
                                    {
                                        if (dgvShowExcel.Rows[i].Cells[j].ToString().Trim() != "")
                                            /* if( j == 7)
                                             {
                                                 BiologyDt.Rows[i][comTabl[x + ((comTabl.Length) / 2)]] = Convert.ToInt32( dgvShowExcel.Rows[i].Cells[j].Value.ToString().Trim().Substring(0,2));
                                             }
                                         else*/
                                            BiologyDt.Rows[i][comTabl[x + ((comTabl.Length) / 2)]] = dgvShowExcel.Rows[i].Cells[j].Value;
                                        break;
                                    }
                                }
                            }

                        }
                        //小組資料庫
                        for (int j = 0; j < SlaverDt.Columns.Count; j++)
                        {
                            //+2是因為編號跟新檢體位置
                            SlaverDt.Rows[i][j] = dgvShowExcel.Rows[i].Cells[j + ComColumn + 2].Value;
                        }
                    }
                        //更新資料庫
                        SqlCommandBuilder builderSec = new SqlCommandBuilder(SecDtAdapter);
                        SecDtAdapter.Update(SecDt);
                        SqlCommandBuilder builderBio = new SqlCommandBuilder(BiologyDtAdapter);
                        BiologyDtAdapter.Update(BiologyDt);
                        SqlCommandBuilder builderSlaver = new SqlCommandBuilder(SlaverDtAdapter);
                        SlaverDtAdapter.Update(SlaverDt);
                        MessageBox.Show("匯入成功!!");
                        buttonSaveToDB.Visible = false;
                        // dgvShowExcel.Height = 208;
                        dgvShowMsg.Visible = true;
                        buttonPrintLabNo.Visible = true;
                        cboTeamNo.Enabled = true;
                        buttonPrint.Visible = true;
                }
            }
            return true;
        }
        //編新檢體號碼
        private string RandomNumber(string type, int Row)
        {
            string newNumber = "";
            Random randomtemp = new Random(Guid.NewGuid().GetHashCode());
            for (int i = 0; i < Row; i++)
            {
                if (dgvShowExcel.Rows[Row].Cells["個案碼"].Value.ToString().Trim() == dgvShowExcel.Rows[i].Cells["個案碼"].Value.ToString().Trim()
                    && dgvShowExcel.Rows[Row].Cells["檢體採集日期"].Value.ToString().Trim() == dgvShowExcel.Rows[i].Cells["檢體採集日期"].Value.ToString().Trim())
                {
                    while (true)
                    {
                        Boolean frag = true;
                        newNumber = dgvShowMsg.Rows[i].Cells[0].Value.ToString().Substring(0, 7) + type.Substring(0, 1) + randomtemp.Next(0, 10);
                        for (int j = 0; j < Row; j++)
                        {
                            if (dgvShowMsg.Rows[i].Cells[0].Value.ToString() == newNumber)
                                frag = false;
                        }
                        if (frag == true)
                            return newNumber;
                    }
                }
            }
            //第一碼為小組代碼
            newNumber = cboTeamNo.Text.Substring(0, 1).ToLower();
            Boolean SameFlag = false;
            while (true)
            {
                SameFlag = false;
                //二~六碼為隨機亂碼
                newNumber += randomtemp.Next(100000, 1000000);
                using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                {
                    conn.Open();
                    //找是否已有相同亂數編號
                    SqlDataAdapter SearchDateAdapter = new SqlDataAdapter("Select  chLabNo From BioPerMasterTbl where SUBSTRING(chLabNo,0,6)='" + newNumber + "'", conn);
                    DataTable dtDate = new DataTable();
                    SearchDateAdapter.Fill(dtDate);
                    if (dtDate.Rows.Count == 0)
                    {
                        for (int i = 0; i < Row; i++)
                        {
                            if (newNumber == dgvShowMsg.Rows[i].Cells[0].Value.ToString().Substring(0, 7))
                            {
                                SameFlag = true;
                                break;
                            }
                        }
                        if (SameFlag == false)
                            break;
                    }
                }
            }
            //第七碼為檢體種類+第八碼隨機碼
            newNumber += type.Substring(0, 1) + randomtemp.Next(1, 10);
            return newNumber;
        }

        //列印
        private void buttonPrint_Click(object sender, EventArgs e)
        {
            string headerText = "";
            string[] tempPath = textBoxFilePath.Text.Split('.', '\\');
            headerText = tempPath[tempPath.Length - 2] + "\n 時間 :" + DateTime.Now + "\n";
            if (dgvShowMsg.Columns.Count == 4)
                headerText += "檢體編號";
            else
                headerText += "錯誤清單";
            //列印gridview資料
            ClsPrint _ClsPrint = new ClsPrint(dgvShowMsg, headerText);
            _ClsPrint.PrintForm();

        }
        //列印新檢體編碼(貼紙)
        private void buttonPrintLabNo_Click(object sender, EventArgs e)
        {
            var PD = new PrintDocument();

            if (printFunction.IsPrinterExist("CAB MACH4/300"))
            {
                PD.PrinterSettings.PrinterName = "CAB MACH4/300";
                try
                {
                    for (int i = 0; i < dgvShowMsg.Rows.Count; i++)
                    //for (int i = 0; i < 1; i++)
                    {
                        printNum = dgvShowMsg.Rows[i].Cells[0].Value.ToString();
                        PD.PrintPage += new PrintPageEventHandler(PD_PrintPage);
                        PD.Print();
                    }
                    //printNum = "U121516777";
                    //PD.PrintPage += new PrintPageEventHandler(PD_PrintPage);
                    //PD.Print();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void comboBoxCase_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (cboTeamNo.Text != "")
            {
                panel1.Visible = true;
            }
        }
        //登出按鈕
        private void buttonSignOut_Click(object sender, EventArgs e)
        {
            //登出
            LogIn login = new LogIn();
            login.Show();
            this.Close();
        }
        private void dgvShowExcel_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            for (int i = 0; i < dgvShowExcel.Rows.Count; i++)
                dgvShowExcel.Rows[i].HeaderCell.Value = (i + 2).ToString();
        }
        //重新選擇小組按鈕(清空)
        private void buttonClear_Click(object sender, EventArgs e)
        {
            InitImportPage();
            cboTeamNo.Enabled = true;
            buttonPrint.Visible = false;
            buttonPrintLabNo.Visible = false;
            dgvShowExcel.Rows.Clear();
            dgvShowMsg.Rows.Clear();
            textBoxFilePath.Text = "";
        }
        //Next按鈕(略過警告通知)
        private void buttonPass_Click(object sender, EventArgs e)
        {
            CheckFile(ExcelDt);
        }

        //查看-顯示入庫時間
        private void StorageTimeRecord()
        {
            using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
            {
                conn.Open();
                SqlDataAdapter SearchDateAdapter = new SqlDataAdapter(@"select chInComeDate as 入庫日期, left(chLabNo,1) as 收案小組代碼,count(chInComeDate) as 筆數 
                from BioPerMasterTbl (nolock) group by chInComeDate, left(chLabNo,1) order by chInComeDate desc, left(chLabNo,1)", conn);
                DataTable dtDate = new DataTable();
                SearchDateAdapter.Fill(dtDate);
                dgvStorageTime.DataSource = dtDate;
                dgvStorageTime.ClearSelection();
            }
            int count = 0;
            for (int i = 0; i < dgvStorageTime.Rows.Count; i++)
                count += Convert.ToInt32(dgvStorageTime.Rows[i].Cells[2].Value);
            StorageCount.Text = "共" + count + "筆";
        }
        //查看-點選入庫時間
        private void dgvStorageTime_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string comeDate = dgvStorageTime.Rows[e.RowIndex].Cells[0].Value.ToString();
            string chCase = dgvStorageTime.Rows[e.RowIndex].Cells[1].Value.ToString();
            string sql = " select * from BioPerMasterTbl where chInComeDate='" + comeDate + "' and left(chLabNo,1)='" + chCase + "'";
            using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
            {
                conn.Open();
                //insert Event Log: 8. --流灠匯入Excel記錄--
                ClsShareFunc.insEvenLogt("8", ClsShareFunc.sUserName, "", "", "流灠匯入Excel記錄--" + comeDate + chCase + sExcelName);
                SqlDataAdapter SearchDateAdapter = new SqlDataAdapter(sql, conn);
                DataTable dtDate = new DataTable();
                SearchDateAdapter.Fill(dtDate);
                for (int i = 0; i < StorageRecordColumns.Length; i++)
                    dtDate.Columns[i].ColumnName = StorageRecordColumns[i];
                dgvStorageRecord.DataSource = dtDate;
            }

        }
        //查看-列印
        private void buttonRecordPrint_Click(object sender, EventArgs e)
        {
            DataGridView RecordTemp = new DataGridView();
            RecordTemp.Columns.Add("新檢體管號碼", "新檢體管號碼");
            RecordTemp.Columns.Add("舊檢體管號碼", "舊檢體管號碼");
            RecordTemp.Columns.Add("新檢體位置", "新檢體位置");
            RecordTemp.Columns.Add("舊檢體位置", "舊檢體位置");
            //using (SqlConnection conn = new SqlConnection(ClsShareFunc.DB_SECConnection()))
            using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
            {
                conn.Open(); for (int i = 0; i < dgvStorageRecord.Rows.Count; i++)
                {
                    //找出舊檢體編號
                    SqlDataAdapter SecDtAdapter = new SqlDataAdapter("select chLabPieNo from BioPerMappingTbl where chLabNo='" + dgvStorageRecord.Rows[i].Cells[0].Value + "'", conn);
                    DataTable SecDt = new DataTable();
                    SecDtAdapter.Fill(SecDt);
                    string _MRNo = dec_AES(SecDt.Rows[0][0].ToString().Trim(), dgvStorageRecord.Rows[i].Cells["檢體管號碼"].Value.ToString().Trim(), dgvStorageRecord.Rows[i].Cells["入庫日期"].Value.ToString().Trim().Substring(0, 3));
                    RecordTemp.Rows.Add(dgvStorageRecord.Rows[i].Cells["檢體管號碼"].Value, _MRNo, dgvStorageRecord.Rows[i].Cells["新檢體位置"].Value, dgvStorageRecord.Rows[i].Cells["舊檢體位置"].Value);
                }
            }

            ClsPrint _ClsPrint = new ClsPrint(RecordTemp, "對照表");
            _ClsPrint.PrintForm();
        }
        //查詢-搜尋按鈕
        private void buttonSearch_Click(object sender, EventArgs e)
        {
            dgvSearchData.Rows.Clear();
            string sql = "select * from BioPerMasterTbl";
            //string sql = "select chSubStock, chLabType, chStoreageMethod, chSickPortion, chDiagName1, chDiagName2, chDiagName3 from BioPerMasterTbl";
            //小組欄位
            //if (comboBoxCase2.Text.Trim() != "" && comboBoxColumn.Text.Trim() != "" && textBoxCaseSearch.Text.Trim() != "")
            //{
            //    sql+=", 'BioPerSlaver "+comboBoxCase2.Text.Substring(0,1)+"Tbl' where "
            //}
            //else

            sql += " where ";

            //性別
            if (checkBoxSexM.Checked == true && checkBoxSexF.Checked == false)
                sql += " chSex='男' and ";
            else if (checkBoxSexM.Checked == false && checkBoxSexF.Checked == true)
                sql += " chSex='女' and ";
            //年齡
            if (textBoxAge1.Text.Trim() != "" && textBoxAge2.Text.Trim() != "")
            {
                if (textBoxAge1.Text.CompareTo(textBoxAge2.Text) < 0)
                    sql += " intAge>=" + textBoxAge1.Text + " and intAge<= " + textBoxAge2.Text + " and ";
                else
                    sql += " intAge>=" + textBoxAge2.Text + " and intAge<= " + textBoxAge1.Text + " and ";
            }
            else if (textBoxAge1.Text.Trim() != "" && textBoxAge2.Text.Trim() == "")
                sql += " intAge=" + textBoxAge1.Text + " and ";
            else if (textBoxAge1.Text.Trim() == "" && textBoxAge2.Text.Trim() != "")
                sql += " intAge=" + textBoxAge2.Text + " and ";
            //檢體部位
            if (cboAdoptPortion.Text.Trim() != "")
            {
                sql += "chAdoptPortion like'%" + cboAdoptPortion.Text + "%' and ";
            }
            //檢體種類
            int typeCount = 0;
            foreach (Control c in groupBoxLabType.Controls)
            {
                if (c is CheckBox)
                {
                    CheckBox chk = (CheckBox)c;
                    if (chk.Checked)
                    {
                        sql += " chLabType='" + chk.Text.Substring(0, 1) + "' or ";
                        typeCount++;
                    }
                }
            }
            if (sql.Substring(sql.Length - 4, 4).Trim() == "or")
            {
                sql = sql.Substring(0, sql.Length - 4) + " and ";
            }
            
            //診斷
            if (textBoxDiag1.Text != "" || textBoxDiag2.Text != "" || textBoxDiag3.Text != "")
            {
                sql += " (";
                for (int i = 1; i < 4; i++)
                {
                    TextBox text = (TextBox)this.Controls.Find("textBoxDiag" + i, true).FirstOrDefault();
                    if (text.Text.Trim() != "")
                    {
                        if (i == 2)
                        {
                            if (textBoxDiag1.Text.Trim() != "")
                                sql += comboBoxRel1.Text + " ";
                        }
                        else if (i == 3)
                        {
                            if (textBoxDiag2.Text.Trim() != "")
                                sql += comboBoxRel2.Text + " ";
                        }
                        sql += "(chDiagName1+' '+chDiagName2+' '+chDiagName3 like'%" + text.Text + "%') ";

                    }
                }
                sql += " ) and ";
            }
            
            //出庫
            if (chkGetOut.Text != "")
            {
                string takeStr = chkGetOut.Text.ToString();
                switch (takeStr)
                {
                    case "未出庫":
                        sql += " chTakeOutDate is null ";
                        break;
                    case "出庫":
                        sql += " chTakeOutDate is not null ";
                        break;
                }
                sql += "and ";
            }

            //收案小組
            if (cbGroup.Text != "")
            {
                string gpStr = cbGroup.Text.ToString();
                switch (gpStr)
                {
                    case "腎臟分庫":
                        sql += " chSubStock = '腎臟分庫' ";
                        break;
                    case "婦科腫瘤分庫":
                        sql += " chSubStock = '婦科腫瘤分庫' ";
                        break;
                    case "肝病收案小組":
                        sql += " chSubStock = '肝病收案小組' ";
                        break;
                    case "消化系器官組織檢體資料庫":
                        sql += " chSubStock = '消化系器官組織檢體資料庫' ";
                        break;
                    case "惡性腦瘤組織庫":
                        sql += " chSubStock = '惡性腦瘤組織庫' ";
                        break;                                                                                       
                }
                sql += "and ";
            }

            //檢體保存期限
            if (txtSDate.Text.ToString().Trim() != "" || txtEDate.Text.ToString().Trim() != "")
            {
                if (txtSDate.Text.ToString().Trim() != "" && txtEDate.Text.ToString().Trim() != "")
                {
                    if (txtSDate.Text.ToString().Trim() == txtEDate.Text.ToString().Trim())
                    {
                        sql += "chUseExpireDate = '" + txtSDate.Text.ToString().Trim() + "' and ";
                    }
                    else
                    {
                        sql += "chUseExpireDate >= '" + txtSDate.Text.ToString().Trim() + "' and ";
                        sql += "chUseExpireDate <= '" + txtEDate.Text.ToString().Trim() + "' and ";
                    }
                }
                else
                {
                    if (txtSDate.Text.ToString().Trim() == "")
                    {
                        txtSDate.Text = txtEDate.Text.ToString().Trim();
                    }
                    if (txtEDate.Text.ToString().Trim() == "")
                    {
                        txtEDate.Text = txtSDate.Text.ToString().Trim();
                    }
                    sql += "chUseExpireDate = '" + txtSDate.Text.ToString().Trim() + "' and ";
                }
            }


            //如果都沒有選擇任何一項
            if (sql.Trim().Substring(sql.Length - 6, 5) == "where")
                sql = sql.Substring(0, sql.Length - 7);
            else
                sql = sql.Substring(0, sql.Length - 4);

            //全欄位檢索
            string strAll = txtSearchAll.Text.Trim();
            if (strAll != "")
            {
                sql = "select * from (" + sql + ") as a where "
                    + "chLabNo like '%" + strAll + "%'"
                    + " or chOldLabPosition like '%" + strAll + "%'"
                    + " or chNewLabPositon like '%" + strAll + "%'"
                    + " or chSex like '%" + strAll + "%'"
                    + " or intAge like '%" + strAll + "%'"
                    + " or chLabType like '%" + strAll + "%'"
                    + " or chLabAdoptDate like '%" + strAll + "%'"
                    + " or chStoreageMethod like '%" + strAll + "%'"
                    + " or chLabLeaveBodyDatetime like '%" + strAll + "%'"
                    + " or chLabDealDatetime like '%" + strAll + "%'"
                    + " or chLabLeaveBodyEnvir like '%" + strAll + "%'"
                    + " or chLabLeaveBodyHour like '%" + strAll + "%'"
                    + " or chSubStock like '%" + strAll + "%'"
                    + " or chSickPortion like '%" + strAll + "%'"
                    + " or chDiagName1 like '%" + strAll + "%'"
                    + " or chDiagName2 like '%" + strAll + "%'"
                    + " or chDiagName3 like '%" + strAll + "%'"
                    + " or chClerkName like '%" + strAll + "%'"
                    + " or chPlanAgreeDate like '%" + strAll + "%'"
                    + " or chAgreeNoDate like '%" + strAll + "%'"
                    + " or chUseExpireDate like '%" + strAll + "%'"
                    + " or chChangeRange like '%" + strAll + "%'"
                    + " or chStatus like '%" + strAll + "%'"
                    + " or chNote like '%" + strAll + "%'"
                    + " or chTakeOutName like '%" + strAll + "%'"
                    + " or chTakeOutDate like '%" + strAll + "%'"
                    + " or chTakeOutApplicant like '%" + strAll + "%'"
                    + " or chTakeOutPlanNo like '%" + strAll + "%'"
                    + " or chTakeOutNote like '%" + strAll + "%'"
                    + " or chInComeDate like '%" + strAll + "%' order by chLabType";
             }
            else{
                sql += " order by chLabType";
            }

            
            //insert Event Log: 11. --篩選(查詢)--
            ClsShareFunc.insEvenLogt("11", ClsShareFunc.sUserName, "", "", "篩選(查詢)--");
            using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
            {
                conn.Open();

                string sLabNo, sNewLabPositon, sSex, sAge, sLabType, sLabAdoptDate, sAdoptPortion,
                    sStoreageMethod, sLabLeaveBodyDatetime, sLabDealDatetime, sLabLeaveBodyEnvir,
                    sLabLeaveBodyHour, sSubStock, sSickPortion, sDiagName1, sDiagName2, sDiagName3,
                    sClerkName, sPlanAgreeDate, sAgreeNoDate, sUseExpireDate, sChangeRange,
                    sStatus, sNote, sTakeOutName, sTakeOutDate, sTakeOutApplicant, sTakeOutPlanNo, sTakeOutNote,
                    sInComeDate, sPrintSeqNo;
                sLabNo = ""; sNewLabPositon = ""; sSex = ""; sAge = ""; sLabType = ""; sLabAdoptDate = ""; sAdoptPortion = "";
                sStoreageMethod = ""; sLabLeaveBodyDatetime = ""; sLabDealDatetime = ""; sLabLeaveBodyEnvir = "";
                sLabLeaveBodyHour = ""; sSubStock = ""; sSickPortion = ""; sDiagName1 = ""; sDiagName2 = ""; sDiagName3 = "";
                sClerkName = ""; sPlanAgreeDate = ""; sAgreeNoDate = ""; sUseExpireDate = ""; sChangeRange = "";
                sStatus = ""; sNote = ""; sTakeOutName = ""; sTakeOutDate = ""; sTakeOutApplicant = ""; sTakeOutPlanNo = ""; sTakeOutNote = "";
                sInComeDate = ""; sPrintSeqNo = "";

                try
                {
                    using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                    {
                        sCon.Open();

                        SqlCommand sCmd = new SqlCommand(sql, sCon);
                        SqlDataReader sRead = sCmd.ExecuteReader();
                        int ctRowNum = 1;
                        Dictionary<string, string> dicExpired = new Dictionary<string, string>();
                        btnSearchOut.Visible = false;

                        while (sRead.Read())
                        {
                            sLabNo = ClsShareFunc.gfunCheck(sRead["chLabNo"].ToString());
                            sNewLabPositon = ClsShareFunc.gfunCheck(sRead["chNewLabPositon"].ToString());
                            sSex = ClsShareFunc.gfunCheck(sRead["chSex"].ToString());
                            sAge = ClsShareFunc.gfunCheck(sRead["intAge"].ToString());
                            sLabType = ClsShareFunc.gfunCheck(sRead["chLabType"].ToString());
                            sLabAdoptDate = ClsShareFunc.gfunCheck(sRead["chLabAdoptDate"].ToString());
                            sAdoptPortion = ClsShareFunc.gfunCheck(sRead["chAdoptPortion"].ToString());
                            sStoreageMethod = ClsShareFunc.gfunCheck(sRead["chStoreageMethod"].ToString());
                            sLabLeaveBodyDatetime = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyDatetime"].ToString());
                            sLabDealDatetime = ClsShareFunc.gfunCheck(sRead["chLabDealDatetime"].ToString());
                            sLabLeaveBodyEnvir = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyEnvir"].ToString());
                            sLabLeaveBodyHour = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyHour"].ToString());
                            sSubStock = ClsShareFunc.gfunCheck(sRead["chSubStock"].ToString());
                            sSickPortion = ClsShareFunc.gfunCheck(sRead["chSickPortion"].ToString());
                            sDiagName1 = ClsShareFunc.gfunCheck(sRead["chDiagName1"].ToString());
                            sDiagName2 = ClsShareFunc.gfunCheck(sRead["chDiagName2"].ToString());
                            sDiagName3 = ClsShareFunc.gfunCheck(sRead["chDiagName3"].ToString());
                            sClerkName = ClsShareFunc.gfunCheck(sRead["chClerkName"].ToString());
                            sPlanAgreeDate = ClsShareFunc.gfunCheck(sRead["chPlanAgreeDate"].ToString());
                            sAgreeNoDate = ClsShareFunc.gfunCheck(sRead["chAgreeNoDate"].ToString());
                            sUseExpireDate = ClsShareFunc.gfunCheck(sRead["chUseExpireDate"].ToString());
                            sChangeRange = ClsShareFunc.gfunCheck(sRead["chChangeRange"].ToString());
                            sStatus = ClsShareFunc.gfunCheck(sRead["chStatus"].ToString());
                            sNote = ClsShareFunc.gfunCheck(sRead["chNote"].ToString());
                            sTakeOutName = ClsShareFunc.gfunCheck(sRead["chTakeOutName"].ToString());
                            sTakeOutDate = ClsShareFunc.gfunCheck(sRead["chTakeOutDate"].ToString());
                            sTakeOutApplicant = ClsShareFunc.gfunCheck(sRead["chTakeOutApplicant"].ToString());
                            sTakeOutPlanNo = ClsShareFunc.gfunCheck(sRead["chTakeOutPlanNo"].ToString());
                            sTakeOutNote = ClsShareFunc.gfunCheck(sRead["chTakeOutNote"].ToString());
                            sInComeDate = ClsShareFunc.gfunCheck(sRead["chInComeDate"].ToString());
                            sPrintSeqNo = ClsShareFunc.gfunCheck(sRead["intPrintSeqNo"].ToString());

                            dicExpired.Add(sLabNo, sUseExpireDate);
                            //dgvSearchData.Rows.Add(false, sLabNo, sNewLabPositon, sSex, sAge, sLabType, sLabAdoptDate, sAdoptPortion,
                            //    sStoreageMethod, sLabLeaveBodyDatetime, sLabDealDatetime, sLabLeaveBodyEnvir,
                            //    sLabLeaveBodyHour, sSubStock, sSickPortion, sDiagName1, sDiagName2, sDiagName3,
                            //    sClerkName, sPlanAgreeDate, sAgreeNoDate, sUseExpireDate, sChangeRange,
                            //    sStatus, sNote,sTakeOutName, sTakeOutDate, sTakeOutApplicant, sTakeOutPlanNo, sTakeOutNote,
                            //    sInComeDate, sPrintSeqNo);
                            dgvSearchData.Rows.Add(ctRowNum, false, sLabNo, sSubStock, sLabType, sStoreageMethod, sUseExpireDate, sAdoptPortion, sDiagName1, sDiagName2, sDiagName3);

                            //有選未出庫的情況
                            if (chkGetOut.Text != "")
                            {
                                string takeStr = chkGetOut.Text.ToString();
                                if (takeStr == "出庫")
                                {
                                    //MessageBox.Show(dgvSearchData.Rows[ctRowNum-1].Cells[1].Value.ToString());
                                    dgvSearchData.Rows[ctRowNum - 1].Cells[1].ReadOnly = true;
                                    dgvSearchData.Rows[ctRowNum - 1].Cells[1].Value = true;
                                }
                                else{
                                    btnSearchOut.Visible = true;
                                }
                            }
                            //沒有選出庫未出庫的情況
                            else
                            {
                                if (sTakeOutDate != "")
                                {
                                    dgvSearchData.Rows[ctRowNum - 1].Cells[1].ReadOnly = true;
                                    dgvSearchData.Rows[ctRowNum - 1].Cells[1].Value = true;
                                }
                            }

                            ctRowNum++;
                        }
                        sRead.Close();

                        if (txtSDate.Text.ToString() == "" && txtEDate.Text.ToString() == "" && dicExpired.Count != 0)
                        {
                            checkExpired(dicExpired);
                        }
                    }
                }
                 catch (Exception ex)
                {
                    MessageBox.Show("QryLReqNo: " + ex.Message.ToString());
                }
            }
        }
        //檢查那些篩選出來的檢體已過期
        private void checkExpired(Dictionary<string,string> dic)
        {
            StringBuilder sbExpired = new StringBuilder();
            Boolean expiredFlag = false;
            int count = 1;
            sbExpired.AppendLine("以下檢體編號已過期:");
            foreach (KeyValuePair<string, string> item in dic)
            {
                if (item.Value != "")
                {
                    if (Convert.ToInt32(item.Value) < Convert.ToInt32(GetTime()))
                    {
                        sbExpired.AppendLine(count + " : " + item.Key);
                        expiredFlag = true;
                    }
                }
                count++;
            }
            if (expiredFlag)
            {
                MessageBox.Show(sbExpired.ToString());
            }
        }

        ////查詢-選擇小組欄位
        //private void comboBoxCase2_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (cbGroup.Text != "")
        //    {
        //        string sql = "SELECT COLUMN_NAME as a FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME ='BioPerSlaver" + cbGroup.Text.Substring(0, 1) + "Tbl'"; ;
        //        using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
        //        {
        //            conn.Open();
        //            SqlDataAdapter combo = new SqlDataAdapter(sql, conn);
        //            DataTable dtDate = new DataTable();
        //            combo.Fill(dtDate);
        //            foreach (DataRow row in dtDate.Rows)
        //            {
        //                row[0] = row[0].ToString().Trim();
        //            }
        //            comboBoxColumn.DisplayMember = "a";
        //            comboBoxColumn.DataSource = dtDate;
        //            comboBoxColumn.SelectedIndex = -1;
        //        }
        //    }
        //}
        //查詢-列印
        private void buttonPrintfSearch_Click(object sender, EventArgs e)
        {
            //insert Event Log: 12. --篩選(列印)--
            ClsShareFunc.insEvenLogt("12", ClsShareFunc.sUserName, "", "", "篩選(列印)--");
            ClsPrint _ClsPrint = new ClsPrint(dgvSearchData, "查詢列印");
            _ClsPrint.PrintForm();
        }
        
        /*查詢 BioCommonLoginTbl 中成員資料*/
        private void cmdQuery(DataGridView dgv)
        {
            string sSQL = "";
            string sID, sName, sEnable, sDepartment, sPwd;
            sID = ""; sName = ""; sEnable = ""; sDepartment = ""; sPwd = "";

            try
            {
                //dgv.DataSource = null;
                dgv.Rows.Clear();

                //using (SqlConnection sCon = new SqlConnection(ClsShareFunc.DB_SECConnection()))
                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                {
                    sCon.Open();
                    sSQL = " select *  from BioCommonLoginTbl (nolock)  ";
                    SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                    SqlDataReader sRead = sCmd.ExecuteReader();
                    while (sRead.Read())
                    {
                        sID = ClsShareFunc.gfunCheck(sRead["chUserID"].ToString());
                        sName = ClsShareFunc.gfunCheck(sRead["chUserName"].ToString());
                        sEnable = ClsShareFunc.gfunCheck(sRead["chEnableFlag"].ToString());
                        sDepartment = ClsShareFunc.gfunCheck(sRead["chBioEmpFlag"].ToString());
                        sPwd = ClsShareFunc.gfunCheck(sRead["chPassword"].ToString());
                        dgv.Rows.Add(sID, sName, sEnable, sDepartment, sPwd);
                    }
                    sRead.Close();

                    /*SqlDataAdapter sAdp = new SqlDataAdapter(sSQL, sCon);
                    DataTable dt = new DataTable();
                    sAdp.Fill(dt);//DataAdapter Fill to DataTable
                    dgv.DataSource = dt;//DataAdapter Bind to DataGridView*/
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(" (cmdQuery): " + ex.Message.ToString());
                return;
            }
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

        /*check Id,Pwd*/
        private Boolean gfunCheckFormat(string sId, string sName, string sPwd, string sPwdVer)
        {
            Boolean Check = false;

            /*********** 1.Check 欄位不可空白 ***********/
            if (sId != "" && sName != "" && sPwd != "" && sPwdVer != "")
            {
                /*********** 2.Check 身分證正確性 ***********/
                if (gfunCheckId(sId) == true)
                {
                    /*********** 3.Check 密碼正確性 ***********/
                    if (ClsShareFunc.gfunCheckPwd(sPwd) == true && ClsShareFunc.gfunCheckPwd(sPwdVer) == true)
                    {
                        /*********** 4.Check 密碼是否一致 ***********/
                        if (sPwd == sPwdVer)
                            Check = true;
                        else
                        {
                            MessageBox.Show("密碼不一致。請重新輸入!");
                            CleartxtPwd();
                            Check = false;
                        }
                    }
                    else
                    {
                        MessageBox.Show("密碼長度需大於8且含有英文及數字!");
                        CleartxtPwd();
                        Check = false;
                    }
                }
                else
                {
                    MessageBox.Show("ID不為身分證格式!");
                    CleartxtId();
                    Check = false;
                }
            }
            else
            {
                if (sId == "")
                    MessageBox.Show("帳號不可為空白!");
                if (sName == "")
                    MessageBox.Show("姓名不可為空白!");
                if (sPwd == "")
                    MessageBox.Show("密碼不可為空白!");
                if (sPwdVer == "")
                    MessageBox.Show("密碼確認不可為空白!");
            }
            return Check;
        }

        /*主管維護 - 清除ID txtbox*/
        private void CleartxtId()
        {
            switch (this.tabForm.SelectedIndex)
            {
                case 6:
                    txtID.Text = "";
                    break;
                case 7:
                    txtID_Admin.Text = "";
                    break;
            }
        }

        /*主管維護 - 清除密碼txtbox*/
        private void CleartxtPwd()
        {
            switch (this.tabForm.SelectedIndex)
            {
                case 6:
                    txtPwd_Admin.Text = "";
                    txtPwdVer_Admin.Text = "";
                    break;
            }
        }

        /*===================檢驗身分證=================*/
        private Boolean gfunCheckId(string sId)
        {
            Boolean Check = false;
            string[] Match = { "A", "B", "C", "D", "E", "F", "G", "H", "J", "K", "L", "M", "N",
                                 "P", "Q", "R", "S", "T", "U", "V", "X", "Y", "W", "Z", "I","O" }; //10~35

            /*--------------- (1)分割成單一碼(except first chracter) ----------------*/
            int[] Charcter = new int[10];
            for (int i = 1; i < 10; i++)
                Charcter[i] = Convert.ToInt32(sId.Substring(i, 1).ToUpper());

            /*--------------- (2)Get 第一碼對應之數字 --------------------------------*/
            string FirstCha = "";
            int num = 0;
            FirstCha = sId.Substring(0, 1);
            for (int i = 10; i <= 35; i++)
                if (FirstCha == Match[i - 10])
                    num = i;

            /*--------------- (3) Check 第二碼數字(1:男/2:女) -----------------------*/
            int SecondCha = 0;
            SecondCha = Charcter[1];
            if (SecondCha == 1 || SecondCha == 2)
            {
                /*--------------- (4) Counting... ------------------------------------------------
                                               F1  F2  C1  C2  C3  C4  C5  C6  C7  C8  C9     
                                               x1   x9  x8   x7   x6   x5   x4  x3   x2   x1   keep    */

                //(4-1)先算第一碼
                int sum = 0;
                int num1 = num / 10;
                int num2 = num % 10;
                sum = num1 * 1 + num2 * 9;

                //(4-2)再算2~8碼[2~8碼 -> 8~2]
                int k = 8;
                for (int i = 1; i <= 8; i++)
                {
                    sum += Charcter[i] * k;
                    k--;
                }

                //(4-3)最後加上第9碼
                sum += Charcter[9];

                /*----------- (5 Finally: sum除以10餘0為正確身分證字號 ------------*/
                if (sum % 10 == 0)
                    Check = true;
            }
            return Check;
        }

        /*資訊室維護－dgv show to textbox*/
        private void dgvInfo_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string sId = "";
            string sName = "";
            string sDepartment = "";

            /*1.ID 2.Name 3.Enabled 4.Department 5.Pwd*/
            sId = dgvInfo.Rows[e.RowIndex].Cells[0].Value.ToString().Trim();
            sName = dgvInfo.Rows[e.RowIndex].Cells[1].Value.ToString().Trim();
            sDepartment = dgvInfo.Rows[e.RowIndex].Cells[3].Value.ToString().Trim();

            txtID.Text = sId;
            txtName.Text = sName;
            cboDepartment.Text = sDepartment;
        }

        /*=====================新增====================*/
        private void AddMemberData()
        {
            string sId = "";
            string sName = "";
            string sPwd = "";
            string sPwdVer = "";
            string sDepartment = "";
            DataGridView dgv = null;
            string sOtherValue = "";

            try
            {
                sId = txtID.Text.Trim();
                sName = txtName.Text.Trim();
                sPwd = txtPwd.Text.Trim();
                sPwdVer = txtPwdVer.Text.Trim();
                sDepartment = cboDepartment.Text.Trim();
                dgv = dgvInfo;

                if (sId.Length != 10)
                {
                    MessageBox.Show("帳號--必須為 10 碼!");
                    return;
                }
                if (sName.Length < 3)
                {
                    MessageBox.Show("姓名--必須大於 3 碼!");
                    return;
                }
                if (sPwd != sPwdVer)
                {
                    MessageBox.Show("輸入兩次的密碼不一致!");
                    return;
                }
                //function裡已經有判斷了解
                //if (sPwd.Length <8)
                //{
                //    MessageBox.Show("密碼--至少須 8 碼以上且有英文及數字!");
                //    return;
                //}

                /*1.Check 帳號and密碼格式*/
                if (gfunCheckFormat(sId, sName, sPwd, sPwdVer) == true)
                {
                    /*2.Check Administrator 中有沒有此帳號存在*/
                    if (ClsShareFunc.CheckInDb(ClsShareFunc.DbAdmin(), sId, "insert") == false)
                    {
                        /*3.Check Common 中有沒有此帳號存在*/
                        if (ClsShareFunc.CheckInDb(ClsShareFunc.DbCom(), sId, "insert") == false)
                        {
                            /*4.建立此帳號*/
                            if (VerAction("新增") == false)
                                return;

                            string insertSQL = "";
                            //using (SqlConnection insertCon = new SqlConnection(ClsShareFunc.DB_SECConnection()))
                            using (SqlConnection insertCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                            {
                                insertCon.Open();
                                insertSQL = " insert into BioCommonLoginTbl (chUserID,chUserName,chPassword,chEnableFlag,chBioEmpFlag, chCreateDateTime, chLastModPwdDT) " +
                                    " values ('" + sId + "','" + sName + "','" + ClsShareFunc.GetMD5(sPwd) + "','N','" + sDepartment + "',dbo.GetDateToDate13(getdate()), dbo.GetDateToDate13(getdate()))";
                                SqlCommand insertCmd = new SqlCommand(insertSQL, insertCon);
                                insertCmd.ExecuteNonQuery();

                                sOtherValue = " insert into BioCommonLoginTbl (chUserID,chUserName,chPassword,chEnableFlag,chBioEmpFlag, chCreateDateTime, chLastModPwdDT) " +
                                    " values (" + sId + "," + sName + "," + "PWD" + ",N," + sDepartment + ",dbo.GetDateToDate13(getdate()), dbo.GetDateToDate13(getdate()))";
                                //insert Event Log: 20. --資訊室主管新增帳號    --       
                                ClsShareFunc.insEvenLogt("20", ClsShareFunc.sUserName, "", sId, "資訊室主管新增帳號--" + sOtherValue);

                                MessageBox.Show("新增成功!");
                                InitTabMIS();
                                cmdQuery(dgv);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("身份證號--驗證錯誤!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("AddMemberData(" + this.tabForm.SelectedIndex + "): " + ex.Message.ToString());
                return;
            }
        }

        /*=====================查詢====================*/
        private void QueryMemberData()
        {
            string sId = "";
            string sName = "";
            DataGridView dgv = null;
            try
            {
                QryKey();

                switch (this.tabForm.SelectedIndex)
                {
                    case 6://主管維護                  
                        sId = txtID_Admin.Text.Trim();
                        sName = txtName_Admin.Text.Trim();
                        dgv = dgvInfo_Admin;
                        break;
                    case 7://資訊室維護
                        sId = txtID.Text.Trim();
                        sName = txtName.Text.Trim();
                        dgv = dgvInfo;
                        break;
                };

                if (sId == "")
                {
                    cmdQuery(dgv);
                }
                else
                {


                    cmdQuery(dgv);

                    //int i = 0;

                    //for (int j = 0; j < dgv.RowCount; j++)
                    //{
                    //    if (dgv.Rows[i].Cells[0].Value.ToString().Trim() != sId)
                    //        dgv.Rows.RemoveAt(i);
                    //    else
                    //        i++;
                    //}

                    //if (dgv.RowCount == 1)
                    if (dgv.RowCount == 0)
                    {
                        MessageBox.Show("無此帳號!");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("資料維護 (QueryMemberData(" + this.tabForm.SelectedIndex + ")): " + ex.Message.ToString());
                return;
            }
        }

        /*==================修改====================
                         * use: sId       update: sName,sPwd,sEnableFlg      */
        private void ModifyMemberData(int sCase)
        {
            string sId = "";
            string sName = "";
            string sPwd = "";
            string sEnableFlg = "";
            string sPwdVer = "";
            string sDepartment = "";
            string sOtherValue = "";

            try
            {
                switch (this.tabForm.SelectedIndex)
                {
                    case 6:
                        sId = txtID_Admin.Text;
                        sName = txtName_Admin.Text;
                        sPwd = txtPwd_Admin.Text;
                        sEnableFlg = cboEnableFlg_Admin.Text;
                        sPwdVer = txtPwdVer_Admin.Text;
                        break;
                    case 7:
                        sId = txtID.Text;
                        sName = txtName.Text;
                        sEnableFlg = txtEnableFlg.Text;
                        sDepartment = cboDepartment.Text;
                        break;
                };


                if (sName == "")
                {
                    MessageBox.Show("姓名不可為空白!");
                    return;
                }
                if (sName.Trim().Length > 20)
                {
                    MessageBox.Show("姓名長度不可超出20 bytes!");
                    return;
                }

                switch (sCase)
                {
                    case 1://不須修改密碼
                        if (ClsShareFunc.CheckInDb(ClsShareFunc.DbCom(), sId, "modify") == true)
                        {
                            if (VerAction("修改(一般)") == false)
                                return;

                            string modSQL = "";

                            /*2.update*/
                            using (SqlConnection modCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                            {
                                modCon.Open();
                                if (tabForm.SelectedIndex == 7) //資訊室主管可修改: 名字 & 所屬部門
                                {
                                    modSQL = "update BioCommonLoginTbl set chUserName = '" + sName
                                   + "',chBioEmpFlag = '" + sDepartment + "' where chUserID = '" + sId + "' ";
                                    //sOtherValue為event log 的OtherValue
                                    sOtherValue = "update BioCommonLoginTbl set chUserName = " + sName
                                   + ",chBioEmpFlag = " + sDepartment + " where chUserID = " + sId;
                                }
                                else if (tabForm.SelectedIndex == 6)//生物醫學主管可修改: 名字 & 權限開啟否
                                {
                                    modSQL = "update BioCommonLoginTbl set chUserName = '" + sName
                                        + "',chEnableFlag = '" + sEnableFlg + "' where chUserID = '" + sId + "' ";
                                    //sOtherValue為event log 的OtherValue
                                    sOtherValue = "update BioCommonLoginTbl set chUserName = " + sName
                                        + ",chEnableFlag = " + sEnableFlg + " where chUserID = " + sId;
                                }
                                SqlCommand modCmd = new SqlCommand(modSQL, modCon);
                                modCmd.ExecuteNonQuery();

                                //insert Event Log: 21. --資訊室主管修改帳號資料   --
                                if (tabForm.SelectedIndex == 7)  //資訊室主管可修改: 名字 & 所屬部門                             
                                    ClsShareFunc.insEvenLogt("21", ClsShareFunc.sUserName, "", sId, "資訊室主管修改帳號資料--" + sOtherValue);
                                //insert Event Log: 18. --生物主管修改帳號--
                                if (tabForm.SelectedIndex == 6)  //生物醫學主管可修改: 名字 & 權限開啟否                             
                                    ClsShareFunc.insEvenLogt("18", ClsShareFunc.sUserName, "", sId, "生物主管修改帳號--" + sOtherValue);

                                MessageBox.Show("修改成功!");
                                //999999以下,修改完存檔是應該重新query以更新spread,但會出現-無此帳號-的訊息,故先mard掉
                                QueryMemberData();
                            }
                        }
                        else
                            MessageBox.Show("無此帳號!");
                        break;

                    case 2://須修改密碼
                        /*1.Check 帳號and密碼格式*/
                        if (gfunCheckFormat(sId, sName, sPwd, sPwdVer) == true)
                        {
                            if (ClsShareFunc.CheckInDb(ClsShareFunc.DbCom(), sId, "modify") == true)
                            {
                                if (VerAction("修改(密碼)") == false)
                                    return;

                                string modSQL = "";
                                /*2.update*/
                                //using (SqlConnection modCon = new SqlConnection(ClsShareFunc.DB_SECConnection()))
                                using (SqlConnection modCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                                {
                                    modCon.Open();
                                    modSQL = "update BioCommonLoginTbl set chUserName = '" + sName + "',chPassword ='" + ClsShareFunc.GetMD5(sPwd)
                                        + "',chEnableFlag = '" + sEnableFlg + "',chLastModPwdDT = dbo.GetDateToDate13(getdate())" + " where chUserID = '" + sId + "' ";
                                    //sOtherValue為event log 的OtherValue
                                    sOtherValue = "update BioCommonLoginTbl set chUserName = " + sName + ",chPassword =" + "PWD"
                                        + ",chEnableFlag = " + sEnableFlg + " where chUserID = " + sId;
                                    SqlCommand modCmd = new SqlCommand(modSQL, modCon);
                                    modCmd.ExecuteNonQuery();

                                    //insert Event Log: 19. --生物主管修改密碼-- 
                                    ClsShareFunc.insEvenLogt("19", ClsShareFunc.sUserName, "", sId, "生物主管修改密碼--" + sOtherValue);

                                    MessageBox.Show("修改成功!");
                                    QueryMemberData();
                                }
                            }
                            else
                                MessageBox.Show("無此帳號!");
                        }
                        break;
                };

            }
            catch (Exception ex)
            {
                MessageBox.Show("ModifyMemberData(" + this.tabForm.SelectedIndex + "): " + ex.Message.ToString());
                return;
            }
        }

        /*==================切換tab時動作====================*/
        private void tabForm_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (tabForm.SelectedIndex)
            {
                case 0://資料匯入
                    //StorageTimeRecord();
                    //LoadCase(cbGroup);
                    break;
                case 1://匯入紀錄       
                    StorageTimeRecord();
                    dgvStorageRecord.DataSource = null;
                    break;
                case 2://篩選
                    searchCharge();
                    break;
                case 3://查詢/修改
                    //InitTabQryandMod();
                    break;
                case 4://出庫:
                    InitTabOut();
                    break;
                case 5://出庫紀錄
                    QryOutRecord();
                    break;
                case 6://主管維護
                    InitTabAdmin();
                    pnlKey.Enabled = true;
                    pnlAdmin.Enabled = true;
                    if (bolKeyPass == true)
                    {
                        frmVerPwd NewFrm = new frmVerPwd();
                        NewFrm.pEntrySource = "Function6";

                        NewFrm.ShowDialog();
                        if (NewFrm.PassVerPwd == true)
                        {
                            QueryMemberData();
                            bolKeyPass = false;
                        }
                        else
                        {
                            pnlKey.Enabled = false;
                            pnlAdmin.Enabled = false;
                        }
                    }
                    else
                        QueryMemberData();

                    break;
                case 7://資訊室維護
                    InitTabMIS();
                    pnlInform.Enabled = true;
                    if (bolKeyPass == true)
                    {
                        frmVerPwd NewFrmV = new frmVerPwd();
                        NewFrmV.pEntrySource = "Function7";

                        NewFrmV.ShowDialog();
                        if (NewFrmV.PassVerPwd == true)
                        {
                            QueryMemberData();
                            bolKeyPass = false;
                        }
                        else
                            pnlInform.Enabled = false;
                    }
                    else
                        QueryMemberData();
                    break;
                case 8: //特殊權限
                    InitTabAuth();
                    gbID2LReqNo.Enabled = false;
                    gbLReqNo2All.Enabled = false;

                    //每次點特殊權限Tab, 每次都需問 Administrator password
                    pFunction8_AdminID = "";
                    frmVerPwd NewFrmA = new frmVerPwd();
                    NewFrmA.pEntrySource = "Function8";
                    NewFrmA.ShowDialog();

                    if (NewFrmA.PassVerPwd == true)
                    {
                        gbID2LReqNo.Enabled = true;
                        gbLReqNo2All.Enabled = true;

                    }


                    break;
                case 9://備份
                    gbRemote.Enabled = true;
                    gbLocal.Enabled = true;
                    GetLocalEnv();
                    GetRemoteEnv();
                    break;
                case 10: //LOG
                    InitTabAuth();
                    txtEventDate.Enabled = true;
                    txtEventNo.Enabled = true;
                    Refresh.Enabled = true;
                    if (bolKeyPass == true)
                    {
                        frmVerPwd NewFrmB = new frmVerPwd();
                        NewFrmB.pEntrySource = "Function10";
                        NewFrmB.ShowDialog();
                        if (NewFrmB.PassVerPwd == true)
                        {
                            bolKeyPass = false;
                        }
                        else
                        {
                            txtEventDate.Enabled = false;
                            txtEventNo.Enabled = false;
                            Refresh.Enabled = false;
                        }
                    }
                    break;
            };
        }

        /*=======================初始介面==========================*/
        /*清除介面 - 查詢/修改*/
        private void InitTabQryandMod()
        {
            txtModLReqNo.Text = "";
            CleardgvQryLReqNo();
            CleardgvShowLReqNo();
        }

        /*清除介面 - 出庫*/
        private void InitTabOut()
        {
            //txtOutLReqNo.Text = "";
            richtxtMsg.Text = "";
            //CleardgvOutLReqNo();
            for (int j = 1; j <= 31; j++)
            {
                switch (j)
                {
                    case 0:
                    case 1:
                    case 27:
                    case 28:
                    case 29:
                        dgvOutLReqNo.Columns[j].Visible = true;
                        break;
                    default:
                        dgvOutLReqNo.Columns[j].Visible = false;
                        break;
                };
            }
            CleardgvOutDetail();
        }

        /*清除介面 - 主管維護*/
        private void InitTabAdmin()
        {
            cboInterface.Text = "";
            txtID_Admin.Text = "";
            txtName_Admin.Text = "";
            txtPwd_Admin.Text = "";
            txtPwdVer_Admin.Text = "";
            cboEnableFlg_Admin.Text = "N";

            CleardgvInfo_Admin();
            dgvKey.DataSource = null;
            lblPwd_Admin.Visible = false;
            lblPwdVer_Admin.Visible = false;
            txtPwd_Admin.Visible = false;
            txtPwdVer_Admin.Visible = false;
        }

        /*清除介面 - 資訊室維護*/
        private void InitTabMIS()
        {
            txtID.Text = "";
            txtName.Text = "";
            txtPwd.Text = "";
            txtPwdVer.Text = "";
            cboDepartment.Text = "B";
            txtEnableFlg.Text = "N";

            txtID.Enabled = false;

            lblPwd.Visible = false;
            lblPwdVer.Visible = false;
            lblEnableFlg.Visible = false;
            txtPwd.Visible = false;
            txtPwdVer.Visible = false;
            txtEnableFlg.Visible = false;
            btnAdd.Visible = false;
            btnCancel.Visible = false;
            lnklblAddMember.Visible = true;
            btnModify.Visible = true;

            CleardgvInfo();
        }

        /*清除介面 - 特殊權限作業*/
        private void InitTabAuth()
        {
            txtID2LReqNo.Text = "";
            txtLReqNo2All.Text = "";
            CleardgvID2LReqNo();
            CleardgvLReqNo2All();
            bolClear = true;
        }

        /*=======================查詢/修改==========================*/
        /*查詢/修改 - use LReqNo -> data*/
        private void QryLReqNo(string sLReqNo, DataGridView dgv)
        {
            string sSQL = "";
            string sLabNo, sNewLabPositon, sSex, sAge, sLabType, sLabAdoptDate, sAdoptPortion,
                sStoreageMethod, sLabLeaveBodyDatetime, sLabDealDatetime, sLabLeaveBodyEnvir,
                sLabLeaveBodyHour, sSubStock, sSickPortion, sDiagName1, sDiagName2, sDiagName3,
                sClerkName, sPlanAgreeDate, sAgreeNoDate, sUseExpireDate, sChangeRange,
                sStatus, sNote, sTakeOutName, sTakeOutDate, sTakeOutApplicant, sTakeOutPlanNo, sTakeOutNote,
                sInComeDate, sPrintSeqNo;
            sLabNo = ""; sNewLabPositon = ""; sSex = ""; sAge = ""; sLabType = ""; sLabAdoptDate = ""; sAdoptPortion = "";
            sStoreageMethod = ""; sLabLeaveBodyDatetime = ""; sLabDealDatetime = ""; sLabLeaveBodyEnvir = "";
            sLabLeaveBodyHour = ""; sSubStock = ""; sSickPortion = ""; sDiagName1 = ""; sDiagName2 = ""; sDiagName3 = "";
            sClerkName = ""; sPlanAgreeDate = ""; sAgreeNoDate = ""; sUseExpireDate = ""; sChangeRange = "";
            sStatus = ""; sNote = ""; sTakeOutName = ""; sTakeOutDate = ""; sTakeOutApplicant = ""; sTakeOutPlanNo = ""; sTakeOutNote = "";
            sInComeDate = ""; sPrintSeqNo = "";

            try
            {
                switch (dgv.Name)
                {
                    case "dgvQryLReqNo":
                        CleardgvQryLReqNo();
                        break;
                    case "dgvOutLReqNo":
                        CleardgvOutLReqNo();
                        break;
                }

                sLReqNo = sLReqNo.ToUpper();
                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                {
                    sCon.Open();
                    sSQL = "select *  from BioPerMasterTbl (nolock) where chLabNo like '" + sLReqNo + "%'";

                    SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                    SqlDataReader sRead = sCmd.ExecuteReader();
                    while (sRead.Read())
                    {
                        sLabNo = ClsShareFunc.gfunCheck(sRead["chLabNo"].ToString());
                        sNewLabPositon = ClsShareFunc.gfunCheck(sRead["chNewLabPositon"].ToString());
                        sSex = ClsShareFunc.gfunCheck(sRead["chSex"].ToString());
                        sAge = ClsShareFunc.gfunCheck(sRead["intAge"].ToString());
                        sLabType = ClsShareFunc.gfunCheck(sRead["chLabType"].ToString());
                        sLabAdoptDate = ClsShareFunc.gfunCheck(sRead["chLabAdoptDate"].ToString());
                        sAdoptPortion = ClsShareFunc.gfunCheck(sRead["chAdoptPortion"].ToString());
                        sStoreageMethod = ClsShareFunc.gfunCheck(sRead["chStoreageMethod"].ToString());
                        sLabLeaveBodyDatetime = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyDatetime"].ToString());
                        sLabDealDatetime = ClsShareFunc.gfunCheck(sRead["chLabDealDatetime"].ToString());
                        sLabLeaveBodyEnvir = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyEnvir"].ToString());
                        sLabLeaveBodyHour = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyHour"].ToString());
                        sSubStock = ClsShareFunc.gfunCheck(sRead["chSubStock"].ToString());
                        sSickPortion = ClsShareFunc.gfunCheck(sRead["chSickPortion"].ToString());
                        sDiagName1 = ClsShareFunc.gfunCheck(sRead["chDiagName1"].ToString());
                        sDiagName2 = ClsShareFunc.gfunCheck(sRead["chDiagName2"].ToString());
                        sDiagName3 = ClsShareFunc.gfunCheck(sRead["chDiagName3"].ToString());
                        sClerkName = ClsShareFunc.gfunCheck(sRead["chClerkName"].ToString());
                        sPlanAgreeDate = ClsShareFunc.gfunCheck(sRead["chPlanAgreeDate"].ToString());
                        sAgreeNoDate = ClsShareFunc.gfunCheck(sRead["chAgreeNoDate"].ToString());
                        sUseExpireDate = ClsShareFunc.gfunCheck(sRead["chUseExpireDate"].ToString());
                        sChangeRange = ClsShareFunc.gfunCheck(sRead["chChangeRange"].ToString());
                        sStatus = ClsShareFunc.gfunCheck(sRead["chStatus"].ToString());
                        sNote = ClsShareFunc.gfunCheck(sRead["chNote"].ToString());
                        sTakeOutName = ClsShareFunc.gfunCheck(sRead["chTakeOutName"].ToString());
                        sTakeOutDate = ClsShareFunc.gfunCheck(sRead["chTakeOutDate"].ToString());
                        sTakeOutApplicant = ClsShareFunc.gfunCheck(sRead["chTakeOutApplicant"].ToString());
                        sTakeOutPlanNo = ClsShareFunc.gfunCheck(sRead["chTakeOutPlanNo"].ToString());
                        sTakeOutNote = ClsShareFunc.gfunCheck(sRead["chTakeOutNote"].ToString());
                        sInComeDate = ClsShareFunc.gfunCheck(sRead["chInComeDate"].ToString());
                        sPrintSeqNo = ClsShareFunc.gfunCheck(sRead["intPrintSeqNo"].ToString());

                        dgv.Rows.Add(false, sLabNo, sNewLabPositon, sSex, sAge, sLabType, sLabAdoptDate, sAdoptPortion,
                            sStoreageMethod, sLabLeaveBodyDatetime, sLabDealDatetime, sLabLeaveBodyEnvir,
                            sLabLeaveBodyHour, sSubStock, sSickPortion, sDiagName1, sDiagName2, sDiagName3,
                            sClerkName, sPlanAgreeDate, sAgreeNoDate, sUseExpireDate, sChangeRange,
                            sStatus, sNote, sTakeOutName, sTakeOutDate, sTakeOutApplicant, sTakeOutPlanNo, sTakeOutNote,
                            sInComeDate, sPrintSeqNo);
                    }
                    sRead.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("QryLReqNo: " + ex.Message.ToString());
            }
        }

        /*查詢/修改 - 輸入檢體編號時動作*/
        private void txtModLReqNo_TextChanged(object sender, EventArgs e)
        {
            string sLReqNo = "";
            sLReqNo = txtModLReqNo.Text;
            if (sLReqNo.Length > 0)
                QryLReqNo(sLReqNo, dgvQryLReqNo);
            else
                dgvQryLReqNo.Rows.Clear();
            dgvShowLReqNo.Rows.Clear();
        }

        /*查詢/修改 - 修改檢體資訊*/
        private void btnModLReqNo_Click(object sender, EventArgs e)
        {

            string tmpLReqNo = "";
            string sLReqNo = ""; //不變
            string sYear = "";
            string sRange = "";
            string sStatus = "";
            string sNote = "";
            string sSQL = "";
            string sPosition = "";
            Dictionary<string, string> dicPos = new Dictionary<string, string>();
            ArrayList arrPos = new ArrayList();
            try
            {

                tmpLReqNo = txtModLReqNo.Text;
                sLReqNo = dgvShowLReqNo.Rows[0].Cells[0].Value.ToString().Trim();
                sPosition = dgvShowLReqNo.Rows[0].Cells[1].Value == null ? "" : dgvShowLReqNo.Rows[0].Cells[1].Value.ToString().Trim();
                sYear = dgvShowLReqNo.Rows[0].Cells[2].Value.ToString().Trim();
                sRange = dgvShowLReqNo.Rows[0].Cells[3].Value.ToString().Trim();
                sStatus = dgvShowLReqNo.Rows[0].Cells[4].Value.ToString().Trim();
                sNote = dgvShowLReqNo.Rows[0].Cells[5].Value.ToString().Trim();
                if (sPosition == "")
                {
                    MessageBox.Show("檢體位置不得為空白!");
                    return;
                }
                else
                {
                    if (sStatus.Length > 0 && sStatus == "退出")
                    {
                        sPosition = sPosition + "(退)";
                    }
                    dicPos.Add("1", sPosition);
                    arrPos = checkPos(dicPos);
                    if (arrPos.Count > 0)
                    {
                        MessageBox.Show("檢體位置重複!");
                        return;
                    }
                }

                if (sYear.Length > 7)
                {
                    MessageBox.Show("使用年限--長度不可超出7 bytes!");
                    return;
                }
                if (sRange.Length > 100)
                {
                    MessageBox.Show("變更範圍--長度不可超出100 bytes!");
                    return;
                }
                if (sRange.Length > 30)
                {
                    MessageBox.Show("狀態--長度不可超出30 bytes!");
                    return;
                }
                if (sNote.Length > 100)
                {
                    MessageBox.Show("備註--長度不可超出100 bytes!");
                    return;
                }
                

                //insert Event Log: 14. --修改檢體資料--
                ClsShareFunc.insEvenLogt("14", ClsShareFunc.sUserName, sLReqNo, "", "修改檢體資料--");
                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                {
                    sCon.Open();
                    sSQL = "update BioPerMasterTbl set chNewLabPositon='" + sPosition + "', chUseExpireDate='" + sYear + "' , chChangeRange='" + sRange + "' , chStatus='" + sStatus;
                    sSQL = sSQL + "' , chNote='" + sNote + "' where chLabNo='" + sLReqNo + "'";

                    SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                    sCmd.ExecuteNonQuery();

                    dgvShowLReqNo.Rows.Clear();
                    dgvQryLReqNo.Rows.Clear();
                    QryLReqNo(tmpLReqNo, dgvQryLReqNo);
                }
                MessageBox.Show("修改完成!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("buttonModify_Click: " + ex.Message.ToString());
            }
        }

        /*查詢/修改 - 帶入出庫畫面*/
        private void btn2Out_Click(object sender, EventArgs e)
        {
            tabForm.SelectedIndex = 4;
            dgvOutLReqNo.Rows.Clear();
            string[] aStr = new string[30];
            for (int i = 0; i < dgvQryLReqNo.Rows.Count; i++)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)dgvQryLReqNo.Rows[i].Cells[0];
                if (chk.Value.ToString() == "True")
                {
                    for (int j = 1; j < 30; j++)
                    {
                        aStr[j] = dgvQryLReqNo.Rows[i].Cells[j].Value.ToString();
                    }
                    dgvOutLReqNo.Rows.Add(aStr);
                }
            }
        }
        //
        private void dgvQryLReqNo_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                string sLReqNo = "";
                string sYear = "";
                string sRange = "";
                string sStatus = "";
                string sNote = "";
                string sPosition = "";
                int sRow = e.RowIndex;

                CleardgvShowLReqNo();
                if (e.ColumnIndex != 0 && sRow >= 0)
                {
                    sLReqNo = dgvQryLReqNo.Rows[sRow].Cells[1].Value.ToString();
                    sPosition = dgvQryLReqNo.Rows[sRow].Cells[2].Value.ToString();
                    sYear = dgvQryLReqNo.Rows[sRow].Cells[21].Value.ToString();
                    sRange = dgvQryLReqNo.Rows[sRow].Cells[22].Value.ToString();
                    sStatus = dgvQryLReqNo.Rows[sRow].Cells[23].Value.ToString().Trim();
                    sNote = dgvQryLReqNo.Rows[sRow].Cells[24].Value.ToString();

                    dgvShowLReqNo.Rows.Add(sLReqNo, sPosition, sYear, sRange, sStatus, sNote);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("dgvQryLReqNo_RowHeaderMouseDoubleClick: " + ex.Message.ToString());
            }
        }
        /*查詢/修改 - cell2dgv*/
        private void dgvQryLReqNo_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            
        }

        /*查詢/修改 - 清除DataGridView(dgvQryLReqNo)*/
        private void CleardgvQryLReqNo()
        {
            dgvQryLReqNo.Rows.Clear();
        }

        /*查詢/修改 - 清除DataGridView(dgvShowLReqNo)*/
        private void CleardgvShowLReqNo()
        {
            dgvShowLReqNo.Rows.Clear();
        }

        /*查詢/修改 - 全選*/
        private void chkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dgvQryLReqNo.Rows.Count; i++)
                dgvQryLReqNo.Rows[i].Cells[0].Value = chkSelectAll.Checked;
        }
        /*=======================出庫==========================*/
        /*出庫 - cell2dgv*/
        private void dgvOutLReqNo_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                int sRow = e.RowIndex;
                string[] aStr1 = new string[10];
                string[] aStr2 = new string[10];
                string[] aStr3 = new string[10];

                CleardgvOutDetail();

                for (int j = 1; j < dgvOutLReqNo.Columns.Count - 1; j++)
                {
                    if (j >= 1 && j <= 10)
                        aStr1[j - 1] = dgvOutLReqNo.Rows[sRow].Cells[j].Value == null ? "" : dgvOutLReqNo.Rows[sRow].Cells[j].Value.ToString();
                    else if (j >= 10 && j <= 20)
                        aStr2[j - 11] = (dgvOutLReqNo.Rows[sRow].Cells[j].Value == null ? "" : dgvOutLReqNo.Rows[sRow].Cells[j].Value.ToString());
                    else if (j >= 20 && j <= 31)
                        aStr3[j - 21] = (dgvOutLReqNo.Rows[sRow].Cells[j].Value == null ? "" : dgvOutLReqNo.Rows[sRow].Cells[j].Value.ToString());
                }

                dgvOutDetail1.Rows.Add(aStr1);
                dgvOutDetail2.Rows.Add(aStr2);
                dgvOutDetail3.Rows.Add(aStr3);
            }
        }

        /*出庫 - 將check的LReqNo出庫*/
        private void btnOutLReqNo_Click(object sender, EventArgs e)
        {
            string sSQL = "";
            string sSQLOut = "";
            string sLReqNo = "";
            string sDBTakeOutDate = "";
            string sTakeOutName = "";
            string sTakeOutDate = "";
            string sTakeOutApplicant = "";
            string sTakeOutPlanNo = "";
            string sTakeOutNote = "";
            int sStartIndex = 0;
            int sEndIndex = 0;

            /*1. 清空detail畫面*/
            InitTabOut();

            //要填入的出庫人名字
            sTakeOutName = ClsShareFunc.sUserName;
            //sTakeOutDate = ClsShareFunc.ChangeDateFormat(GetTime());
            //要填入的出庫日期(今天的yyymmdd)
            sTakeOutDate = GetTime();

            for (int i = 0; i < dgvOutLReqNo.Rows.Count; )
            {
                sLReqNo = dgvOutLReqNo.Rows[i].Cells[1].Value.ToString().Trim();
                sTakeOutApplicant = dgvOutLReqNo.Rows[i].Cells[27].Value.ToString().Trim();
                sTakeOutPlanNo = dgvOutLReqNo.Rows[i].Cells[28].Value.ToString().Trim();
                sTakeOutNote = dgvOutLReqNo.Rows[i].Cells[29].Value.ToString().Trim();

                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)dgvOutLReqNo.Rows[i].Cells[0];
                string sChkValue = "";
                if (chk.Value == null)
                    sChkValue = "";
                else
                    sChkValue = chk.Value.ToString();

                if (sChkValue == "True") //有勾選就出庫
                {
                    /*1.先check是否出庫過*/
                    using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                    {
                        sCon.Open();
                        sSQL = "select chTakeOutDate from BioPerMasterTbl (nolock) where chLabNo = '" + sLReqNo + "' ";
                        SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                        SqlDataReader sRead = sCmd.ExecuteReader();

                        while (sRead.Read())
                        {
                            sDBTakeOutDate = ClsShareFunc.gfunCheck(sRead["chTakeOutDate"].ToString().Trim());
                        }
                        sRead.Close();

                        if (sDBTakeOutDate == "")
                        {
                            /*2.check應填欄位不可空白*/
                            if (sTakeOutApplicant != "" && sTakeOutPlanNo != "")
                            {
                                /*3.出庫*/
                                //insert Event Log: 15. --修改檢體資料--
                                ClsShareFunc.insEvenLogt("15", ClsShareFunc.sUserName, sLReqNo, "", "檢體出庫--");
                                using (SqlConnection sConOut = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                                {
                                    sConOut.Open();
                                    sSQLOut = "update BioPerMasterTbl set chTakeOutDate = '" + sTakeOutDate + "' , chTakeOutApplicant = '" + sTakeOutApplicant + "' ,";
                                    sSQLOut = sSQLOut + "chTakeOutPlanNo = '" + sTakeOutPlanNo + "' , chTakeOutNote = '" + sTakeOutNote + "' , chTakeOutName = '" + sTakeOutName + "' ";
                                    sSQLOut = sSQLOut + " where chLabNo = '" + sLReqNo + "' ";
                                    SqlCommand sCmdOut = new SqlCommand(sSQLOut, sConOut);
                                    sCmdOut.ExecuteNonQuery();

                                    sStartIndex = richtxtMsg.Text.Length;
                                    richtxtMsg.Text = richtxtMsg.Text + "檢體編號【" + sLReqNo + "】已出庫\r\n";
                                    sEndIndex = richtxtMsg.Text.Length;
                                    richtxtMsg.Select(sStartIndex, sEndIndex);
                                    richtxtMsg.SelectionColor = System.Drawing.Color.Black;
                                    dgvOutLReqNo.Rows.RemoveAt(i);
                                }
                            }
                            else
                            {
                                MessageBox.Show("必填欄位不可空白!");
                                break;
                            }
                        }
                        else
                        {
                            //insert Event Log: 15-1. --檢體出庫 (Fail 已曾出庫過)--
                            ClsShareFunc.insEvenLogt("15-1", ClsShareFunc.sUserName, sLReqNo, "", "檢體出庫 (Fail 已曾出庫過)--");
                            sStartIndex = richtxtMsg.Text.Length;
                            richtxtMsg.Text = richtxtMsg.Text + "檢體編號【" + sLReqNo + "】已出庫過!\r\n";
                            sEndIndex = richtxtMsg.Text.Length;
                            richtxtMsg.Select(sStartIndex, sEndIndex);
                            richtxtMsg.SelectionColor = System.Drawing.Color.Red;
                            dgvOutLReqNo.Rows[i].Cells[0].Value = false;
                            i++;
                        }
                    }
                }
                else
                {
                    i++;
                }
            }


        }

        /*出庫 - copy to dt*/
        private DataTable Dgv2Dt(DataGridView dgv)
        {
            DataTable dt = new DataTable();
            /*1. add header & columns*/
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                DataColumn dc = new DataColumn();
                dc.ColumnName = dgv.Columns[i].HeaderText.ToString();
                dt.Columns.Add(dc);
            }

            /*2. add rows*/
            for (int i = 0; i < dgv.Rows.Count - 1; i++)
            {
                DataRow dr = dt.NewRow();
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    dr[j] = dgv.Rows[i].Cells[j].Value;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        /*出庫 - 輸入檢體編號時動作*/
        private void txtOutLReqNo_TextChanged(object sender, EventArgs e)
        {
            string sLReqNo = "";
            sLReqNo = txtOutLReqNo.Text;
            if (sLReqNo.Length > 0)
                QryLReqNo(sLReqNo, dgvOutLReqNo);
            else
                dgvOutLReqNo.Rows.Clear();
        }

        /*出庫 - 清除DataGridView(dgvOutLReqNo)*/
        private void CleardgvOutLReqNo()
        {
            dgvOutLReqNo.Rows.Clear();
            for (int j = 1; j <= 30; j++)
            {
                switch (j)
                {
                    case 0:
                    case 1:
                    case 27:
                    case 28:
                    case 29:
                        dgvOutLReqNo.Columns[j].Visible = true;
                        break;
                    default:
                        dgvOutLReqNo.Columns[j].Visible = false;
                        break;
                };
            }
        }

        /*出庫 - 清除DataGridView(dgvOutDetail)*/
        private void CleardgvOutDetail()
        {
            dgvOutDetail1.Rows.Clear();
            dgvOutDetail2.Rows.Clear();
            dgvOutDetail3.Rows.Clear();
        }

        /*出庫 - 右鍵移除row*/
        private void dgvOutLReqNo_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                dgvOutLReqNo.Rows.RemoveAt(e.RowIndex);
            }
        }

        /*出庫 - 全選*/
        private void chkSelectAllOut_CheckedChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < dgvOutLReqNo.Rows.Count; i++)
                dgvOutLReqNo.Rows[i].Cells[0].Value = chkSelectAllOut.Checked;
        }

        /*=======================主管維護=========================*/

        /*------------------------------金鑰管理------------------------------*/

        private void QryKey()
        {
            DataGridView dgv = null;

            switch (tabForm.SelectedIndex)
            {
                case 6://主管維護
                    dgv = dgvKey;
                    break;
                case 7://資訊室維護
                    dgv = dgvKey_Inform;
                    break;
            };

            string sSQL = "";
            string sYear = "";
            string sGenerate = "";

            try
            {
                dgv.Rows.Clear();

                //using (SqlConnection sCon = new SqlConnection(ClsShareFunc.DB_SECConnection()))
                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                {
                    sCon.Open();
                    sSQL = " select chYear,是否已產生='Y' from BioMasterKeyTbl (nolock) ";
                    SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                    SqlDataReader sRead = sCmd.ExecuteReader();
                    while (sRead.Read())
                    {
                        sYear = ClsShareFunc.gfunCheck(sRead["chYear"].ToString());
                        sGenerate = ClsShareFunc.gfunCheck(sRead["是否已產生"].ToString());
                        dgv.Rows.Add(sYear, sGenerate);
                    }
                    sRead.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("QryKey(" + tabForm.SelectedIndex + ")" + ex.Message.ToString());
                return;
            }
        }

        /*帳號管理 - 查詢金鑰*/
        private void btnQryKey_Click(object sender, EventArgs e)
        {
            QryKey();
        }

        /*帳號管理 - 新增金鑰*/
        private void btnAddKey_Click(object sender, EventArgs e)
        {
            string sYear = "";
            string sKey = "";
            string sSQL = "";

            try
            {
                //using (SqlConnection sCon = new SqlConnection(ClsShareFunc.DB_SECConnection()))
                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                {
                    sCon.Open();
                    sSQL = "select top 1 chYear from BioMasterKeyTbl (nolock) order by chYear desc";
                    SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                    SqlDataReader sRead = sCmd.ExecuteReader();
                    if (sRead.HasRows)
                        while (sRead.Read())
                            sYear = ClsShareFunc.gfunCheck(sRead["chYear"]).Trim();
                    sRead.Close();
                }

                if (sYear != "")
                {
                    sYear = (Convert.ToInt32(sYear) + 1).ToString();
                    sKey = AES.GetMD5(sYear + "TzuchiBiology");


                    //insert Event Log: 17. --新增年度金鑰  --
                    ClsShareFunc.insEvenLogt("17", ClsShareFunc.sUserName, "", "", "新增年度金鑰--" + sYear + "年度");
                    using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                    {
                        sCon.Open();
                        sSQL = "insert into BioMasterKeyTbl (chYear,chMasterKey) values " +
                            "('" + sYear + "','" + sKey + "')";
                        SqlCommand insertCmd = new SqlCommand(sSQL, sCon);
                        insertCmd.ExecuteNonQuery();

                        MessageBox.Show("新增成功!");
                    }
                }

                QryKey();
            }
            catch (Exception ex)
            {
                //insert Event Log: 17-1. --新增年度金鑰 (Fail)  --
                ClsShareFunc.insEvenLogt("17-1", ClsShareFunc.sUserName, "", "", "新增年度金鑰 (Fail)--" + sYear + "年度");
                MessageBox.Show("主管資料維護 (btnAddKey_Click): " + ex.Message.ToString());
                return;
            }
        }

        /*------------------------------帳號管理------------------------------*/

        /*帳號管理－修改*/
        private void btnMod_Admin_Click(object sender, EventArgs e)
        {
            if (cboInterface.Text == "修改(密碼)")
                ModifyMemberData(2);//須修改密碼
            else
                ModifyMemberData(1);
        }

        /*帳號管理－dgv show to textbox*/
        private void dgvInfo_Admin_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string sId = "";
            string sName = "";
            string sEnableFlg = "";

            sId = dgvInfo_Admin.Rows[e.RowIndex].Cells[0].Value.ToString().Trim();
            sName = dgvInfo_Admin.Rows[e.RowIndex].Cells[1].Value.ToString().Trim();
            sEnableFlg = dgvInfo_Admin.Rows[e.RowIndex].Cells[2].Value.ToString().Trim();

            txtID_Admin.Text = sId;
            txtName_Admin.Text = sName;
            cboEnableFlg_Admin.Text = sEnableFlg;
        }

        /*帳號管理 - 介面選擇1*/
        private void cboInterface_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sInterface = "";
            sInterface = cboInterface.Text;
            SwitchcboInterfate(sInterface);
        }

        /*帳號管理 - 介面選擇2*/
        private void SwitchcboInterfate(string sInterface)
        {
            /*--------------------------default:---------------------------------------*/
            btnMod_Admin.Visible = false;

            lblId_Admin.Visible = true;
            lblName_Admin.Visible = true;
            lblEnableFlg_Admin.Visible = true;
            lblPwd_Admin.Visible = false;
            lblPwdVer_Admin.Visible = false;

            txtID_Admin.Visible = true;
            txtName_Admin.Visible = true;
            cboEnableFlg_Admin.Visible = true;
            txtPwd_Admin.Visible = false;
            txtPwdVer_Admin.Visible = false;
            /*----------------------------------:---------------------------------------*/

            switch (sInterface)
            {
                case "修改(一般)":
                    btnMod_Admin.Visible = true;
                    break;

                case "修改(密碼)":
                    btnMod_Admin.Visible = true;
                    lblPwd_Admin.Visible = true;
                    lblPwdVer_Admin.Visible = true;
                    txtPwd_Admin.Visible = true;
                    txtPwdVer_Admin.Visible = true;
                    break;

                default:
                    btnMod_Admin.Visible = false;
                    break;
            };
        }

        /*帳號管理 - 清除DataGridView(dgvInfo_Admin)*/
        private void CleardgvInfo_Admin()
        {
            dgvInfo_Admin.DataSource = null;
        }

        /*=======================資訊室維護=======================*/
        /*資訊維護－查詢*/
        private void btnQuery_Click(object sender, EventArgs e)
        {
            QueryMemberData();
        }

        /*資訊維護－新增*/
        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddMemberData();
        }

        /*資訊維護－修改*/
        private void btnModify_Click(object sender, EventArgs e)
        {
            ModifyMemberData(1);//不須修改密碼
        }

        /*資訊維護－新增成員*/
        private void lnklblAddMember_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            txtID.Enabled = true;

            lnklblAddMember.Visible = false;
            btnModify.Visible = false;
            btnAdd.Visible = true;
            btnCancel.Visible = true;
            lblPwd.Visible = true;
            lblPwdVer.Visible = true;
            lblEnableFlg.Visible = true;
            txtPwd.Visible = true;
            txtPwdVer.Visible = true;
            txtEnableFlg.Visible = true;

            txtID.Text = "";
            txtName.Text = "";
            txtEnableFlg.Text = "N";
            txtPwd.Text = "";
            txtPwdVer.Text = "";
            cboDepartment.Text = "B";
        }

        /*資訊維護－清除DataGridView(dgvInfo)*/
        private void CleardgvInfo()
        {
            dgvInfo.DataSource = null;
        }

        /*資訊維護－清除DataGridView(Key_Inform)*/
        private void CleardgvdgvKey_Inform()
        {
            dgvKey_Inform.DataSource = null;
        }

        /*資訊維護－清除DataGridView(Key_Inform)*/
        private void btnCancel_Click(object sender, EventArgs e)
        {
            InitTabMIS();
        }


        /*======================特殊權限作業======================*/

        /*特殊權限作業 - ID -> LReqNo*/
        private void txtID2LReqNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.ToString() == Keys.Enter.ToString())
            {
                string sMRNo = "";
                string sSQL = "";
                string sOriMRNo = "";
                string sDeMRNo = "";
                string sLReqNo = "";
                string sInComYear = "";

                try
                {
                    /*1.Initialize DataGridView*/
                    CleardgvID2LReqNo();

                    txtID2LReqNo.Text = txtID2LReqNo.Text.Trim().ToUpper();
                    sMRNo = txtID2LReqNo.Text;

                    if (sMRNo != "")
                    {
                        /*2. Query Data*/
                        //using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                        //insert Event Log: 22. --特殊權限--查詢病歷號碼  --
                        ClsShareFunc.insEvenLogt("22", ClsShareFunc.sUserName, "", sMRNo, "特殊權限(查詢病歷號碼)-- (Cosign Administrator: " + pFunction8_AdminID.Trim());
                        using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                        {
                            sCon.Open();
                            sSQL = "SELECT 病歷號 = b.chMRNo, 檢體編號 = a.chLabNo,chInComeDate  FROM [DB_BIO].[dbo].[BioPerMasterTbl] a (nolock) ";
                            sSQL = sSQL + " inner join [DB_SEC].[dbo].[BioPerMappingTbl] b (nolock) ";
                            sSQL = sSQL + " on a.chLabNo = b.chLabNo collate Chinese_Taiwan_Stroke_CI_AS ";
                            SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                            SqlDataReader sRead = sCmd.ExecuteReader();

                            if (sRead.HasRows)
                            {
                                while (sRead.Read())
                                {
                                    sLReqNo = ClsShareFunc.gfunCheck(sRead["檢體編號"].ToString().Trim());
                                    sOriMRNo = ClsShareFunc.gfunCheck(sRead["病歷號"].ToString().Trim());
                                    sInComYear = ClsShareFunc.gfunCheck(sRead["chInComeDate"].ToString().Trim()).Substring(0, 3);
                                    sDeMRNo = dec_AES(sOriMRNo, sLReqNo, sInComYear);
                                    if (sDeMRNo == "")
                                    {
                                        MessageBox.Show("解密時出現錯誤, 請通知資訊人員!");
                                        return;
                                    }
                                    /*3. Get Data of the MRNo*/
                                    if (sMRNo == sDeMRNo)
                                        dgvID2LReqNo.Rows.Add(sDeMRNo, sLReqNo);
                                }
                                CleargbLReqNo2All();
                            }
                            else
                                MessageBox.Show("無此病歷號之相關資訊!");
                            sRead.Close();
                        }
                    }
                    else
                        MessageBox.Show("請輸入病歷號!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("txtID2LReqNo_KeyDown: " + ex.Message.ToString());
                }
            }
        }

        /*特殊權限作業 - Enter */
        private void txtLReqNo2All_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.ToString() == Keys.Enter.ToString())
            {
                txtLReqNo2All.Text = txtLReqNo2All.Text.ToString().Trim().ToUpper();
                string sLReqNo = "";
                sLReqNo = txtLReqNo2All.Text;
                QryLReqNo2All(sLReqNo);
            }
        }

        /*特殊權限作業 - LReqNo -> All Data */
        private void QryLReqNo2All(string sLReqNo)
        {
            string sSQL = "";
            string sYear = "";

            string sLabPieNo, sLabNo, sOldLabPosition, sNewLabPositon, sPerCaseNo, sOldMRNo, sMRNo,
                sPtName, sSex, sBirthday, sAge, sLabType, sLabAdoptDate, sAdoptPortion, sStoreageMethod,
                sLabLeaveBodyDatetime, sLabDealDatetime, sLabLeaveBodyEnvir, sLabLeaveBodyHour,
                sSubStock, sSickPortion, sDiagName1, sDiagName2, sDiagName3,
                sClerkName, sPlanAgree, sPlanAgreeDate, sAgreeNo, sAgreeNoDate,
                sUseExpireDate, sChangeRange, sStatus, sNote, sTakeOutName, sTakeOutDate,
                sTakeOutApplicant, sTakeOutPlanNo, sTakeOutNote, sInComeDate, sPrintSeqNo;
            sLabPieNo = ""; sLabNo = ""; sOldLabPosition = ""; sNewLabPositon = ""; sPerCaseNo = ""; sOldMRNo = ""; sMRNo = "";
            sPtName = ""; sSex = ""; sBirthday = ""; sAge = ""; sLabType = ""; sLabAdoptDate = ""; sAdoptPortion = ""; sStoreageMethod = "";
            sLabLeaveBodyDatetime = ""; sLabDealDatetime = ""; sLabLeaveBodyEnvir = ""; sLabLeaveBodyHour = "";
            sSubStock = ""; sSickPortion = ""; sDiagName1 = ""; sDiagName2 = ""; sDiagName3 = "";
            sClerkName = ""; sPlanAgree = ""; sPlanAgreeDate = ""; sAgreeNo = ""; sAgreeNoDate = "";
            sUseExpireDate = ""; sChangeRange = ""; sStatus = ""; sNote = ""; sTakeOutName = ""; sTakeOutDate = "";
            sTakeOutApplicant = ""; sTakeOutPlanNo = ""; sTakeOutNote = ""; sInComeDate = ""; sPrintSeqNo = "";

            try
            {
                /*1.Initialize*/
                CleardgvLReqNo2All();

                if (sLReqNo != "")
                {
                    /*2.Fill to Datatable*/
                    //using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                    //insert Event Log: 23. --特殊權限--查詢檢體號碼    --
                    ClsShareFunc.insEvenLogt("23", ClsShareFunc.sUserName, sLReqNo.Substring(0, 9).Trim(), "", "特殊權限(查詢檢體號碼)" + sLReqNo + " --(Cosign Administrator: " + pFunction8_AdminID.Trim());
                    using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                    {
                        sCon.Open();
                        sSQL = "select a.chLabNo,b.chLabPieNo,a.chOldLabPosition,a.chNewLabPositon,b.chPerCaseNo, ";
                        sSQL = sSQL + " b.chOldMRNo,b.chMRNo,b.chPtName,a.chSex,b.chBirthday,a.intAge,";
                        sSQL = sSQL + " a.chLabType,a.chLabAdoptDate,a.chAdoptPortion,a.chStoreageMethod,";
                        sSQL = sSQL + " a.chLabLeaveBodyDatetime,a.chLabDealDatetime,a.chLabLeaveBodyEnvir,a.chLabLeaveBodyHour,";
                        sSQL = sSQL + " a.chSubStock,a.chSickPortion,a.chDiagName1,a.chDiagName2,a.chDiagName3,";
                        sSQL = sSQL + " a.chClerkName,b.chPlanAgree,a.chPlanAgreeDate,b.chAgreeNo,a.chAgreeNoDate,";
                        sSQL = sSQL + " a.chUseExpireDate,a.chChangeRange,a.chStatus,a.chNote,a.chTakeOutName,a.chTakeOutDate,";
                        sSQL = sSQL + " a.chTakeOutApplicant,a.chTakeOutPlanNo,a.chTakeOutNote,a.chInComeDate,a.intPrintSeqNo";
                        sSQL = sSQL + "  FROM [DB_BIO].[dbo].[BioPerMasterTbl] a (nolock) ";
                        sSQL = sSQL + " inner join [DB_SEC].[dbo].[BioPerMappingTbl] b (nolock) ";
                        sSQL = sSQL + " on a.chLabNo = b.chLabNo collate Chinese_Taiwan_Stroke_CI_AS ";
                        sSQL = sSQL + " where a.chLabNo = '" + sLReqNo + "' ";
                        SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                        SqlDataReader sRead = sCmd.ExecuteReader();
                        if (sRead.HasRows)
                        {
                            while (sRead.Read())
                            {
                                sYear = ClsShareFunc.gfunCheck(sRead["chInComeDate"].ToString()).Substring(0, 3);
                                sLabPieNo = dec_AES(ClsShareFunc.gfunCheck(sRead["chLabPieNo"].ToString()), sLReqNo, sYear);
                                sLabNo = ClsShareFunc.gfunCheck(sRead["chLabNo"].ToString());
                                sOldLabPosition = ClsShareFunc.gfunCheck(sRead["chOldLabPosition"].ToString());
                                sNewLabPositon = ClsShareFunc.gfunCheck(sRead["chNewLabPositon"].ToString());
                                sPerCaseNo = dec_AES(ClsShareFunc.gfunCheck(sRead["chPerCaseNo"].ToString()), sLReqNo, sYear);
                                sOldMRNo = dec_AES(ClsShareFunc.gfunCheck(sRead["chOldMRNo"].ToString()), sLReqNo, sYear);
                                sMRNo = dec_AES(ClsShareFunc.gfunCheck(sRead["chMRNo"].ToString()), sLReqNo, sYear);
                                sPtName = dec_AES(ClsShareFunc.gfunCheck(sRead["chPtName"].ToString()), sLReqNo, sYear);
                                sSex = ClsShareFunc.gfunCheck(sRead["chSex"].ToString());
                                sBirthday = dec_AES(ClsShareFunc.gfunCheck(sRead["chBirthday"].ToString()), sLReqNo, sYear);
                                sAge = ClsShareFunc.gfunCheck(sRead["intAge"].ToString());
                                sLabType = ClsShareFunc.gfunCheck(sRead["chLabType"].ToString());
                                sLabAdoptDate = ClsShareFunc.gfunCheck(sRead["chLabAdoptDate"].ToString());
                                sAdoptPortion = ClsShareFunc.gfunCheck(sRead["chAdoptPortion"].ToString());
                                sStoreageMethod = ClsShareFunc.gfunCheck(sRead["chStoreageMethod"].ToString());
                                sLabLeaveBodyDatetime = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyDatetime"].ToString());
                                sLabDealDatetime = ClsShareFunc.gfunCheck(sRead["chLabDealDatetime"].ToString());
                                sLabLeaveBodyEnvir = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyEnvir"].ToString());
                                sLabLeaveBodyHour = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyHour"].ToString());
                                sSubStock = ClsShareFunc.gfunCheck(sRead["chSubStock"].ToString());
                                sSickPortion = ClsShareFunc.gfunCheck(sRead["chSickPortion"].ToString());
                                sDiagName1 = ClsShareFunc.gfunCheck(sRead["chDiagName1"].ToString());
                                sDiagName2 = ClsShareFunc.gfunCheck(sRead["chDiagName2"].ToString());
                                sDiagName3 = ClsShareFunc.gfunCheck(sRead["chDiagName3"].ToString());
                                sClerkName = ClsShareFunc.gfunCheck(sRead["chClerkName"].ToString());
                                sPlanAgree = dec_AES(ClsShareFunc.gfunCheck(sRead["chPlanAgree"].ToString()), sLReqNo, sYear);
                                sPlanAgreeDate = ClsShareFunc.gfunCheck(sRead["chPlanAgreeDate"].ToString());
                                sAgreeNo = dec_AES(ClsShareFunc.gfunCheck(sRead["chAgreeNo"].ToString()), sLReqNo, sYear);
                                sAgreeNoDate = ClsShareFunc.gfunCheck(sRead["chAgreeNoDate"].ToString());
                                sUseExpireDate = ClsShareFunc.gfunCheck(sRead["chUseExpireDate"].ToString());
                                sChangeRange = ClsShareFunc.gfunCheck(sRead["chChangeRange"].ToString());
                                sStatus = ClsShareFunc.gfunCheck(sRead["chStatus"].ToString());
                                sNote = ClsShareFunc.gfunCheck(sRead["chNote"].ToString());
                                sTakeOutName = ClsShareFunc.gfunCheck(sRead["chTakeOutName"].ToString());
                                sTakeOutDate = ClsShareFunc.gfunCheck(sRead["chTakeOutDate"].ToString());
                                sTakeOutApplicant = ClsShareFunc.gfunCheck(sRead["chTakeOutApplicant"].ToString());
                                sTakeOutPlanNo = ClsShareFunc.gfunCheck(sRead["chTakeOutPlanNo"].ToString());
                                sTakeOutNote = ClsShareFunc.gfunCheck(sRead["chTakeOutNote"].ToString());
                                sInComeDate = ClsShareFunc.gfunCheck(sRead["chInComeDate"].ToString());
                                sPrintSeqNo = ClsShareFunc.gfunCheck(sRead["intPrintSeqNo"].ToString());

                                dgvLReqNo2All1.Rows.Add(sLabPieNo, sLabNo, sOldLabPosition, sNewLabPositon,
                                    sPerCaseNo, sOldMRNo, sMRNo, sPtName, sSex, sBirthday);
                                dgvLReqNo2All2.Rows.Add(sAge, sLabType, sLabAdoptDate, sAdoptPortion, sStoreageMethod,
                                    sLabLeaveBodyDatetime, sLabDealDatetime, sLabLeaveBodyEnvir, sLabLeaveBodyHour, sSubStock);
                                dgvLReqNo2All3.Rows.Add(sSickPortion, sDiagName1, sDiagName2, sDiagName3, sClerkName,
                                    sPlanAgree, sPlanAgreeDate, sAgreeNo, sAgreeNoDate, sUseExpireDate, sChangeRange);
                                dgvLReqNo2All4.Rows.Add(sStatus, sNote, sTakeOutName, sTakeOutDate, sTakeOutApplicant,
                                    sTakeOutPlanNo, sTakeOutNote, sInComeDate, sPrintSeqNo);
                            }

                            if (bolClear == true)
                                CleargbID2LReqNo();
                        }
                        else
                            MessageBox.Show("查無此檢體編號!");
                        sRead.Close();
                    }
                }
                else
                    MessageBox.Show("請輸入檢體編號!");

                bolClear = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("txtLReqNo2All_KeyDown: " + ex.Message.ToString());
            }
        }

        private void CleargbID2LReqNo()
        {
            CleardgvID2LReqNo();
            txtID2LReqNo.Text = "";
        }

        private void CleargbLReqNo2All()
        {
            CleardgvLReqNo2All();
            txtLReqNo2All.Text = "";
        }

        /*特殊權限作業 - 清除DataGridView(dgvID2LReqNo)*/
        private void CleardgvID2LReqNo()
        {
            dgvID2LReqNo.Rows.Clear();
        }

        /*特殊權限作業 - 清除DataGridView(dgvLReqNo2All)*/
        private void CleardgvLReqNo2All()
        {
            dgvLReqNo2All1.Rows.Clear();
            dgvLReqNo2All2.Rows.Clear();
            dgvLReqNo2All3.Rows.Clear();
            dgvLReqNo2All4.Rows.Clear();
        }

        /*特殊權限作業 - LReqNo - > Show All */
        private void dgvID2LReqNo_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int sRow = e.RowIndex;
            string sLReqNo = "";
            sLReqNo = dgvID2LReqNo.Rows[sRow].Cells[1].Value.ToString();
            txtLReqNo2All.Text = sLReqNo;
            bolClear = false;
            QryLReqNo2All(sLReqNo);
        }

        /*======================備份======================*/
        /*備份 - 取得本機資訊*/
        private void GetLocalEnv()
        {
            txtRemoteDestination.Text = @"\\192.168.1.3\Batch\";
            string sLocalIP = "";

            // 取得本機名稱
            String sLocalName = Dns.GetHostName();

            // 取得本機的 IpHostEntry 類別實體
            IPHostEntry iphostentry = Dns.GetHostByName(sLocalName);

            // 取得所有 IP 位址
            int num = 1;
            foreach (IPAddress ipaddress in iphostentry.AddressList)
            {
                sLocalIP = ipaddress.ToString();
                num = num + 1;
            }

            //999999
            //if (sLocalIP == "192.168.1.2")
            //{
            //    gbRemote.Enabled = false;
            //    txtRemoteUserName.Text = "";
            //    txtRemoteName.Text = "";
            //    txtRemoteIP.Text = "";
            //    txtLocalSource.Text = "";
            //}
            //else
            //{
            //    gbLocal.Enabled = false;
            //    txtUserName.Text = "";
            //    txtLocalName.Text = "";
            //    txtLocalIP.Text = "";
            //    txtRemoteDestination.Text = "";
            //}

            txtUserName.Text = ClsShareFunc.sUserId;
            txtLocalName.Text = sLocalName;
            txtLocalIP.Text = sLocalIP;
            if (txtLocalName.Text.Trim().Substring(txtLocalName.Text.Trim().Length - 2, 2) == "-1")
            {
                btnBackup.Enabled = true;
                btnRestore.Enabled = false;
            }
            else if (txtLocalName.Text.Trim().Substring(txtLocalName.Text.Trim().Length - 2, 2) == "-2")
            {
                btnBackup.Enabled = false;
                btnRestore.Enabled = true;
            }
            else
            {
                btnBackup.Enabled = false;
                btnRestore.Enabled = false;
            }

        }

        /*備份 - 取得副機資訊*/
        private void GetRemoteEnv()
        {
            txtLocalSource.Text = @"C:\Batch\";
            txtRemoteUserName.Text = ClsShareFunc.sUserId;
            txtRemoteIP.Text = "192.168.1.3";
            //txtRemoteIP.Text = "10.2.5.201";

            Ping pingSender = new Ping();
            PingReply reply = pingSender.Send(txtRemoteIP.Text);
            if (reply.Status != IPStatus.Success)
            {
                MessageBox.Show("請確認副機已開啟!");
                return;
            }

            IPHostEntry hostInfo = Dns.GetHostByAddress(txtRemoteIP.Text);
            txtRemoteName.Text = hostInfo.HostName;
            if (txtRemoteName.Text.Trim().Substring(txtRemoteName.Text.Trim().Length - 2, 2) == "-2")
                btnRestore.Enabled = true;
            else
                btnRestore.Enabled = false;
        }


        private void btnBackup_Click(object sender, EventArgs e)
        {

            //using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
            using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
            {
                try
                {
                    //備份作業
                    //1.如果有舊的ZIP,先delete
                    if (File.Exists(@"C:\Batch\DB_BIO.zip"))
                    {
                        File.Delete(@"C:\Batch\DB_BIO.zip");
                    }
                    if (File.Exists(@"C:\Batch\DB_SEC.zip"))
                    {
                        File.Delete(@"C:\Batch\DB_SEC.zip");
                    }

                    //insert Event Log: 24. --備份 --
                    if (chkSelfBackupOnly.Checked != true)
                        ClsShareFunc.insEvenLogt("24", ClsShareFunc.sUserName, "", "", "備份--");
                    else
                        ClsShareFunc.insEvenLogt("24-1", ClsShareFunc.sUserName, "", "", "備份(本機)--");

                    //2.Backup 成.bak
                    sCon.Open();
                    SqlCommand updateCmd = new SqlCommand(@"BACKUP DATABASE DB_BIO TO DISK='c:\Batch\DB_BIO.bak'", sCon);
                    updateCmd.ExecuteNonQuery();

                    updateCmd = new SqlCommand(@"BACKUP DATABASE DB_SEC TO DISK='c:\Batch\DB_SEC.bak'", sCon);
                    updateCmd.ExecuteNonQuery();

                    //3.加密壓縮成ZIP FILE
                    //壓縮使用參考的BioBank_Conn.dll, .ZIP("Input File.","Output File.")
                    BioBank_Conn.Class_biobank_Compress.ZIP(@"C:\Batch\DB_BIO.bak", @"C:\Batch\DB_BIO.zip");
                    BioBank_Conn.Class_biobank_Compress.ZIP(@"C:\Batch\DB_SEC.bak", @"C:\Batch\DB_SEC.zip");

                    File.Delete(@"C:\Batch\DB_BIO.bak");
                    File.Delete(@"C:\Batch\DB_SEC.bak");
                    //4.如果是備份至副機,就執行以下code:將ZIP File copy 至副機; 後delete 主機的ZIP
                    //如果是備份至主機self,就不執行以下code
                    if (chkSelfBackupOnly.Checked != true)
                    {

                        //1.如果有舊的ZIP,先delete
                        if (File.Exists(@"\\192.168.1.3\Batch\DB_BIO.zip"))
                        {
                            File.Delete(@"\\192.168.1.3\Batch\DB_BIO.zip");
                        }
                        if (File.Exists(@"\\192.168.1.3\Batch\DB_SEC.zip"))
                        {
                            File.Delete(@"\\192.168.1.3\Batch\DB_SEC.zip");
                        }
                        File.Copy(@"C:\Batch\DB_BIO.zip", @"\\192.168.1.3\Batch\DB_BIO.zip");
                        File.Copy(@"C:\Batch\DB_SEC.zip", @"\\192.168.1.3\Batch\DB_SEC.zip");
                        File.Delete(@"C:\Batch\DB_BIO.zip");
                        File.Delete(@"C:\Batch\DB_SEC.zip");
                    }

                    //Process myProcess = new Process();
                    //myProcess.StartInfo.FileName = @"C:\Batch\BackupBio.bat";
                    //myProcess.StartInfo.UseShellExecute = false;
                    //myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    //myProcess.Start();
                    //myProcess.WaitForExit();

                    MessageBox.Show("備份完成!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Backup Error : " + ex.Message.ToString());
                    return;
                }
            }


            // Process myProcess = new Process();
            // myProcess.StartInfo.FileName = @"C:\Batch\BackupBio.bat";
            // myProcess.StartInfo.UseShellExecute = false;
            // myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

            // pbBackup.Minimum = 1;
            // pbBackup.Maximum = 10000;
            // pbBackup.Step = 1;

            ////int sTotalTime = (int)myProcess.TotalProcessorTime.TotalMilliseconds;
            // int i = 0;
            // Boolean bolStatus = true;
            //while (bolStatus == true)
            //{
            //    pbBackup.PerformStep();

            //    if (i == 0)
            //    {
            //        myProcess.Start();
            //        i = 1;
            //    }

            //    if (pbBackup.Value == 10000)
            //    {
            //        bolStatus = false;
            //    }
            //}

            // // Wait for the sort process to write the sorted text lines.
            //  myProcess.WaitForExit();
            //MessageBox.Show("備份完成!");

        }

        /*====================出庫紀錄===================*/
        /*出庫紀錄 - 查詢&分類出庫紀錄 */
        private void QryOutRecord()
        {
            string sSQL = "";
            string sOutTime = "";
            string sTeam = "";
            string sCount = "";

            try
            {
                dgvOutTime.Rows.Clear();

                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                {
                    sCon.Open();
                    sSQL = "SELECT chTakeOutDate,收案小組=substring(chLabNo,1,1),筆數=COUNT(*)";
                    sSQL = sSQL + " FROM [DB_BIO].[dbo].[BioPerMasterTbl] (nolock) ";
                    sSQL = sSQL + " where isnull(chTakeOutDate,'') <> '' ";
                    sSQL = sSQL + " group by chTakeOutDate,substring(chLabNo,1,1) ";
                    SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                    SqlDataReader sRead = sCmd.ExecuteReader();
                    if (sRead.HasRows)
                    {
                        while (sRead.Read())
                        {
                            sOutTime = ClsShareFunc.gfunCheck(sRead["chTakeOutDate"].ToString());
                            sTeam = ClsShareFunc.gfunCheck(sRead["收案小組"].ToString());
                            sCount = ClsShareFunc.gfunCheck(sRead["筆數"].ToString());
                            dgvOutTime.Rows.Add(sOutTime, sTeam, sCount);
                        }

                    } sRead.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("QryOutRecord: " + ex.Message.ToString());
            }
        }

        /*出庫紀錄 - 顯示每筆詳細記錄 */
        private void dgvOutTime_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string sOutTime = "";
            string sSQL = "";
            string sLabNo, sNewLabPositon, sSex, sAge, sLabType, sLabAdoptDate, sAdoptPortion,
                sStoreageMethod, sLabLeaveBodyDatetime, sLabDealDatetime, sLabLeaveBodyEnvir,
                sLabLeaveBodyHour, sSubStock, sSickPortion, sDiagName1, sDiagName2, sDiagName3,
                sClerkName, sPlanAgreeDate, sAgreeNoDate, sUseExpireDate, sChangeRange,
                sStatus, sNote, sTakeOutName, sTakeOutDate, sTakeOutApplicant, sTakeOutPlanNo, sTakeOutNote,
                sInComeDate, sPrintSeqNo;
            sLabNo = ""; sNewLabPositon = ""; sSex = ""; sAge = ""; sLabType = ""; sLabAdoptDate = ""; sAdoptPortion = "";
            sStoreageMethod = ""; sLabLeaveBodyDatetime = ""; sLabDealDatetime = ""; sLabLeaveBodyEnvir = "";
            sLabLeaveBodyHour = ""; sSubStock = ""; sSickPortion = ""; sDiagName1 = ""; sDiagName2 = ""; sDiagName3 = "";
            sClerkName = ""; sPlanAgreeDate = ""; sAgreeNoDate = ""; sUseExpireDate = ""; sChangeRange = "";
            sStatus = ""; sNote = ""; sTakeOutName = ""; sTakeOutDate = ""; sTakeOutApplicant = ""; sTakeOutPlanNo = ""; sTakeOutNote = "";
            sInComeDate = ""; sPrintSeqNo = "";

            sOutTime = dgvOutTime.Rows[e.RowIndex].Cells[0].Value.ToString();
            if (sOutTime == "")
                return;

            //insert Event Log: 16. --出庫記錄查詢  --
            ClsShareFunc.insEvenLogt("16", ClsShareFunc.sUserName, "", "", "出庫記錄查詢--" + sOutTime);

            try
            {
                dgvOutRecord.Rows.Clear();

                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                {
                    sCon.Open();
                    sSQL = "SELECT *  FROM [DB_BIO].[dbo].[BioPerMasterTbl] ";
                    sSQL = sSQL + " where chTakeOutDate = '" + sOutTime + "' ORDER BY chLabNo";
                    SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                    SqlDataReader sRead = sCmd.ExecuteReader();
                    if (sRead.HasRows)
                    {
                        while (sRead.Read())
                        {
                            sLabNo = ClsShareFunc.gfunCheck(sRead["chLabNo"].ToString());
                            //sSex = ClsShareFunc.gfunCheck(sRead["chSex"].ToString());
                            //sAge = ClsShareFunc.gfunCheck(sRead["intAge"].ToString());
                            //sLabType = ClsShareFunc.gfunCheck(sRead["chLabType"].ToString());
                            //sLabAdoptDate = ClsShareFunc.gfunCheck(sRead["chLabAdoptDate"].ToString());
                            //sAdoptPortion = ClsShareFunc.gfunCheck(sRead["chAdoptPortion"].ToString());
                            //sStoreageMethod = ClsShareFunc.gfunCheck(sRead["chStoreageMethod"].ToString());
                            //sLabLeaveBodyDatetime = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyDatetime"].ToString());
                            //sLabDealDatetime = ClsShareFunc.gfunCheck(sRead["chLabDealDatetime"].ToString());
                            //sLabLeaveBodyEnvir = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyEnvir"].ToString());
                            //sLabLeaveBodyHour = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyHour"].ToString());
                            //sSubStock = ClsShareFunc.gfunCheck(sRead["chSubStock"].ToString());
                            //sSickPortion = ClsShareFunc.gfunCheck(sRead["chSickPortion"].ToString());
                            //sDiagName1 = ClsShareFunc.gfunCheck(sRead["chDiagName1"].ToString());
                            //sDiagName2 = ClsShareFunc.gfunCheck(sRead["chDiagName2"].ToString());
                            //sDiagName3 = ClsShareFunc.gfunCheck(sRead["chDiagName3"].ToString());
                            //sClerkName = ClsShareFunc.gfunCheck(sRead["chClerkName"].ToString());
                            //sPlanAgreeDate = ClsShareFunc.gfunCheck(sRead["chPlanAgreeDate"].ToString());
                            //sAgreeNoDate = ClsShareFunc.gfunCheck(sRead["chAgreeNoDate"].ToString());
                            //sUseExpireDate = ClsShareFunc.gfunCheck(sRead["chUseExpireDate"].ToString());
                            //sChangeRange = ClsShareFunc.gfunCheck(sRead["chChangeRange"].ToString());
                            //sStatus = ClsShareFunc.gfunCheck(sRead["chStatus"].ToString());
                            //sNote = ClsShareFunc.gfunCheck(sRead["chNote"].ToString());
                            sTakeOutDate = ClsShareFunc.gfunCheck(sRead["chTakeOutDate"].ToString());
                            sTakeOutName = ClsShareFunc.gfunCheck(sRead["chTakeOutName"].ToString());
                            sTakeOutApplicant = ClsShareFunc.gfunCheck(sRead["chTakeOutApplicant"].ToString());
                            sNewLabPositon = ClsShareFunc.gfunCheck(sRead["chNewLabPositon"].ToString());
                            sTakeOutPlanNo = ClsShareFunc.gfunCheck(sRead["chTakeOutPlanNo"].ToString());
                            //sTakeOutNote = ClsShareFunc.gfunCheck(sRead["chTakeOutNote"].ToString());
                            //sInComeDate = ClsShareFunc.gfunCheck(sRead["chInComeDate"].ToString());
                            //sPrintSeqNo = ClsShareFunc.gfunCheck(sRead["intPrintSeqNo"].ToString());

                            //dgvOutRecord.Rows.Add(sLabNo, sNewLabPositon, sSex, sAge, sLabType, sLabAdoptDate, sAdoptPortion,
                            //     sStoreageMethod, sLabLeaveBodyDatetime, sLabDealDatetime, sLabLeaveBodyEnvir,
                            //     sLabLeaveBodyHour, sSubStock, sSickPortion, sDiagName1, sDiagName2, sDiagName3,
                            //     sClerkName, sPlanAgreeDate, sAgreeNoDate, sUseExpireDate, sChangeRange,
                            //     sStatus, sNote, sTakeOutName, sTakeOutDate, sTakeOutApplicant, sTakeOutPlanNo, sTakeOutNote,
                            //     sInComeDate, sPrintSeqNo);

                            dgvOutRecord.Rows.Add(sTakeOutDate, sTakeOutName, sTakeOutApplicant, sTakeOutPlanNo,
                                sNewLabPositon, sLabNo);
                        }

                    } sRead.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("QryOutRecord: " + ex.Message.ToString());
            }
        }

        private void btnRestore_Click(object sender, EventArgs e)
        {
            //還原作業
            //insert Event Log: 25. --還原--
            if (chkSelfBackupOnly.Checked != true)
                ClsShareFunc.insEvenLogt("25", ClsShareFunc.sUserName, "", "", "還原--");
            using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
            {
                try
                {
                    if (File.Exists(@"C:\Batch\DB_BIO.zip") == false || (File.Exists(@"C:\Batch\DB_SEC.zip") == false))
                    {
                        MessageBox.Show(@"C:\Batch\下的ZIP不存在, 按任一鍵離開.");
                        return;
                    }
                    //1.Copy  .ZIP from C:\Batch to C:\Batch2 
                    //先delete, 以免有殘留, filecopy會當機
                    File.Delete(@"C:\Batch2\DB_BIO.zip");
                    File.Delete(@"C:\Batch2\DB_SEC.zip");
                    File.Delete(@"C:\Batch2\Batch\DB_BIO.bak");
                    File.Delete(@"C:\Batch2\Batch\DB_SEC.bak");

                    File.Copy(@"C:\Batch\DB_BIO.zip", @"C:\Batch2\DB_BIO.zip");
                    File.Copy(@"C:\Batch\DB_SEC.zip", @"C:\Batch2\DB_SEC.zip");

                    //Process myProcess = new Process();
                    //myProcess.StartInfo.FileName = @"C:\Batch2\RestoreBio-1.bat";
                    //myProcess.StartInfo.UseShellExecute = false;
                    //myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    //myProcess.Start();
                    //myProcess.WaitForExit();

                    //2.解壓縮ZIP FILE成 .bak
                    //解壓縮使用參考的BioBank_Conn.dll, ....unZIP("Input File.","Output DIRECTORY.")
                    BioBank_Conn.Class_biobank_Compress.unZIP(@"C:\Batch2\DB_BIO.zip", @"C:\Batch2");
                    BioBank_Conn.Class_biobank_Compress.unZIP(@"C:\Batch2\DB_SEC.zip", @"C:\Batch2");

                    //3.Database Restore
                    sCon.Open();
                    SqlCommand updateCmd = new SqlCommand(@"ALTER DATABASE [DB_BIO] SET  SINGLE_USER WITH ROLLBACK IMMEDIATE", sCon);
                    updateCmd.ExecuteNonQuery();
                    updateCmd = new SqlCommand(@"RESTORE DATABASE DB_BIO FROM DISK='c:\Batch2\Batch\DB_BIO.bak' with Replace", sCon);
                    updateCmd.ExecuteNonQuery();

                    updateCmd = new SqlCommand(@"use master", sCon);
                    updateCmd.ExecuteNonQuery();
                    updateCmd = new SqlCommand(@"ALTER DATABASE [DB_SEC] SET  SINGLE_USER WITH ROLLBACK IMMEDIATE", sCon);
                    updateCmd.ExecuteNonQuery();
                    updateCmd = new SqlCommand(@"RESTORE DATABASE DB_SEC FROM DISK='c:\Batch2\Batch\DB_SEC.bak' with Replace", sCon);
                    updateCmd.ExecuteNonQuery();

                    //4.delete .bak 以防止,資安問題  
                    File.Delete(@"C:\Batch2\DB_BIO.zip");
                    File.Delete(@"C:\Batch2\DB_SEC.zip");
                    File.Delete(@"C:\Batch\DB_BIO.zip");
                    File.Delete(@"C:\Batch\DB_SEC.zip");
                    File.Delete(@"C:\Batch2\Batch\DB_BIO.bak");
                    File.Delete(@"C:\Batch2\Batch\DB_SEC.bak");


                    MessageBox.Show("還原完成!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Restore Error : " + ex.Message.ToString());
                    return;
                }
            }
            //Process myProcess = new Process();
            //myProcess.StartInfo.FileName = @"C:\Batch2\RestoreBio.bat";
            //myProcess.StartInfo.UseShellExecute = false;
            //myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            //// Wait for the sort process to write the sorted text lines.
            //myProcess.Start();
            //myProcess.WaitForExit();
            //MessageBox.Show("還原完成!");
        }

        private void textBoxFilePath_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtID2LReqNo_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtLReqNo2All_TextChanged(object sender, EventArgs e)
        {

        }

        private void tpBackup_Click(object sender, EventArgs e)
        {

        }

        private void txtUserName_TextChanged(object sender, EventArgs e)
        {

        }

        private void TEST_Click(object sender, EventArgs e)
        {
            //1.如果有舊的ZIP,先delete
            MessageBox.Show("1111");
            if (File.Exists(@"\\192.168.1.3\Batch\DB_BIO.zip"))
            {
                File.Delete(@"\\192.168.1.3\Batch\DB_BIO.zip");
            }
            MessageBox.Show("22222");
            if (File.Exists(@"\\192.168.1.3\Batch\DB_SEC.zip"))
            {
                File.Delete(@"\\192.168.1.3\Batch\DB_SEC.zip");
            }
            MessageBox.Show("333333");
            File.Copy(@"C:\Batch\DB_BIO.zip", @"\\192.168.1.3\Batch\DB_BIO.zip");
            MessageBox.Show("4444");
            File.Copy(@"C:\Batch\DB_SEC.zip", @"\\192.168.1.3\Batch\DB_SEC.zip");
        }

        private void txtID_Admin_TextChanged(object sender, EventArgs e)
        {

        }

        private void dgvOutLReqNo_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void txtPwd_TextChanged(object sender, EventArgs e)
        {

        }

        private void dgvStorageTime_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgvOutTime_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Refresh_Click(object sender, EventArgs e)
        {

            dgvEventLog.Rows.Clear();

            try
            {
                string sSQL = "";
                string sEventDateTime = "";
                string sEventNo = "";
                string sCLerkName = "";
                string sLabNo = "";
                string sMRNo = "";
                string sOtherValue = "";

                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_SEC_conn())
                {
                    sCon.Open();
                    sSQL = "select *  from BioEventLogTbl (nolock) where chEventDateTime like '" + txtEventDate.Text.ToString().Trim() + "%' and chEventNo like '" + txtEventNo.Text.Trim() + "%' order by chEventDateTime desc";

                    SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                    SqlDataReader sRead = sCmd.ExecuteReader();
                    while (sRead.Read())
                    {
                        sEventDateTime = ClsShareFunc.gfunCheck(sRead["chEventDateTime"].ToString().Substring(0, 7)) + "-" +
                                ClsShareFunc.gfunCheck(sRead["chEventDateTime"].ToString().Substring(7, 2)) + ":" + ClsShareFunc.gfunCheck(sRead["chEventDateTime"].ToString().Substring(9, 2)) + ":" + ClsShareFunc.gfunCheck(sRead["chEventDateTime"].ToString().Substring(11));
                        sEventNo = ClsShareFunc.gfunCheck(sRead["chEventNo"].ToString());
                        sCLerkName = ClsShareFunc.gfunCheck(sRead["chClerkName"].ToString());
                        sLabNo = ClsShareFunc.gfunCheck(sRead["chLabNo"].ToString());
                        //1050105:花蓮大林主任共同同意,病歷號不show出
                        //sMRNo = ClsShareFunc.gfunCheck(sRead["chMRNo"].ToString());
                        sMRNo = (ClsShareFunc.gfunCheck(sRead["chMRNo"].ToString()) == "" ? "" : "*");
                        sOtherValue = ClsShareFunc.gfunCheck(sRead["chOtherValue"].ToString());

                        dgvEventLog.Rows.Add(sEventDateTime, sEventNo, sCLerkName, sLabNo, sMRNo, sOtherValue);
                    }
                    sRead.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("QryEventLog: " + ex.Message.ToString());
            }

        }

        private void butClear_Click(object sender, EventArgs e)
        {
            //dgvShowExcel.Rows.Clear();
            //dgvShowMsg.Rows.Clear();

            //if (dgvShowExcel.DataSource != null)
            //{
            dgvShowExcel.DataSource = null;
            dgvShowExcel.DataSource = null;
            dgvShowExcel.Rows.Clear();     //clear row item 
            //dgvShowExcel.Columns.Clear();  //clear column  item     
            //}
            //if (dgvShowMsg.DataSource != null)
            //{
            dgvShowMsg.DataSource = null;
            dgvShowMsg.DataSource = null;
            dgvShowMsg.Rows.Clear();     //clear row item 
            //dgvShowMsg.Columns.Clear();  //clear column  item     
            //}
        }

        private void dgvQryLReqNo_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tpSpAuth_Click(object sender, EventArgs e)
        {

        }

        private void BioBank_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        public void cab_Button_Click(object sender, EventArgs e)
        {
            var PD = new PrintDocument();

            if (printFunction.IsPrinterExist("CAB MACH4/300"))
            {
                PD.PrinterSettings.PrinterName = "CAB MACH4/300";
                try
                {
                    for (int i = 0; i < dgvStorageRecord.Rows.Count; i++)
                    //for (int i = 0; i < 1; i++)
                    {
                        printNum = dgvStorageRecord.Rows[i].Cells["檢體管號碼"].Value.ToString();
                        PD.PrintPage += new PrintPageEventHandler(PD_PrintPage);
                        PD.Print();
                    }
                    //printNum = "U121516777";
                    //PD.PrintPage += new PrintPageEventHandler(PD_PrintPage);
                    //PD.Print();
                }
                catch (Exception ex){
                    MessageBox.Show(ex.Message);
                }
            }
            
        }
        public void PD_PrintPage(object sender, PrintPageEventArgs e)
        {
            e.Graphics.DrawString("*" + printNum + "*", new Font("Free 3 of 9 Extended", 14, FontStyle.Regular), Brushes.Black, 5, 0);
            e.Graphics.DrawString("*" + printNum + "*", new Font("Free 3 of 9 Extended", 14, FontStyle.Regular), Brushes.Black, 5, 12);
            e.Graphics.DrawString(printNum, new Font("新細明體", 8, FontStyle.Regular), Brushes.Black, 25, 26);
        }

        private void btnSearchOut_Click(object sender, EventArgs e)
        {
            tabForm.SelectedIndex = 4;
            dgvOutLReqNo.Rows.Clear();
            ArrayList chkNumAL = new ArrayList();
            string sSQL = "";

            for (int i = 0; i < dgvSearchData.Rows.Count; i++)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)dgvSearchData.Rows[i].Cells[1];
                
                //if (chk.Value.ToString() == "True")
                //{
                //    for (int j = 1; j < 8; j++)
                //    {
                //        aStr[j] = dgvSearchData.Rows[i].Cells[j].Value.ToString();
                //    }
                //    dgvOutLReqNo.Rows.Add(aStr);
                //}
                //檢體代碼暫存
                if (chk.Value.ToString() == "True")
                {
                    chkNumAL.Add(dgvSearchData.Rows[i].Cells[2].Value.ToString());
                }
            }

            try
            {
                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                {
                    sCon.Open();

                    for (int i = 0; i <= chkNumAL.Count-1; i++)
                    {
                        string[] aStr = new string[32];
                        sSQL = "SELECT * from BioPerMasterTbl WHERE chLabNo = '" + chkNumAL[i].ToString().Trim() + "'";
                        SqlCommand sCmd = new SqlCommand(sSQL, sCon);
                        SqlDataReader sRead = sCmd.ExecuteReader();
                        while (sRead.Read())
                        {
                            aStr[1] = ClsShareFunc.gfunCheck(sRead["chLabNo"].ToString());
                            aStr[2] = ClsShareFunc.gfunCheck(sRead["chNewLabPositon"].ToString());
                            aStr[3] = ClsShareFunc.gfunCheck(sRead["chSex"].ToString());
                            aStr[4] = ClsShareFunc.gfunCheck(sRead["intAge"].ToString());
                            aStr[5] = ClsShareFunc.gfunCheck(sRead["chLabType"].ToString());
                            aStr[6] = ClsShareFunc.gfunCheck(sRead["chLabAdoptDate"].ToString());
                            aStr[7] = ClsShareFunc.gfunCheck(sRead["chAdoptPortion"].ToString());
                            aStr[8] = ClsShareFunc.gfunCheck(sRead["chStoreageMethod"].ToString());
                            aStr[9] = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyDatetime"].ToString());
                            aStr[10] = ClsShareFunc.gfunCheck(sRead["chLabDealDatetime"].ToString());
                            aStr[11] = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyEnvir"].ToString());
                            aStr[12] = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyHour"].ToString());
                            aStr[13] = ClsShareFunc.gfunCheck(sRead["chSubStock"].ToString());
                            aStr[14] = ClsShareFunc.gfunCheck(sRead["chSickPortion"].ToString());
                            aStr[15] = ClsShareFunc.gfunCheck(sRead["chDiagName1"].ToString());
                            aStr[16] = ClsShareFunc.gfunCheck(sRead["chDiagName2"].ToString());
                            aStr[17] = ClsShareFunc.gfunCheck(sRead["chDiagName3"].ToString());
                            aStr[18] = ClsShareFunc.gfunCheck(sRead["chClerkName"].ToString());
                            aStr[19] = ClsShareFunc.gfunCheck(sRead["chPlanAgreeDate"].ToString());
                            aStr[20] = ClsShareFunc.gfunCheck(sRead["chAgreeNoDate"].ToString());
                            aStr[21] = ClsShareFunc.gfunCheck(sRead["chUseExpireDate"].ToString());
                            aStr[22] = ClsShareFunc.gfunCheck(sRead["chChangeRange"].ToString());
                            aStr[23] = ClsShareFunc.gfunCheck(sRead["chStatus"].ToString());
                            aStr[24] = ClsShareFunc.gfunCheck(sRead["chNote"].ToString());
                            aStr[25] = ClsShareFunc.gfunCheck(sRead["chTakeOutName"].ToString());
                            aStr[26] = ClsShareFunc.gfunCheck(sRead["chTakeOutDate"].ToString());
                            aStr[27] = ClsShareFunc.gfunCheck(sRead["chTakeOutApplicant"].ToString());
                            aStr[28] = ClsShareFunc.gfunCheck(sRead["chTakeOutPlanNo"].ToString());
                            aStr[29] = ClsShareFunc.gfunCheck(sRead["chTakeOutNote"].ToString());
                            aStr[30] = ClsShareFunc.gfunCheck(sRead["chInComeDate"].ToString());
                            aStr[31] = ClsShareFunc.gfunCheck(sRead["intPrintSeqNo"].ToString());
                        }
                        //把值傳到代入出庫表
                        dgvOutLReqNo.Rows.Add(aStr);
                    }
                    sCon.Close();
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Source);
            }
        }
        /* GridView 匯出Execel 事件 ---------- Start */
        private void btnOutExcel_Click(object sender, EventArgs e)
        {
            frmSaveFiles frmExlExport = new frmSaveFiles();
            frmExlExport.Show();
            ClsShareFunc.nowDGV = dgvSearchData;
        }

        private void btnOPExcel_Click(object sender, EventArgs e)
        {
            frmSaveFiles frmExlExport = new frmSaveFiles();
            frmExlExport.Show();
            ClsShareFunc.nowDGV = dgvOutRecord;
        }
        /* GridView 匯出Execel 事件 ---------- End */

        /* GridView 右鍵複製功能 ---------- Start */
        private void dgvStorageRecord_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                menu.Show(dgvStorageRecord, new Point(e.X, e.Y));//顯示右鍵選單
                ClsShareFunc.nowDGV = dgvStorageRecord;
            }
        }

        private void dgvSearchData_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                menu.Show(dgvSearchData, new Point(e.X, e.Y));//顯示右鍵選單
                ClsShareFunc.nowDGV = dgvSearchData;
            }
        }

        private void contexMenuuu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            ToolStripItem item = e.ClickedItem;
            switch (item.Text)
            {
                case "複製":
                    Clipboard.SetDataObject(ClsShareFunc.nowDGV.CurrentCell.Value.ToString(), false);
                    break;
            }
        }
        /* GridView 右鍵複製功能 ---------- End */

        private void btnPrintChkOut_Click(object sender, EventArgs e)
        {
            //insert Event Log: 12. --篩選(列印)--
            ClsShareFunc.insEvenLogt("12", ClsShareFunc.sUserName, "", "", "出庫紀錄(列印)--");
            ClsPrint _ClsPrint = new ClsPrint(dgvOutRecord, "查詢列印");
            _ClsPrint.PrintForm();
        }

        //輸入檢體編碼時按enter會做的事情
        private void txtModLReqNo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(Convert.ToInt32(e.KeyChar) == 13){
                txtModLReqNo.SelectAll();
            }
        }

        private void dgvQryLReqNo_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvQryLReqNo.IsCurrentCellDirty)
            {
                dgvQryLReqNo.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dgvQryLReqNo_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string sLReqNo = "";
                string sYear = "";
                string sRange = "";
                string sStatus = "";
                string sNote = "";

                CleardgvShowLReqNo();
                if (e.ColumnIndex != 0 && dgvQryLReqNo.RowCount > 0)
                {
                    sLReqNo = dgvQryLReqNo.Rows[0].Cells[1].Value.ToString();
                    sYear = dgvQryLReqNo.Rows[0].Cells[21].Value.ToString();
                    sRange = dgvQryLReqNo.Rows[0].Cells[22].Value.ToString();
                    sStatus = dgvQryLReqNo.Rows[0].Cells[23].Value.ToString().Trim();
                    sNote = dgvQryLReqNo.Rows[0].Cells[24].Value.ToString();

                    dgvShowLReqNo.Rows.Add(sLReqNo, sYear, sRange, sStatus, sNote);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("dgvQryLReqNo_RowHeaderMouseDoubleClick: " + ex.Message.ToString());
            }
        }
        //印列單筆貼紙
        private void singPrtBtn_Click(object sender, EventArgs e)
        {
            var PD = new PrintDocument();

            if (printFunction.IsPrinterExist("CAB MACH4/300"))
            {
                PD.PrinterSettings.PrinterName = "CAB MACH4/300";
                try
                {
                    int rowNo = 0;
                    rowNo = dgvStorageRecord.CurrentCell.RowIndex;
                    for (int i = 0; i < dgvStorageRecord.Rows.Count; i++)
                    printNum = dgvStorageRecord.Rows[rowNo].Cells["檢體管號碼"].Value.ToString();
                    PD.PrintPage += new PrintPageEventHandler(PD_PrintPage);
                    PD.Print();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void btnSerarchAll_Click(object sender, EventArgs e)
        {
            dgvSearchData.Rows.Clear();
            string strAll = txtSearchAll.Text.ToString();
            string strSQL = "select * from BioPerMasterTbl where "
                + "chLabNo like '%" + strAll + "%'"
                + " or chOldLabPosition like '%" + strAll + "%'"
                + " or chNewLabPositon like '%" + strAll + "%'"
                + " or chSex like '%" + strAll + "%'"
                + " or intAge like '%" + strAll + "%'"
                + " or chLabType like '%" + strAll + "%'"
                + " or chLabAdoptDate like '%" + strAll + "%'"
                + " or chStoreageMethod like '%" + strAll + "%'"
                + " or chLabLeaveBodyDatetime like '%" + strAll + "%'"
                + " or chLabDealDatetime like '%" + strAll + "%'"
                + " or chLabLeaveBodyEnvir like '%" + strAll + "%'"
                + " or chLabLeaveBodyHour like '%" + strAll + "%'"
                + " or chSubStock like '%" + strAll + "%'"
                + " or chSickPortion like '%" + strAll + "%'"
                + " or chDiagName1 like '%" + strAll + "%'"
                + " or chDiagName2 like '%" + strAll + "%'"
                + " or chDiagName3 like '%" + strAll + "%'"
                + " or chClerkName like '%" + strAll + "%'"
                + " or chPlanAgreeDate like '%" + strAll + "%'"
                + " or chAgreeNoDate like '%" + strAll + "%'"
                + " or chUseExpireDate like '%" + strAll + "%'"
                + " or chChangeRange like '%" + strAll + "%'"
                + " or chStatus like '%" + strAll + "%'"
                + " or chNote like '%" + strAll + "%'"
                + " or chTakeOutName like '%" + strAll + "%'"
                + " or chTakeOutDate like '%" + strAll + "%'"
                + " or chTakeOutApplicant like '%" + strAll + "%'"
                + " or chTakeOutPlanNo like '%" + strAll + "%'"
                + " or chTakeOutNote like '%" + strAll + "%'"
                + " or chInComeDate like '%" + strAll + "%'";
        
        //insert Event Log: 11. --篩選(查詢)--
            ClsShareFunc.insEvenLogt("11", ClsShareFunc.sUserName, "", "", "篩選(查詢)--");
            using (SqlConnection conn = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
            {
                conn.Open();

                string sLabNo, sNewLabPositon, sSex, sAge, sLabType, sLabAdoptDate, sAdoptPortion,
                    sStoreageMethod, sLabLeaveBodyDatetime, sLabDealDatetime, sLabLeaveBodyEnvir,
                    sLabLeaveBodyHour, sSubStock, sSickPortion, sDiagName1, sDiagName2, sDiagName3,
                    sClerkName, sPlanAgreeDate, sAgreeNoDate, sUseExpireDate, sChangeRange,
                    sStatus, sNote, sTakeOutName, sTakeOutDate, sTakeOutApplicant, sTakeOutPlanNo, sTakeOutNote,
                    sInComeDate, sPrintSeqNo;
                sLabNo = ""; sNewLabPositon = ""; sSex = ""; sAge = ""; sLabType = ""; sLabAdoptDate = ""; sAdoptPortion = "";
                sStoreageMethod = ""; sLabLeaveBodyDatetime = ""; sLabDealDatetime = ""; sLabLeaveBodyEnvir = "";
                sLabLeaveBodyHour = ""; sSubStock = ""; sSickPortion = ""; sDiagName1 = ""; sDiagName2 = ""; sDiagName3 = "";
                sClerkName = ""; sPlanAgreeDate = ""; sAgreeNoDate = ""; sUseExpireDate = ""; sChangeRange = "";
                sStatus = ""; sNote = ""; sTakeOutName = ""; sTakeOutDate = ""; sTakeOutApplicant = ""; sTakeOutPlanNo = ""; sTakeOutNote = "";
                sInComeDate = ""; sPrintSeqNo = "";

                try
                {
                    using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                    {
                        sCon.Open();

                        SqlCommand sCmd = new SqlCommand(strSQL, sCon);
                        SqlDataReader sRead = sCmd.ExecuteReader();
                        int ctRowNum = 1;
                        btnSearchOut.Visible = false;

                        while (sRead.Read())
                        {
                            sLabNo = ClsShareFunc.gfunCheck(sRead["chLabNo"].ToString());
                            sNewLabPositon = ClsShareFunc.gfunCheck(sRead["chNewLabPositon"].ToString());
                            sSex = ClsShareFunc.gfunCheck(sRead["chSex"].ToString());
                            sAge = ClsShareFunc.gfunCheck(sRead["intAge"].ToString());
                            sLabType = ClsShareFunc.gfunCheck(sRead["chLabType"].ToString());
                            sLabAdoptDate = ClsShareFunc.gfunCheck(sRead["chLabAdoptDate"].ToString());
                            sAdoptPortion = ClsShareFunc.gfunCheck(sRead["chAdoptPortion"].ToString());
                            sStoreageMethod = ClsShareFunc.gfunCheck(sRead["chStoreageMethod"].ToString());
                            sLabLeaveBodyDatetime = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyDatetime"].ToString());
                            sLabDealDatetime = ClsShareFunc.gfunCheck(sRead["chLabDealDatetime"].ToString());
                            sLabLeaveBodyEnvir = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyEnvir"].ToString());
                            sLabLeaveBodyHour = ClsShareFunc.gfunCheck(sRead["chLabLeaveBodyHour"].ToString());
                            sSubStock = ClsShareFunc.gfunCheck(sRead["chSubStock"].ToString());
                            sSickPortion = ClsShareFunc.gfunCheck(sRead["chSickPortion"].ToString());
                            sDiagName1 = ClsShareFunc.gfunCheck(sRead["chDiagName1"].ToString());
                            sDiagName2 = ClsShareFunc.gfunCheck(sRead["chDiagName2"].ToString());
                            sDiagName3 = ClsShareFunc.gfunCheck(sRead["chDiagName3"].ToString());
                            sClerkName = ClsShareFunc.gfunCheck(sRead["chClerkName"].ToString());
                            sPlanAgreeDate = ClsShareFunc.gfunCheck(sRead["chPlanAgreeDate"].ToString());
                            sAgreeNoDate = ClsShareFunc.gfunCheck(sRead["chAgreeNoDate"].ToString());
                            sUseExpireDate = ClsShareFunc.gfunCheck(sRead["chUseExpireDate"].ToString());
                            sChangeRange = ClsShareFunc.gfunCheck(sRead["chChangeRange"].ToString());
                            sStatus = ClsShareFunc.gfunCheck(sRead["chStatus"].ToString());
                            sNote = ClsShareFunc.gfunCheck(sRead["chNote"].ToString());
                            sTakeOutName = ClsShareFunc.gfunCheck(sRead["chTakeOutName"].ToString());
                            sTakeOutDate = ClsShareFunc.gfunCheck(sRead["chTakeOutDate"].ToString());
                            sTakeOutApplicant = ClsShareFunc.gfunCheck(sRead["chTakeOutApplicant"].ToString());
                            sTakeOutPlanNo = ClsShareFunc.gfunCheck(sRead["chTakeOutPlanNo"].ToString());
                            sTakeOutNote = ClsShareFunc.gfunCheck(sRead["chTakeOutNote"].ToString());
                            sInComeDate = ClsShareFunc.gfunCheck(sRead["chInComeDate"].ToString());
                            sPrintSeqNo = ClsShareFunc.gfunCheck(sRead["intPrintSeqNo"].ToString());

                            //dgvSearchData.Rows.Add(false, sLabNo, sNewLabPositon, sSex, sAge, sLabType, sLabAdoptDate, sAdoptPortion,
                            //    sStoreageMethod, sLabLeaveBodyDatetime, sLabDealDatetime, sLabLeaveBodyEnvir,
                            //    sLabLeaveBodyHour, sSubStock, sSickPortion, sDiagName1, sDiagName2, sDiagName3,
                            //    sClerkName, sPlanAgreeDate, sAgreeNoDate, sUseExpireDate, sChangeRange,
                            //    sStatus, sNote,sTakeOutName, sTakeOutDate, sTakeOutApplicant, sTakeOutPlanNo, sTakeOutNote,
                            //    sInComeDate, sPrintSeqNo);
                            dgvSearchData.Rows.Add(ctRowNum, false, sLabNo, sSubStock, sLabType, sStoreageMethod, sSickPortion, sDiagName1, sDiagName2, sDiagName3);

                            //有選未出庫的情況
                            if (chkGetOut.Text != "")
                            {
                                string takeStr = chkGetOut.Text.ToString();
                                if (takeStr == "出庫")
                                {
                                    //MessageBox.Show(dgvSearchData.Rows[ctRowNum-1].Cells[1].Value.ToString());
                                    dgvSearchData.Rows[ctRowNum - 1].Cells[1].ReadOnly = true;
                                    dgvSearchData.Rows[ctRowNum - 1].Cells[1].Value = true;
                                }
                                else{
                                    btnSearchOut.Visible = true;
                                }
                            }
                            //沒有選出庫未出庫的情況
                            else
                            {
                                if (sTakeOutDate != "")
                                {
                                    dgvSearchData.Rows[ctRowNum - 1].Cells[1].ReadOnly = true;
                                    dgvSearchData.Rows[ctRowNum - 1].Cells[1].Value = true;
                                }
                            }

                            ctRowNum++;
                        }
                        sRead.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("QryLReqNo: " + ex.Message.ToString());
                }
            }
        }

        //預設篩選預設欄位
        private void searchCharge()
        {
            string strSQL = "SELECT DISTINCT chDiagName1, chDiagName2, chDiagName3 from BioPerMasterTbl";
            AutoCompleteStringCollection acc1 = new AutoCompleteStringCollection();
            AutoCompleteStringCollection acc2 = new AutoCompleteStringCollection();
            AutoCompleteStringCollection acc3 = new AutoCompleteStringCollection();
            string temp1 = "";
            string temp2 = "";
            string temp3 = "";

            textBoxDiag1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            textBoxDiag2.AutoCompleteSource = AutoCompleteSource.CustomSource;
            textBoxDiag3.AutoCompleteSource = AutoCompleteSource.CustomSource;

            try
            {
                using (SqlConnection sCon = BioBank_Conn.Class_biobank_conn.DB_BIO_conn())
                {
                    sCon.Open();

                    SqlCommand sCmd = new SqlCommand(strSQL, sCon);
                    SqlDataReader sRead = sCmd.ExecuteReader();
                    while (sRead.Read())
                    {
                        temp1 = ClsShareFunc.gfunCheck(sRead["chDiagName1"].ToString()).Trim();
                        temp2 = ClsShareFunc.gfunCheck(sRead["chDiagName2"].ToString()).Trim();
                        temp3 = ClsShareFunc.gfunCheck(sRead["chDiagName3"].ToString()).Trim();

                        if (temp1 != "")
                            acc1.Add(temp1);
                        if (temp2 != "")
                            acc2.Add(temp2);
                        if (temp3 != "")
                            acc3.Add(temp3);
                    }
                }
                textBoxDiag1.AutoCompleteCustomSource = acc1;
                textBoxDiag2.AutoCompleteCustomSource = acc2;
                textBoxDiag3.AutoCompleteCustomSource = acc3;
            }
            catch (Exception e)
            {
                MessageBox.Show("diagnosis error: " + e.Message.ToString());
            }
        }

        private void txtSDate_MouseClick(object sender, MouseEventArgs e)
        {
            txtSDate.SelectAll();
        }

        private void txtEDate_MouseClick(object sender, MouseEventArgs e)
        {
            txtEDate.SelectAll();
        }

        private void txtSDate_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar == Convert.ToChar(13))
            {
                txtEDate.Focus();
            }
        }

    }
}
