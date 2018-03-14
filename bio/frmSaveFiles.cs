using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BioBank
{
    public partial class frmSaveFiles : Form
    {
        public frmSaveFiles()
        {
            InitializeComponent();
            txtPath.Tag = "請選擇路徑";
            txtPath.Text = (string)txtPath.Tag;
            txtFileName.Tag = "請輸入檔名";
            txtFileName.Text = (string)txtFileName.Tag;
            
        }

        private void frmSaveFiles_Load(object sender, EventArgs e)
        {
        }

        //資料夾目錄
        private void btnPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtPath.Text = dlg.SelectedPath;
            }
        }
        //檔案名稱，匯出EXCEL
        private void btnExlExp_Click(object sender, EventArgs e)
        {
            string fileName = txtFileName.Text;
            if (txtPath.Text != "" && txtFileName.Text != "")
            {
                ClsShareFunc.OutPutExcel(ClsShareFunc.nowDGV, txtPath.Text, fileName);
                this.Close();
            }
            else
                MessageBox.Show("請選擇路徑，檔名不可為空白!");
        }

        private void txtFileName_Click(object sender, EventArgs e)
        {
            this.txtFileName.SelectAll();
        }

        private void txtPath_Click(object sender, EventArgs e)
        {
            this.txtPath.SelectAll();
        }
    }
}
