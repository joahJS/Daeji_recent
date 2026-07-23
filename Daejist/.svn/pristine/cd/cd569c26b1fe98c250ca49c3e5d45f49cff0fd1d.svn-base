using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using ComLib;
using System.IO;

namespace AccAdm
{
    public partial class PopSelDoc : DevExpress.XtraEditors.XtraForm
    {
        public PopSelDoc()
        {
            InitializeComponent();
        }

        //FTP 서버
        private string sInitDir = ComnEtcFunc.FTP_ROOT + @"/ERP/AprlDocs/";
        private string user = ComnEtcFunc.FTP_USER;
        private string pw = ComnEtcFunc.FTP_PW;

        private void PopSelDoc_Load(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DataTable dtDoctp = GetLookupData("0");
            ComGrid.SetLookUpEdit(LkupDoctp, dtDoctp, "CD", "NM", "");
        }

        #region LOOKUP DATA
        private DataTable GetLookupData(string sGB)
        {
            StringBuilder strSql = new StringBuilder();

            if (sGB.Equals("0"))
            {
                strSql.Clear();
                strSql.AppendLine(" SELECT COM_CD AS CD   ");
                strSql.AppendLine("      , COM_NM AS NM   ");
                strSql.AppendLine("   FROM COM_BASE_CD    ");
                strSql.AppendLine("  WHERE CD_GB = 'DOCTP'");
            }

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }
        #endregion

        private void BtnDownload_Click(object sender, EventArgs e)
        {
            string sComcd = LkupDoctp.EditValue?.ToString();

            if (string.IsNullOrEmpty(sComcd))
            {
                XtraMessageBox.Show("다운로드할 양식을 선택해주세요.");
                return;
            }

            using (XtraFolderBrowserDialog folderBrowserDialog = new XtraFolderBrowserDialog())
            {
                folderBrowserDialog.Title = "파일을 저장할 폴더를 선택해주세요.";

                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        Cursor = Cursors.WaitCursor;

                        StringBuilder strSql = new StringBuilder();

                        strSql.AppendLine(" SELECT COM_SUB_NM1 AS FILNM   ");
                        strSql.AppendLine("   FROM COM_BASE_CD    ");
                        strSql.AppendLine("  WHERE CD_GB = 'DOCTP'");
                        strSql.AppendLine("    AND COM_CD = '"+ sComcd + "'   ");

                        DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                        if(dt == null || dt.Rows.Count <= 0)
                        {
                            XtraMessageBox.Show("양식이 없습니다.");
                            return;
                        }

                        string sFileName = dt.Rows[0]["FILNM"]?.ToString();
                        string downPath = folderBrowserDialog.SelectedPath + @"/" + sFileName;

                        if(ComnEtcFunc.FTPFileDownload(sInitDir + sFileName, downPath, user, pw))
                        {
                            XtraMessageBox.Show("다운로드를 완료하였습니다.");
                        }
                        else
                        {
                            XtraMessageBox.Show("다운로드 오류입니다.\r\n관리자에게 문의하세요.");
                        }

                        Cursor = Cursors.Default;
                    }
                    catch (Exception ex)
                    {
                        Cursor = Cursors.Default;
                        XtraMessageBox.Show(ex.ToString());
                        return;
                    }
                }
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void PopSelDoc_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                BtnClose.PerformClick();
        }

        private void LkupDoctp_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnDownload.Focus();
        }
    }
}