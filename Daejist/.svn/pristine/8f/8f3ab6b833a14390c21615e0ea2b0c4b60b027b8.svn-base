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
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using System.Net;
using System.IO;

namespace AccAdm
{
    public partial class PD01001F04 : DevExpress.XtraEditors.XtraForm
    {
        public PD01001F04()
        {
            InitializeComponent();
        }

        public string MAKENO;
        private void PD01001F04_Load(object sender, EventArgs e)
        {
            GridRetr.DataSource = GetInspectInfo(MAKENO);
        }
        
        private void BtnDel_Click(object sender, EventArgs e)
        {
            try
            {
                int[] iArrSelectedIdx = GridViewRetr.GetSelectedRows();
                if(iArrSelectedIdx.Length == 0)
                {
                    XtraMessageBox.Show("삭제하려는 데이터에 체크를 진행하세요.");
                    return;
                }
                
                if (XtraMessageBox.Show(string.Format("선택하신 {0}건에 대하여 검수내역을 삭제하시겠습니까?", iArrSelectedIdx.Length), "삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }

                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < iArrSelectedIdx.Length; i++)
                {
                    DataRow row = GridViewRetr.GetDataRow(iArrSelectedIdx[i]);

                    string sMakeNo = row["MAKENO"]?.ToString();
                    string sMakeNo_Lm = row["MAKENO_LM"]?.ToString();

                    #region 항목삭제시 이미지도 같이 삭제(2021-12-21)
                    //strSql.Clear();
                    //strSql.AppendLine(" SELECT MDATE ");
                    //strSql.AppendLine("   FROM MAKE_M A ");
                    //strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");
                    //strSql.AppendLine("  ORDER BY MDATE DESC ");

                    //DataTable dtM = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    //string sMdate = string.Empty;

                    //if (dtM.Rows.Count > 0)
                    //{
                    //    sMdate = dtM.Rows[0]["MDATE"].ToString();
                    //}

                    //if (!string.IsNullOrEmpty(sMdate))
                    //{
                    //    string sInitDir = string.Format(@"ftp://{0}/Gumsu_Images/{1}/{2}/{3}/{4}_{5}", ComnEtcFunc.FTP_IP, sMdate.Substring(0, 4), sMdate.Substring(4, 2), sMdate.Substring(6, 2), sMakeNo, sMakeNo_Lm);

                    //    //FTP 서버
                    //    string user = ComnEtcFunc.FTP_USER;
                    //    string pw = ComnEtcFunc.FTP_PW;

                    //    string[] fileList = GetFtpFileList(sInitDir, user, pw);

                    //    if(fileList != null)
                    //    {
                    //        for (int j = 0; j < fileList.Length; j++)
                    //        {
                    //            string delFilePath = sInitDir + @"/" + fileList[j];
                    //            DeleteFTPFile(delFilePath, user, pw);
                    //        }
                    //    }
                    //}
                    #endregion

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" DELETE FROM MAKE_4 ");
                    strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");
                    strSql.AppendLine("    AND MAKENO_LM = " + sMakeNo_Lm + " ");
                    
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                }
                
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                Cursor = Cursors.Default;
                XtraMessageBox.Show("삭제 완료했습니다.");

                DialogResult = DialogResult.OK;
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        #region ftp관련
        /// <summary>
        /// FTP 파일 삭제하기
        /// </summary>
        /// <param name="userID">사용자 ID</param>
        /// <param name="password">패스워드</param>
        /// <param name="targetURI">타겟 URI</param>
        /// <returns>처리 결과</returns>
        private bool DeleteFTPFile(string targetURI, string userID, string password)
        {
            try
            {
                FtpWebRequest ftpWebRequest = WebRequest.Create(targetURI) as FtpWebRequest;

                ftpWebRequest.Credentials = new NetworkCredential(userID, password);
                ftpWebRequest.Method = WebRequestMethods.Ftp.DeleteFile;

                FtpWebResponse ftpWebResponse = ftpWebRequest.GetResponse() as FtpWebResponse;
            }
            catch
            {
                return false;
            }

            return true;
        }

        private string[] GetFtpFileList(string targetURI, string userID, string password)
        {
            FtpWebRequest reqFtp = (FtpWebRequest)WebRequest.Create(targetURI);
            reqFtp.Credentials = new NetworkCredential(userID, password);
            reqFtp.Timeout = 10000;
            reqFtp.Method = WebRequestMethods.Ftp.ListDirectory;

            if (IsExistDirectory(targetURI, userID, password))
            {
                FtpWebResponse resFtp = (FtpWebResponse)reqFtp.GetResponse();

                StreamReader reader;
                reader = new StreamReader(resFtp.GetResponseStream());

                string strData;
                strData = reader.ReadToEnd();

                string[] filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                resFtp.Close();

                return filesInDirectory;
            }
            else
            {
                return null;
            }
            
        }

        private bool IsExistDirectory(string Directory, string _FTPuserID, string _FTPpassword)
        {
            try
            {
                var request = (FtpWebRequest)WebRequest.Create(Directory);
                request.Method = WebRequestMethods.Ftp.ListDirectory;
                request.Credentials = new NetworkCredential(_FTPuserID, _FTPpassword);

                using (request.GetResponse())
                {
                    return true;
                }
            }
            catch (WebException)
            {
                return false;
            }
        }
        #endregion

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (GridViewRetr.IsRowSelected(e.RowHandle))
                GridViewRetr.UnselectRow(e.RowHandle);
            else
                GridViewRetr.SelectRow(e.RowHandle);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private DataTable GetInspectInfo(string sMakeNo)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.MAKENO_LM ");
            strSql.AppendLine("      , A.M4_CARNO ");
            strSql.AppendLine("      , A.M4_CVCOD ");
            strSql.AppendLine("      , A.M4_CVNAM ");
            strSql.AppendLine("      , A.M4_CVNAM_IDTNO ");
            strSql.AppendLine("      , A.M4_GRADE ");
            strSql.AppendLine("      , A.M4_GRADE_CD ");
            strSql.AppendLine("      , A.M4_WGT ");
            strSql.AppendLine("      , A.M4_MINUS ");
            strSql.AppendLine("      , A.M4_EVIDENCE ");
            strSql.AppendLine("      , A.IMG_CNT ");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , A.ENT_ID ");
            strSql.AppendLine("      , A.MFY_DT ");
            strSql.AppendLine("      , A.MFY_ID ");
            strSql.AppendLine("   FROM MAKE_4 A ");
            strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void PD01001F04_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.F4)
            {
                BtnDel_Click(null, null);
            }
            else if(e.KeyCode == Keys.Escape)
            {
                BtnClose_Click(null, null);
            }
        }
    }
}