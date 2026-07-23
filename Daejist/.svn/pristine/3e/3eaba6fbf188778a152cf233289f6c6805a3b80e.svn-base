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
using System.IO;
using ComLib;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class SYS001F02 : DevExpress.XtraEditors.XtraForm
    {
        public SYS001F02()
        {
            InitializeComponent();
        }

        public Form PARENT_FORM;
        public DataRow DR_FILE_INFO = null;

        private void SYS001F02_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            if (DR_FILE_INFO != null)
            {
                TxtVersionId.EditValue = DR_FILE_INFO["VERSION_ID"];
                DateEditUpload.EditValue = DR_FILE_INFO["UPLOAD_DT"];
                TxtFileName.EditValue = DR_FILE_INFO["FILE_NAME"];
                TxtFileByte.EditValue = DR_FILE_INFO["FILE_BYTE"];
                MemoVersionRmk.EditValue = DR_FILE_INFO["VERSION_RMK"];

                BtnUpload.Enabled = false;
            }
            else
            {
                DateEditUpload.EditValue = DateTime.Today.ToString("yyyy-MM-dd");
                BtnDownLoad.Enabled = false;
            }
        }

        byte[] FILE = null;
        private void BtnUpload_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;
            
            using (XtraOpenFileDialog openFileDialog = new XtraOpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Exe Files (.exe)|*.exe|All Files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();
                    FILE = new byte[fileStream.Length];
                    fileStream.Read(FILE, 0, FILE.Length);

                    TxtFileByte.EditValue = fileStream.Length;

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }

                }
            }

            XtraMessageBox.Show("저장을 눌러주세요.");
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            string sVersionId = TxtVersionId.EditValue?.ToString();
            string sUploadDt = DateEditUpload.EditValue?.ToString().Substring(0, 10);
            string sFileName = "Daeji";
            string sFileByte = TxtFileByte.EditValue?.ToString();
            string sVersionRmk = MemoVersionRmk.EditValue?.ToString();
            string sId = FmMainToolBar2.UserID;

            if(FILE == null && DR_FILE_INFO == null)
            {
                XtraMessageBox.Show("파일을 업로드하세요");
                return;
            }
            else if (string.IsNullOrEmpty(sVersionId))
            {
                XtraMessageBox.Show("VersionID를 입력하세요(예 : 1.1.10");
                TxtVersionId.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sUploadDt))
            {
                XtraMessageBox.Show("업로드일자를 입력하세요");
                DateEditUpload.Focus();
                return;
            }
            else if (string.IsNullOrEmpty(sVersionRmk))
            {
                XtraMessageBox.Show("비고를 입력하세요");
                MemoVersionRmk.Focus();
                return;
            }
            

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;
                
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                #region [mariaDB]
                //strSql.AppendLine(" INSERT INTO ZSYS_VERSION");
                //strSql.AppendLine("           ( ");
                //strSql.AppendLine("             VERSION_ID ");
                //strSql.AppendLine("           , UPLOAD_DT ");
                //strSql.AppendLine("           , FILE ");
                //strSql.AppendLine("           , FILE_NAME ");
                //strSql.AppendLine("           , FILE_BYTE ");
                //strSql.AppendLine("           , VERSION_RMK ");
                //strSql.AppendLine("           , ENT_DT ");
                //strSql.AppendLine("           , ENT_ID");
                //strSql.AppendLine("           ) ");
                //strSql.AppendLine("      VALUES ");
                //strSql.AppendLine("           ( ");
                //strSql.AppendLine("             @VERSION_ID ");
                //strSql.AppendLine("           , @UPLOAD_DT ");
                //strSql.AppendLine("           , @FILE ");
                //strSql.AppendLine("           , @FILE_NAME ");
                //strSql.AppendLine("           , @FILE_BYTE ");
                //strSql.AppendLine("           , @VERSION_RMK ");
                //strSql.AppendLine("           , NOW() ");
                //strSql.AppendLine("           , @ENT_ID ");
                //strSql.AppendLine("           ) ");
                //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                //strSql.AppendLine("             VERSION_RMK = @VERSION_RMK ");
                //strSql.AppendLine("           , MOD_ID = @MOD_ID ");
                //strSql.AppendLine("           , MOD_DT = NOW() ");
                #endregion

                strSql.AppendLine("IF EXISTS(SELECT* FROM ZSYS_VERSION WHERE VERSION_ID = @VERSION_ID) ");
                strSql.AppendLine("   BEGIN                                                           ");
                strSql.AppendLine("         UPDATE ZSYS_VERSION                                       ");
                strSql.AppendLine("            SET VERSION_RMK = @VERSION_RMK                         ");
	            strSql.AppendLine("              , MOD_ID = @MOD_ID                                   ");
                strSql.AppendLine("              , MOD_DT = GETDATE()                                 ");
                strSql.AppendLine("          WHERE VERSION_ID = @VERSION_ID                           ");
                strSql.AppendLine("     END                                                           ");
                strSql.AppendLine("ELSE                                                               ");
                strSql.AppendLine("   BEGIN                                                           ");
                strSql.AppendLine("         INSERT INTO ZSYS_VERSION                                  ");
                strSql.AppendLine("                   (                                               ");
                strSql.AppendLine("                     VERSION_ID                                    ");
                strSql.AppendLine("                   , UPLOAD_DT                                     ");
                if(FILE != null)
                {
                    strSql.AppendLine("                   , EXEFILE                                       ");
                }
                strSql.AppendLine("                   , FILE_NAME                                     ");
                strSql.AppendLine("                   , FILE_BYTE                                     ");
                strSql.AppendLine("                   , VERSION_RMK                                   ");
                strSql.AppendLine("                   , ENT_ID                                        ");
                strSql.AppendLine("                   )                                               ");
                strSql.AppendLine("                                                                   ");
                strSql.AppendLine("              VALUES                                               ");
                strSql.AppendLine("                   (                                               ");
                strSql.AppendLine("                     @VERSION_ID                                   ");
                strSql.AppendLine("                   , @UPLOAD_DT                                    ");
                if(FILE != null)
                {
                    strSql.AppendLine("                   , @EXEFILE                                      ");
                }
                strSql.AppendLine("                   , @FILE_NAME                                    ");
                strSql.AppendLine("                   , @FILE_BYTE                                    ");
                strSql.AppendLine("                   , @VERSION_RMK                                  ");
                strSql.AppendLine("                   , @ENT_ID                                       ");
                strSql.AppendLine("                   )                                               ");
                strSql.AppendLine("     END                                                           ");

                cmd.Parameters.AddWithValue("@VERSION_ID", sVersionId);
                cmd.Parameters.AddWithValue("@UPLOAD_DT", sUploadDt);
                if(FILE != null)
                {
                    cmd.Parameters.AddWithValue("@EXEFILE", FILE);
                }
                cmd.Parameters.AddWithValue("@FILE_NAME", sFileName);
                cmd.Parameters.AddWithValue("@FILE_BYTE", sFileByte);
                cmd.Parameters.AddWithValue("@VERSION_RMK", sVersionRmk);
                cmd.Parameters.AddWithValue("@ENT_ID", sId);
                cmd.Parameters.AddWithValue("@MOD_ID", sId);

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();

                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                XtraMessageBox.Show("저장이 완료되었습니다.");
                DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnDownLoad_Click(object sender, EventArgs e)
        {
            if (XtraMessageBox.Show("파일을 다운로드 하시겠습니까?", "다운로드 여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                return;

            string sVersionId = TxtVersionId.EditValue?.ToString();
            if (string.IsNullOrEmpty(sVersionId))
            {
                XtraMessageBox.Show("VersionID가 존재하지 않습니다.");
                return;
            }

            try
            {
                byte[] file = null;

                StringBuilder strSql = new StringBuilder();

                strSql.AppendLine(" SELECT EXEFILE ");
                strSql.AppendLine("   FROM zsys_version  ");
                strSql.AppendLine("  WHERE VERSION_ID = '" + sVersionId + "'  ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                if (dt.Rows.Count > 0)
                {
                    file = (byte[])dt.Rows[0]["EXEFILE"];
                }

                if (file != null)
                {
                    string filePath = "";

                    FileDialog fileDlg = new SaveFileDialog();
                    string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    fileDlg.InitialDirectory = sFolderPath;
                    fileDlg.FileName = "Daeji.exe";

                    if (fileDlg.ShowDialog() == DialogResult.OK)
                    {
                        filePath = fileDlg.FileName;
                    }
                    fileDlg.Dispose();

                    FileStream fs;

                    fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
                    fs.Write(file, 0, file.Length);
                    fs.Close();

                    XtraMessageBox.Show("파일다운로드를 완료했습니다.");
                }
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void SYS001F02_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.Escape){
            }
        }
        
    }
}