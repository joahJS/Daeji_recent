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
using System.Data.SqlClient;
using System.Diagnostics;
using System.Threading;

namespace AccAdm
{
    public partial class PopDocFile : DevExpress.XtraEditors.XtraForm
    {
        public PopDocFile()
        {
            InitializeComponent();
        }

        //FTP 서버
        private string sInitDir = ComnEtcFunc.FTP_ROOT + @"/ERP/AprlDocs/";
        private string user = ComnEtcFunc.FTP_USER;
        private string pw = ComnEtcFunc.FTP_PW;

        //임시 파일저장 경로
        private string tempDoctPath = Application.StartupPath + @"/tempDoct/" + FmMainToolBar2.drUser["USRCD"];

        public string _DOCTP;
        public string AddModifyGb { get; set; }
        public DataRow PDataRow;

        public delegate void SendDataHandler(string sComcd);
        public event SendDataHandler DataRowSendEvent;

        private void PopDocFile_Load(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();

            ComnEtcFunc.gp_SetColorFocused(layoutControl2);
            ComnEtcFunc.gp_SetColorFocused(layoutControl4);

            DataTable dtDept = GetLookupData();

            ComLib.ComGrid.SetLookUpEdit(LkupDeptcd, dtDept, "CD", "NM", "");

            if(AddModifyGb == "MOD")
            {
                TxtComnm.EditValue = PDataRow["DOCNM"];
                TxtFileNm.EditValue = PDataRow["FILNM"];
                LkupDeptcd.EditValue = PDataRow["USEDP"];
                TxtCELL1.EditValue = PDataRow["CELL1"];
                TxtCELL2.EditValue = PDataRow["CELL2"];
                TxtCELL3.EditValue = PDataRow["CELL3"];
                TxtCELL4.EditValue = PDataRow["CELL4"];
                TxtCELL5.EditValue = PDataRow["CELL5"];
                TxtCELL6.EditValue = PDataRow["CELL6"];
                TxtCELL7.EditValue = PDataRow["CELL7"];

                layoutControlItem9.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }

            ExcelCloseAndDelete();
        }

        private DataTable GetLookupData()
        {
            try
            {
                StringBuilder strSql = new StringBuilder();

                strSql.AppendLine(" SELECT 'ALL' AS CD  ");
                strSql.AppendLine("      , '공용' AS NM ");
                strSql.AppendLine("  UNION ALL          ");
                strSql.AppendLine(" SELECT DEPT_CD      ");
                strSql.AppendLine("      , DEPT_NM      ");
                strSql.AppendLine("   FROM ACC_DEPT_CD  ");

                return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString()); 
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }

            return null;
        }

        byte[] FILE = null;
        string sExtFileNm;
        private string _DialogGB { get; set; }
        private void BtnUpload_Click(object sender, EventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            using (XtraOpenFileDialog openFileDialog = new XtraOpenFileDialog())
            {
                openFileDialog.Title = "업로드할 파일을 선택해주세요.";
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel files (*.xls,*xlsx)|*.xls;*xlsx|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _DialogGB = "New";
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();
                    FILE = new byte[fileStream.Length];
                    fileStream.Read(FILE, 0, FILE.Length);

                    TxtFileNm.EditValue = Path.GetFileNameWithoutExtension(filePath);
                    sExtFileNm = Path.GetExtension(filePath);

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            string sComNm = string.Empty;
            string sFileNm = string.Empty;
            string sDeptCd = string.Empty;
            string sUser = FmMainToolBar2.UserID;

            sComNm = TxtComnm.EditValue?.ToString();
            sFileNm = TxtFileNm.EditValue?.ToString();
            sDeptCd = LkupDeptcd.EditValue?.ToString(); //공용:ALL

            string sCELL1 = TxtCELL1.EditValue?.ToString();
            string sCELL2 = TxtCELL2.EditValue?.ToString();
            string sCELL3 = TxtCELL3.EditValue?.ToString();
            string sCELL4 = TxtCELL4.EditValue?.ToString();
            string sCELL5 = TxtCELL5.EditValue?.ToString();
            string sCELL6 = TxtCELL6.EditValue?.ToString();
            string sCELL7 = TxtCELL7.EditValue?.ToString();

            if (string.IsNullOrEmpty(sFileNm))
            {
                XtraMessageBox.Show("파일명을 입력해주세요.");
                return;
            }

            if (sFileNm.Split('.').Length < 2)
            {
                sFileNm = sFileNm.Split('.')[0];
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT COUNT(*) AS CNT");
            strSql.AppendLine("   FROM DOCT_K         ");
            strSql.AppendLine("  WHERE FILNM LIKE '%" + sFileNm + "%'");
            if (!string.IsNullOrEmpty(_DOCTP))
            {
                strSql.AppendLine("    AND DOCTP != '" + _DOCTP + "'");
            }

            DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dtChk != null)
            {
                string sCnt = dtChk.Rows[0]["CNT"]?.ToString();
                int iCnt = int.Parse(sCnt);

                if(iCnt > 0)
                {
                    XtraMessageBox.Show("중복된 파일명이 존재합니다.\r\n다른 이름으로 저장해주세요.");
                    TxtFileNm.Focus();
                    return;
                }
            }

            if (AddModifyGb.Equals("ADD"))
            {
                strSql.Clear();
                strSql.AppendLine(" SELECT ISNULL(MAX(CONVERT(NUMERIC, DOCTP)),0)+1 AS DOCTP ");
                strSql.AppendLine("   FROM DOCT_K                                   ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if(dt != null)
                    _DOCTP = dt.Rows[0]["DOCTP"]?.ToString();
            }

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                if (_DialogGB != null && _DialogGB.Equals("Edit"))
                {
                    Process[] prcList = Process.GetProcessesByName("Excel");
                    if (prcList.Length != 0)
                    {
                        if (XtraMessageBox.Show("작성중인 결재 문서를 종료하시겠습니까? \r\n저장하지 않은 데이터는 복구할 수 없습니다.", "엑셀문서 확인", MessageBoxButtons.YesNo) != DialogResult.Yes)
                        {
                            return;
                        }

                        for (int i = 0; i < prcList.Length; i++)
                        {
                            Process prcessInfo = prcList[i];

                            string sProcessFileName = prcessInfo.MainWindowTitle;

                            if (sProcessFileName.Equals("전자결재 양식 작성.xlsx - Excel"))
                            {
                                prcessInfo.Kill();
                            }
                        }

                        Thread.Sleep(1000);
                    }
                }
               
                if (AddModifyGb.Equals("MOD"))
                {
                    string sBefoNm = PDataRow["FILNM"]?.ToString();
                    byte[] DownloadFile = null;
                    DownloadFile = ComnEtcFunc.DownloadFTPFile_ByteArray(sInitDir + sBefoNm, user, pw);

                    if(_DialogGB != null && _DialogGB.Equals("Edit"))
                    {
                        if (FILE == null && DownloadFile != null)
                        {
                            FileInfo fileInfo = new FileInfo(tempDoctPath + @"/전자결재 양식 작성.xlsx");
                            if (fileInfo.Exists)
                            {
                                if (sFileNm.Split('.').Length < 2)
                                {
                                    sFileNm = sFileNm.Split('.')[0];
                                }

                                ComnEtcFunc.DeleteFTPFile(sInitDir + sBefoNm, user, pw);
                                FILE = File.ReadAllBytes(tempDoctPath + @"/전자결재 양식 작성.xlsx");

                            }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(sBefoNm))
                        {
                            if (FILE == null && DownloadFile != null)
                            {
                                if (sBefoNm.Split('.').Length > 2)
                                {
                                    sExtFileNm = "." + sBefoNm.Split('.')[1];
                                }

                                FILE = DownloadFile;
                                ComnEtcFunc.DeleteFTPFile(sInitDir + sBefoNm, user, pw);
                            }
                            else if (FILE != null && DownloadFile != null)
                            {
                                ComnEtcFunc.DeleteFTPFile(sInitDir + sBefoNm, user, pw);
                            }
                        }
                    }
                    
                }

                if (FILE != null)
                {
                    if(_DialogGB != null && _DialogGB.Equals("Edit"))
                    {
                        if (sFileNm.Split('.').Length < 2)
                        {
                            sFileNm += ".xlsx";
                        }
                    }
                    else
                    {
                        if (sFileNm.Split('.').Length < 2)
                        {
                            sFileNm += sExtFileNm;
                        }
                    }

                    ComnEtcFunc.FTPDirectioryCheck(sInitDir, user, pw);
                    ComnEtcFunc.FTPUpload(sInitDir + sFileNm, user, pw, FILE);
                }
                else
                {
                    XtraMessageBox.Show("저장할 파일이 없습니다.");
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    return;
                }

                strSql.Clear();
                strSql.AppendLine(" IF EXISTS(SELECT* FROM DOCT_K WHERE DOCTP = '"+ _DOCTP + "')");
                strSql.AppendLine("    BEGIN                                       ");
                strSql.AppendLine("          UPDATE DOCT_K                         ");
                strSql.AppendLine("             SET USEDP = '"+ sDeptCd + "'                     ");
                strSql.AppendLine("               , DOCNM = '"+ sComNm + "'                     ");
                strSql.AppendLine("               , FILNM = '"+ sFileNm + "'                     ");
                strSql.AppendLine("               , CELL1 = '"+ sCELL1 + "'                     ");
                strSql.AppendLine("               , CELL2 = '"+ sCELL2 +"'                     ");
                strSql.AppendLine("               , CELL3 = '"+ sCELL3 +"'                     ");
                strSql.AppendLine("               , CELL4 = '"+ sCELL4 +"'                     ");
                strSql.AppendLine("               , CELL5 = '"+ sCELL5 +"'                     ");
                strSql.AppendLine("               , CELL6 = '"+ sCELL6 +"'                     ");
                strSql.AppendLine("               , CELL7 = '"+ sCELL7 +"'                     ");
                strSql.AppendLine("               , MUSER = '"+ sUser + "'                     ");
                strSql.AppendLine("               , MDATE = CONVERT(VARCHAR(20), GETDATE(), 20)");
                strSql.AppendLine("           WHERE DOCTP = '"+ _DOCTP + "'");
                strSql.AppendLine("      END                  ");
                strSql.AppendLine(" ELSE                      ");
                strSql.AppendLine("    BEGIN                  ");
                strSql.AppendLine("          INSERT INTO DOCT_K( DOCTP    ");
                strSql.AppendLine("                            , USEDP    ");
                strSql.AppendLine("                            , DOCNM    ");
                strSql.AppendLine("                            , FILNM    ");
                strSql.AppendLine("                            , CELL1    ");
                strSql.AppendLine("                            , CELL2    ");
                strSql.AppendLine("                            , CELL3    ");
                strSql.AppendLine("                            , CELL4    ");
                strSql.AppendLine("                            , CELL5    ");
                strSql.AppendLine("                            , CELL6    ");
                strSql.AppendLine("                            , CELL7    ");
                strSql.AppendLine("                            , CUSER )   ");
                strSql.AppendLine("                      VALUES( '"+ _DOCTP + "'  ");
                strSql.AppendLine("                            , '"+ sDeptCd + "'  ");
                strSql.AppendLine("                            , '"+ sComNm + "'  ");
                strSql.AppendLine("                            , '"+ sFileNm + "'  ");
                strSql.AppendLine("                            , '"+ sCELL1 +"'  ");
                strSql.AppendLine("                            , '"+ sCELL2 +"'  ");
                strSql.AppendLine("                            , '"+ sCELL3 +"'  ");
                strSql.AppendLine("                            , '"+ sCELL4 +"'  ");
                strSql.AppendLine("                            , '"+ sCELL5 +"'  ");
                strSql.AppendLine("                            , '"+ sCELL6 +"'  ");
                strSql.AppendLine("                            , '"+ sCELL7 +"'  ");
                strSql.AppendLine("                            , '"+ sUser + "' ) ");
                strSql.AppendLine("      END                              ");
                                                                                                                     
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");

                DataRowSendEvent(_DOCTP);
                DialogResult = DialogResult.OK;
            }
            catch(Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
            }
        }

        private void TxtFileNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnSave.Focus();
        }

        private void PopDocFile_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
        }

        private void TxtComnm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                TxtFileNm.Focus();
        }

        string _sEXT = ".xlsx";
        private void BtnEditDoct_Click(object sender, EventArgs e)
        {
            _DialogGB = "Edit";
            string sFileName = PDataRow["FILNM"]?.ToString();
            _sEXT = Path.GetExtension(sFileName);

            if (string.IsNullOrEmpty(sFileName))
            {
                XtraMessageBox.Show("저장된 파일이 없습니다.");
                return;
            }

            string downPath = tempDoctPath + @"/전자결재 양식 작성"+ _sEXT;

            if (!Directory.Exists(tempDoctPath))
            {
                Directory.CreateDirectory(tempDoctPath);
            }

            if (ComnEtcFunc.FTPFileDownload(sInitDir + sFileName, downPath, user, pw))
            {
                //XtraMessageBox.Show("다운로드를 완료하였습니다.");
                Process.Start(downPath);
            }
        }

        private void ExcelCloseAndDelete()
        {
            try
            {
                Process[] prcList = Process.GetProcessesByName("Excel");
                if (prcList.Length != 0)
                {
                    for (int i = 0; i < prcList.Length; i++)
                    {
                        Process prcessInfo = prcList[i];

                        string sProcessFileName = prcessInfo.MainWindowTitle;

                        if (sProcessFileName.Equals("전자결재 양식 작성"+ _sEXT + " - Excel"))
                        {
                            prcessInfo.Kill();
                        }
                    }
                }

                Thread.Sleep(1000);

                FileInfo fileInfo = new FileInfo(tempDoctPath + @"/전자결재 양식 작성"+ _sEXT + "");
                if (fileInfo.Exists)
                {
                    //파일삭제
                    fileInfo.Delete();
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }

        }

        private void PopDocFile_FormClosed(object sender, FormClosedEventArgs e)
        {
            ExcelCloseAndDelete();
        }
    }
}