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
using System.Net;
using System.IO;
using System.Runtime.InteropServices;
using ComLib;
using MySql.Data.MySqlClient;
using DevExpress.XtraGrid.Views.Grid;
using System.Windows.Markup;
using System.Globalization;
using System.Windows.Media;
using System.Drawing.Text;
using System.Data.SqlClient;


/*
 * 시작일자 : 2021-01-19 
 * 작성자 : 고혜성
 * 기능 : 스킨선택
 * -----------------------------MODIFY HISTORY---------------------------
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 */
namespace AccAdm
{
    public partial class SYS004F01 : DevExpress.XtraEditors.XtraForm
    {
        Dictionary<string, Image> _SKINS;
        DataRow _TempRow = null;
        PrivateFontSetting FontSetting;

        public SYS004F01()
        {
            InitializeComponent();

            //FTP에서 볼러올 이미지를 담을 변수 초기화
            _SKINS = new Dictionary<string, Image>();
            _SKINS.Add("DevExpress Style", null);
            _SKINS.Add("DevExpress Dark Style", null);
            _SKINS.Add("Office 2016 Colorful", null);
            _SKINS.Add("Office 2016 Dark", null);
            //_SKINS.Add("Office 2016 Black", null);
            _SKINS.Add("Office 2013 White", null);
            //_SKINS.Add("Office 2013 Dark Gray", null);
            _SKINS.Add("Office 2013 Light Gray", null);
            _SKINS.Add("Office 2010 Blue", null);
            //_SKINS.Add("Office 2010 Black", null);
            _SKINS.Add("Office 2010 Silver", null);
            _SKINS.Add("Visual Studio 2013 Blue", null);
            //_SKINS.Add("Visual Studio 2013 Dark", null);
            _SKINS.Add("Visual Studio 2013 Light", null);
            _SKINS.Add("Seven Classic", null);
            _SKINS.Add("Visual Studio 2010", null);
            _SKINS.Add("McSkin", null);
            //_SKINS.Add("Blueprint", null);
            
        }

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
        [DllImport("kernel32")]
        private static extern bool WritePrivateProfileSection(string lpAppName, string lpString, string lpFileName);

        class iniUtil
        {

            private string iniPath;

            public iniUtil(string path)
            {
                this.iniPath = path;  //INI 파일 위치를 생성할때 인자로 넘겨 받음
            }

            public String GetIniValue(String Section, String Key)
            {
                StringBuilder temp = new StringBuilder(255);
                int i = GetPrivateProfileString(Section, Key, "", temp, 255, iniPath);
                return temp.ToString();
            }

            // INI 값을 셋팅
            public void SetIniValue(String Section, String Key, String Value)
            {
                WritePrivateProfileString(Section, Key, Value, iniPath);
            }

            public bool WritePrivateProfileSection(string lpAppName, string lpString)
            {
                return WritePrivateProfileSection(lpAppName, lpString);
            }
        }

        List<string> fontNameList = new List<string>();
        public GridView[] arrGrdView;
        private void SYS004F01_Load(object sender, EventArgs e)
        {
            //this.AutoScaleMode = AutoScaleMode.
            try
            {
                GridRetr.DataSource = GetSkinImages();
                SetFont();
                SetRandomDealerInfo();
                //DataTable dt = (DataTable)GridRetr.DataSource;

                InstalledFontCollection installedFontCollection = new InstalledFontCollection();
                System.Drawing.FontFamily[] fontFamilies;
                // Get the array of FontFamily objects.
                fontFamilies = installedFontCollection.Families;

                // The loop below creates a large string that is a comma-separated
                // list of all font family names.

                int count = fontFamilies.Length;
                for (int j = 0; j < count; ++j)
                {
                    CboFont.Properties.Items.Add(fontFamilies[j].Name);
                    fontNameList.Add(fontFamilies[j].Name);
                }
                
                //XmlLanguage xmlLanguage = XmlLanguage.GetLanguage(CultureInfo.CurrentUICulture.Name);

                //foreach (System.Windows.Media.FontFamily fontFamily in Fonts.SystemFontFamilies)
                //{
                //    if (fontFamily.FamilyNames.ContainsKey(xmlLanguage))
                //    {
                //        CboFont.Properties.Items.Add(fontFamily.FamilyNames[xmlLanguage]);
                //        fontNameList.Add(fontFamily.FamilyNames[xmlLanguage]);
                //    }
                //    else
                //    {
                //        CboFont.Properties.Items.Add(fontFamily.ToString());
                //        fontNameList.Add(fontFamily.ToString());
                //    }
                //}

                arrGrdView = new GridView[] { GridViewTest };
                string sFont = CboFont.EditValue?.ToString();
                int iFontSize = 0;
                int.TryParse(TxtFontSize.EditValue?.ToString(), out iFontSize);
                int iRowHeight = 0;
                int.TryParse(TxtRowHeight.EditValue?.ToString(), out iRowHeight);

                if (string.IsNullOrEmpty(sFont))
                {
                    return;
                }
                else if (iFontSize <= 0)
                {
                    return;
                }
                else if (iRowHeight <= 0)
                {
                    return;
                }

                FontSetting.SetGridView(sFont, iFontSize, iRowHeight, GridViewTest);
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        
        private void BtnApply_Click(object sender, EventArgs e)
        {
            string sApplyInfo = TileView.GetFocusedRowCellValue(TileViewColName)?.ToString();
            SetSkin(sApplyInfo);

            string sSKIN_NAME = DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName;
            SetSkinInfo(sSKIN_NAME);
        }

        private void SetSkinInfo(string sSkinName)
        {
            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("USRCD", FmMainToolBar2.UserID);
                dicParams.Add("SKINNAME", sSkinName);

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" IF EXISTS(SELECT * FROM ZUSRSKIN WHERE USRCD = @USRCD ) ");
                strSql.AppendLine("    BEGIN ");
                strSql.AppendLine("          UPDATE ZUSRSKIN ");
                strSql.AppendLine("             SET SKINNAME = @SKINNAME ");
                strSql.AppendLine("               , MUSER = @USRCD ");
                strSql.AppendLine("               , MDATE = CONVERT(VARCHAR(19),GETDATE(),20) ");
                strSql.AppendLine("           WHERE USRCD = @USRCD; ");
                strSql.AppendLine("      END ");
                strSql.AppendLine(" ELSE ");
                strSql.AppendLine("    BEGIN ");
                strSql.AppendLine("          INSERT INTO ZUSRSKIN ");
                strSql.AppendLine("                    ( USRCD, SKINNAME, CUSER, CDATE ) ");
                strSql.AppendLine("              VALUES( @USRCD, @SKINNAME, @USRCD, CONVERT(VARCHAR(19),GETDATE(),20) ); ");
                strSql.AppendLine("      END ");
                
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.Clear();
                foreach (KeyValuePair<string, string> param in dicParams)
                {
                    cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                }
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();
                
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                XtraMessageBox.Show("스킨정보가 저장되었습니다.");
                Dispose();
            }
            catch(Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void TileView_ItemDoubleClick(object sender, DevExpress.XtraGrid.Views.Tile.TileViewItemClickEventArgs e)
        {
            string sApplyInfo = TileView.GetFocusedRowCellValue(TileViewColName)?.ToString();
            SetSkin(sApplyInfo);
        }

        private void TileView_ItemClick(object sender, DevExpress.XtraGrid.Views.Tile.TileViewItemClickEventArgs e)
        {
            
        }

        private void SetSkin(string sApplyInfo)
        {
            if (sApplyInfo.Equals("Visual Studio 2010"))
            {
                sApplyInfo = "VS2010";
                DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(sApplyInfo);
            }
            else if (sApplyInfo.Equals("Office 2013 White"))
            {
                sApplyInfo = "Office 2013";
                DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(sApplyInfo);
            }
            else if (sApplyInfo.Equals("McSkin"))
            {
                DevExpress.UserSkins.BonusSkins.Register();
                DevExpress.Skins.SkinManager.EnableFormSkins();
                DevExpress.LookAndFeel.UserLookAndFeel.Default.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Skin;
                DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName = "McSkin";
            }
            else
            {
                DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle(sApplyInfo);
            }
            
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private DataTable GetSkinImages()
        {
            DataTable dt = new DataTable();
            dt.TableName = "TABLE";

            //string ftpPath = @"ftp://192.168.0.202/ERP/Skin/";
            string ftpPath = @"ftp://"+ComnEtcFunc.FTP_IP+"/ERP/Skin/";
            string user = ComnEtcFunc.FTP_USER;
            string pw = ComnEtcFunc.FTP_PW;

            FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(ftpPath);

            req1.Credentials = new NetworkCredential(user, pw);
            req1.Method = WebRequestMethods.Ftp.ListDirectory;

            dt.Columns.Add("NAME", typeof(string));
            dt.Columns.Add("IMAGE", typeof(byte[]));
            dt.Columns.Add("SELECT", typeof(string));

            Dictionary<string, Image> dicCopy = new Dictionary<string, Image>();
            foreach(KeyValuePair<string, Image> param in _SKINS)
            {
                dicCopy.Add(param.Key, param.Value);
            }

            using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
            {
                StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                string strData = reader1.ReadToEnd();
                //폴더 내 파일이름
                string[] filesInDirectory = null;
                filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                reader1.Close();

                foreach (KeyValuePair<string, Image> skin in _SKINS)
                {
                    try
                    {
                        Cursor = Cursors.WaitCursor;

                        DataRow row = dt.NewRow();
                        row["NAME"] = skin.Key;

                        int findIndex = Array.FindIndex(filesInDirectory, i => i == string.Format("{0}.png", skin.Key));
                        if (findIndex >= 0)
                        {
                            row["IMAGE"] = DownloadFTPFile(string.Format(@"{0}/{1}", ftpPath, string.Format("{0}.png", skin.Key)), user, pw);
                        }
                        row["SELECT"] = "N";

                        dt.Rows.Add(row);
                        Cursor = Cursors.Default;
                    }
                    catch(Exception ex)
                    {
                        Cursor = Cursors.Default;
                    }
                }
            }
            
            return dt;
        }

        private void SetFont()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT FONT ");
            strSql.AppendLine("      , FONT_SIZE ");
            strSql.AppendLine("      , ROW_HEIGHT ");
            strSql.AppendLine("   FROM ZUSRSKIN ");
            strSql.AppendLine("  WHERE USRCD = " + FmMainToolBar2.UserID + " ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if(dt.Rows.Count == 0)
            {
                DataRow row = null;
                FontSetting = new PrivateFontSetting(row);
            }
            else
            {
                FontSetting = new PrivateFontSetting(dt.Rows[0]);

                CboFont.EditValue = dt.Rows[0]["FONT"];
                TxtFontSize.EditValue = dt.Rows[0]["FONT_SIZE"];
                TxtRowHeight.EditValue = dt.Rows[0]["ROW_HEIGHT"];
            }
        }

        private void SetRandomDealerInfo()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            #region mariaDB
            //strSql.AppendLine(" SELECT DEALER_CD ");
            //strSql.AppendLine("      , DEALER_NM ");
            //strSql.AppendLine("      , IDT_NO ");
            //strSql.AppendLine("   FROM ACC_DEALER_CD ");
            //strSql.AppendLine("  WHERE DEALER_GB = '운송' ");
            //strSql.AppendLine("  ORDER BY RAND() LIMIT 20; ");
            #endregion

            strSql.AppendLine("SELECT TOP 20 DEALER_CD  ");
            strSql.AppendLine("     , DEALER_NM         ");
            strSql.AppendLine("     , IDT_NO            ");
            strSql.AppendLine("  FROM ACC_DEALER_CD     ");
            strSql.AppendLine(" WHERE DEALER_GB = '운송'");
            strSql.AppendLine(" ORDER BY NEWID()        ");

            GridTest.DataSource = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        #region FTP 파일 다운로드하기 - DownloadFTPFile(sourceFileURI, targetFilePath, userID, password)

        /// <summary>
        /// FTP 파일 다운로드하기
        /// </summary>
        /// <param name="sourceFileURI">소스 파일 URI</param>
        /// <param name="targetFilePath">타겟 파일 경로</param>
        /// <param name="userID">사용자 ID</param>
        /// <param name="password">패스워드</param>
        /// <returns>처리 결과</returns>

        public byte[] DownloadFTPFile(string sourceFileURI, string user, string pw)
        {
            byte[] data = null;
            try
            {
                Uri sourceFileUri = new Uri(sourceFileURI);
                FtpWebRequest ftpWebRequest = WebRequest.Create(sourceFileUri) as FtpWebRequest;

                ftpWebRequest.Credentials = new NetworkCredential(user, pw);
                ftpWebRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                FtpWebResponse ftpWebResponse = ftpWebRequest.GetResponse() as FtpWebResponse;

                using (Stream sourceStream = ftpWebResponse.GetResponseStream())
                {
                    Image img = Image.FromStream(sourceStream);
                    data = ImageToByteArray(img);
                }
            }
            catch (Exception ex)
            {
                return null;
            }

            return data;
        }

        #endregion
        public byte[] ImageToByteArray(System.Drawing.Image imageIn)
        {
            using (var ms = new MemoryStream())
            {
                imageIn.Save(ms, imageIn.RawFormat);
                return ms.ToArray();
            }
        }

        private void BtnPreview_Click(object sender, EventArgs e)
        {
            string sFont = CboFont.EditValue?.ToString();
            int iFontSize = 0;
            int.TryParse(TxtFontSize.EditValue?.ToString(), out iFontSize);
            int iRowHeight = 0;
            int.TryParse(TxtRowHeight.EditValue?.ToString(), out iRowHeight);

            if (string.IsNullOrEmpty(sFont))
            {
                XtraMessageBox.Show("폰트를 선택하세요.");
                CboFont.Focus();
                return;
            }
            else if(iFontSize <= 0)
            {
                XtraMessageBox.Show("폰트사이즈를 0 이상 입력하세요.");
                TxtFontSize.SelectAll();
                TxtFontSize.Focus();
                return;
            }
            else if(iRowHeight <= 0)
            {
                XtraMessageBox.Show("행높이를 0 이상 입력하세요.");
                TxtRowHeight.SelectAll();
                TxtRowHeight.Focus();
                return;
            }

            try
            {
                FontSetting.SetGridView(sFont, iFontSize, iRowHeight, GridViewTest);
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewTest_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewTest_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void BtnApply2_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                string sFont = CboFont.EditValue?.ToString();
                int iFontSize = 0;
                int.TryParse(TxtFontSize.EditValue?.ToString(), out iFontSize);
                int iRowHeight = 0;
                int.TryParse(TxtRowHeight.EditValue?.ToString(), out iRowHeight);

                if (string.IsNullOrEmpty(sFont))
                {
                    XtraMessageBox.Show("폰트를 선택하세요.");
                    CboFont.Focus();
                    return;
                }
                else if (iFontSize <= 0)
                {
                    XtraMessageBox.Show("폰트사이즈를 0 이상 입력하세요.");
                    TxtFontSize.SelectAll();
                    TxtFontSize.Focus();
                    return;
                }
                else if (iRowHeight <= 0)
                {
                    XtraMessageBox.Show("행높이를 0 이상 입력하세요.");
                    TxtRowHeight.SelectAll();
                    TxtRowHeight.Focus();
                    return;
                }

                bool bYn = false;
                foreach(string font in fontNameList)
                {
                    if (font.Equals(sFont))
                    {
                        bYn = true;
                        break;
                    }
                }

                if (!bYn)
                {
                    XtraMessageBox.Show("해당 폰트는 사용할 수 없습니다.\r\n다른 폰트를 선택하세요");
                    return;
                }

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("USRCD", FmMainToolBar2.UserID);
                dicParams.Add("FONT", CboFont.EditValue?.ToString());
                dicParams.Add("FONT_SIZE", TxtFontSize.EditValue?.ToString());
                dicParams.Add("ROW_HEIGHT", TxtRowHeight.EditValue?.ToString());

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" IF EXISTS(SELECT * FROM ZUSRSKIN WHERE USRCD = @USRCD )");
                strSql.AppendLine("    BEGIN ");
                strSql.AppendLine("          UPDATE ZUSRSKIN ");
                strSql.AppendLine("             SET FONT = @FONT ");
                strSql.AppendLine("               , FONT_SIZE = @FONT_SIZE ");
                strSql.AppendLine("               , ROW_HEIGHT = @ROW_HEIGHT ");
                strSql.AppendLine("               , MUSER = @USRCD ");
                strSql.AppendLine("               , MDATE = CONVERT(VARCHAR(19),GETDATE(),20)  ");
                strSql.AppendLine("           WHERE USRCD = @USRCD; ");
                strSql.AppendLine("      END ");
                strSql.AppendLine(" ELSE ");
                strSql.AppendLine("    BEGIN ");
                strSql.AppendLine("          INSERT INTO ZUSRSKIN ");
                strSql.AppendLine("                    ( USRCD, SKINNAME, FONT, FONT_SIZE, ROW_HEIGHT, CUSER, CDATE ) ");
                strSql.AppendLine("              VALUES( @USRCD, 'McSkin', @FONT, @FONT_SIZE, @ROW_HEIGHT, @USRCD, CONVERT(VARCHAR(19),GETDATE(),20)  ); ");
                strSql.AppendLine("      END ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.Clear();
                foreach (KeyValuePair<string, string> param in dicParams)
                {
                    cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                }
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                Cursor = Cursors.Default;

                XtraMessageBox.Show("폰트정보가 저장되었습니다.");
                Dispose();
                
                FmMainToolBar2._FontSetting.SetFont(CboFont.EditValue?.ToString());
                FmMainToolBar2._FontSetting.SetFontSize(iFontSize);
                FmMainToolBar2._FontSetting.SetRowHeight(iRowHeight);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }
    }

    public class PrivateFontSetting
    {
        private string FONT = "맑은 고딕";
        private int FONT_SIZE = 9;
        private int ROW_HEIGHT = 23;
        private int COLUMN_HEADER_HEIGHT = 0;

        private int interval = 7;

        public PrivateFontSetting(DataRow row)
        {
            if(row != null)
            {
                string sFont = row["FONT"]?.ToString();
                if (!string.IsNullOrEmpty(sFont))
                {
                    FONT = sFont;
                }

                int iFontSize = 0;
                int.TryParse(row["FONT_SIZE"]?.ToString(), out iFontSize);

                int iRowHeight = 0;
                int.TryParse(row["ROW_HEIGHT"]?.ToString(), out iRowHeight);
                
                if(iFontSize > 0)
                {
                    FONT_SIZE = iFontSize;
                }

                if(iRowHeight > 0)
                {
                    ROW_HEIGHT = iRowHeight;
                }

                COLUMN_HEADER_HEIGHT = ROW_HEIGHT + interval;
            }
        }

        public PrivateFontSetting(string UserID)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT FONT ");
            strSql.AppendLine("      , FONT_SIZE ");
            strSql.AppendLine("      , ROW_HEIGHT ");
            strSql.AppendLine("   FROM ZUSRSKIN ");
            strSql.AppendLine("  WHERE USRCD = " + FmMainToolBar2.UserID + " ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if(dt.Rows.Count == 0)
            {
                COLUMN_HEADER_HEIGHT = ROW_HEIGHT + interval;
            }
            else
            {
                string sFont = dt.Rows[0]["FONT"]?.ToString();
                if (!string.IsNullOrEmpty(sFont))
                {
                    FONT = sFont;
                }

                int iFontSize = 0;
                int.TryParse(dt.Rows[0]["FONT_SIZE"]?.ToString(), out iFontSize);

                int iRowHeight = 0;
                int.TryParse(dt.Rows[0]["ROW_HEIGHT"]?.ToString(), out iRowHeight);

                if (iFontSize > 0)
                {
                    FONT_SIZE = iFontSize;
                }

                if (iRowHeight > 0)
                {
                    ROW_HEIGHT = iRowHeight;
                }

                COLUMN_HEADER_HEIGHT = ROW_HEIGHT + interval;
            }
            
        }

        public void SetFont(string Font)
        {
            this.FONT = Font;
            CalcColumnHeaderPanelHeight(this.ROW_HEIGHT);
        }

        public void SetFontSize(int FontSize)
        {
            this.FONT_SIZE = FontSize;
            CalcColumnHeaderPanelHeight(this.ROW_HEIGHT);
        }

        public void SetRowHeight(int RowHeight)
        {
            this.ROW_HEIGHT = RowHeight;
            CalcColumnHeaderPanelHeight(this.ROW_HEIGHT);
        }

        private void CalcColumnHeaderPanelHeight(int RowHeight)
        {
            this.COLUMN_HEADER_HEIGHT = RowHeight + this.interval;
        } 

        public string GetFont()
        {
            return FONT;
        }

        public int GetFontSize()
        {
            return FONT_SIZE;
        }

        public int GetRowHeight()
        {
            return ROW_HEIGHT;
        }

        public int GetColumnHeader()
        {
            return COLUMN_HEADER_HEIGHT;
        }

        public void SetGridView(string Font, int FontSize, int RowHeight, GridView view)
        {
            Font font = new Font(Font, FontSize);

            view.RowHeight = RowHeight;
            view.Appearance.Row.Font = font;

            view.ColumnPanelRowHeight = RowHeight + interval;
            view.Appearance.HeaderPanel.Font = font;
        }

        public void SetGridView(GridView view)
        {
            Font font = new Font(FONT, FONT_SIZE);

            view.RowHeight = ROW_HEIGHT;
            foreach(DevExpress.XtraGrid.Columns.GridColumn col in view.Columns)
            {
                col.AppearanceHeader.Font = font;
                col.AppearanceCell.Font = font;
            }
            
            view.ColumnPanelRowHeight = COLUMN_HEADER_HEIGHT;
        }
    }
}