using DevExpress.XtraEditors;
using DevExpress.XtraSplashScreen;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Daeji_UPDATE
{
    public partial class Form1 : DevExpress.XtraEditors.XtraForm
    {
        public Form1()
        {
            InitializeComponent();
        }

        public string BIG_CATE = "VERSION";
        public string INIT_NAME = "version";
        public string PROCEDURE_ID = "DP_ZSYS_VERSION_CHK";

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
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
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(SplashScreen1), true, false);
            DevExpress.XtraSplashScreen.SplashScreenManager.CloseForm();
            DBConn.dbCon = DBConn.DbConn();
            string filePath = Application.StartupPath + @"\VersionCheck.ini";
            LblProgress.Text = "최신버전 업데이트 체크를 진행하겠습니다.";
            if (!CheckVersion(filePath))
            {
                LblProgress.ForeColor = Color.Red;
                LblProgress.Text = "현재 버전은 최신버전이 아닙니다. 업데이트 버튼을 눌러 업데이트를 진행하세요.";
                LayoutUpdate.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }
            else
            {
                LblProgress.ForeColor = Color.Black;
                LblProgress.Text = "현재 버전은 최신버전입니다. 로그인 화면으로 이동합니다. 잠시만 기다려주세요...";
                LayoutUpdate.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                DelaySystem(500);
                CloseUpdate();
            }
        }

        private bool CheckVersion(string sPath)
        {
            iniUtil ini = new iniUtil(sPath);
            string sVersion = ini.GetIniValue(BIG_CATE, INIT_NAME);
            string sUpdateYn = GetUpdateRemark(sVersion);
            if (!string.IsNullOrEmpty(sUpdateYn))
            {
                if (sUpdateYn.Equals("Y"))
                    return true;
                else
                    return false;
            }
            else
            {
                return false;
            }
        }

        private string GetUpdateRemark(string sVersionId)
        {
            LblProgress.Text = "업데이트 체크를 진행중입니다...";
            StringBuilder strSql = new StringBuilder();

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Add("CMD", "VER_CHECK");
            dicParams.Add("VERSION_ID", sVersionId);
            
            DataTable dtResult = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
            if (dtResult.Rows.Count > 0)
            {
                return dtResult.Rows?[0]["UPDATE_YN"]?.ToString();
            }
            else
            {
                return null;
            }
        }

        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(SplashScreen1), true, false);

            try
            {
                byte[] file = null;
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("CMD", "LAST_VERSION");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                if (dt.Rows.Count > 0)
                {
                    string sVersionId = dt.Rows[0]["VERSION_ID"]?.ToString();
                    file = (byte[])dt.Rows[0]["EXEFILE"];
                    UpdateFile(file, sVersionId);
                    DevExpress.XtraSplashScreen.SplashScreenManager.CloseForm();
                    Form1_Load(null, null);
                }
                else
                {
                    throw new Exception("DB에 버전이 존재하지 않습니다.\r\n관리자에게 문의하세요.");
                }
            }
            catch (Exception ex)
            {
                DevExpress.XtraSplashScreen.SplashScreenManager.CloseForm();
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void UpdateFile(byte[] file, string sVersionId)
        {
            string filePath = "";

            //string fileDir = @"C:\Users\USER\Desktop\result";
            string fileDir = Application.StartupPath;
            filePath = fileDir + @"\Daeji.exe";
            FileStream fs;

            fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write);
            fs.Write(file, 0, file.Length);
            fs.Close();

            IniSet(BIG_CATE, INIT_NAME, sVersionId);
        }

        private void CloseUpdate()
        {
            string path = string.Format(@"{0}\{1}", Application.StartupPath, "Daeji.exe");
            Application.Exit();
            Process.Start(path);
        }

        private void IniSet(string sBicCate, string key, string var)
        {
            string filePath = Application.StartupPath + @"\VersionCheck.ini";
            iniUtil ini = new iniUtil(filePath);
            ini.SetIniValue(sBicCate, key, var);
        }

        private void DelaySystem(int MS)
        { /* 함수명 : DelaySystem * 1000ms = 1초 * 전달인자 : 얼마나 지연시킬것인가에 대한 변수 * */
            DateTime dtAfter = DateTime.Now;
            TimeSpan dtDuration = new TimeSpan(0, 0, 0, 0, MS);
            DateTime dtThis = dtAfter.Add(dtDuration);
            while (dtThis >= dtAfter)
            {
                System.Windows.Forms.Application.DoEvents(); //현재 시간 얻어 오기... 
                dtAfter = DateTime.Now;
            }
        }
    }
}
