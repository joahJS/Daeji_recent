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
using System.Runtime.InteropServices;
using System.Net;
using System.Diagnostics;


namespace AccAdm
{
    public partial class FmLogin : DevExpress.XtraEditors.XtraForm
    {
        public FmLogin()
        {
            InitializeComponent();
        }

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

        DataRow drUserInfo;
        private void FmLogin_Load(object sender, EventArgs e)
        {
            DBConn.dbCon = DBConn.DbConn();
            string filePath = Application.StartupPath + @"\VersionCheck.ini";
            iniUtil ini = new iniUtil(filePath);
            string LoginCheck = ini.GetIniValue("LOGIN", "cheked");
            string sVersion = ini.GetIniValue("VERSION", "version");
            FmMainToolBar2._VERSION_ID = sVersion;
            if (LoginCheck.Equals("True"))
            {
                TxtId.Text= ini.GetIniValue("LOGIN", "id");

                string sPw = ini.GetIniValue("LOGIN", "pw");

                if (!string.IsNullOrEmpty(sPw))
                {
                    sPw = ComnEtcFunc.Decrypt(sPw, ComnEtcFunc._SECRET_KEY2);
                }

                TxtPw.Text = sPw;
                ChkSave.Checked = true;
            }

            if (!string.IsNullOrEmpty(sVersion))
            {
                LblVersion.Text = string.Format("Ver_{0}({1})", sVersion, GetUpdateRemark(sVersion));
            }
        }

        private void BtnLogin_Click(object sender, EventArgs e)
        {
             if (TxtId.Text.Trim() == "" || TxtPw.Text.Trim() == "")
             {
                 if (MessageBox.Show("아이디와 암호를 입력해 주세요.", "아이디 입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.Retry)
                 {
                     TxtId.Text = "";
                     TxtPw.Text = "";
                 }

             }
             else
             {
                 int res = Authenticate(TxtId.Text.Trim(), TxtPw.Text.Trim());
                 if (res > 0)
                 {

                     if (TxtId.Text == "daejiW")
                     {
                         ProcessStartInfo startInfo = new ProcessStartInfo();
                         string sStartupPath = Application.StartupPath;
                         

                        if(sStartupPath.Contains("AccAdm"))  //디버그 모드와 배포파일의 디렉토리 구조가 달라서 추가한 로직
                        {
                            string sStartupPathnew = sStartupPath.Replace("AccAdm", "WeighingSystem");
                            startInfo.FileName = string.Concat(sStartupPathnew, @"\WeighingSystem.exe");
                        }
                        else
                        {
                            startInfo.FileName = string.Concat(sStartupPath, @"\WeighingSystem\WeighingSystem.exe");
                        }

                         
                         startInfo.Arguments = "";

                         try
                         {
                             this.DialogResult = DialogResult.OK;
                             this.Visible = false;

                             if (ChkSave.Checked == true)
                             {

                                 string id = TxtId.Text;
                                 string pw = TxtPw.Text;
                                 if (!string.IsNullOrEmpty(pw))
                                 {
                                     pw = ComnEtcFunc.Encrypt(pw, ComnEtcFunc._SECRET_KEY2);
                                 }
                                 string cheked = ChkSave.EditValue.ToString();

                                 IniSet("id", id);
                                 IniSet("pw", pw);
                                 IniSet("cheked", cheked);
                             }
                             else
                             {
                                 string cheked = ChkSave.EditValue.ToString();
                                 IniSet("cheked", cheked);
                             }

                             Process process = Process.Start(startInfo);
                             //process.WaitForExit();
                             Application.Exit();
                         }
                         catch (Exception ex)
                         {

                         }
                     }
                     else if (TxtId.Text == "daejiM")
                     {
                         ProcessStartInfo startInfo = new ProcessStartInfo();
                         string sStartupPath = Application.StartupPath;
                         

                        if (sStartupPath.Contains("AccAdm"))  //디버그 모드와 배포파일의 디렉토리 구조가 달라서 추가한 로직
                        {
                            string sStartupPathnew = sStartupPath.Replace("AccAdm", "Daeji_MONITERING");
                            startInfo.FileName = string.Concat(sStartupPathnew, @"\Daeji_MONITERING.exe");
                        }
                        else
                        {
                            startInfo.FileName = string.Concat(sStartupPath, @"\Daeji_MONITERING\Daeji_MONITERING.exe");
                        }

                         startInfo.Arguments = "";

                         try
                         {
                             this.DialogResult = DialogResult.OK;
                             this.Visible = false;

                             if (ChkSave.Checked == true)
                             {

                                 string id = TxtId.Text;
                                 string pw = TxtPw.Text;
                                 if (!string.IsNullOrEmpty(pw))
                                 {
                                     pw = ComnEtcFunc.Encrypt(pw, ComnEtcFunc._SECRET_KEY2);
                                 }
                                 string cheked = ChkSave.EditValue.ToString();

                                 IniSet("id", id);
                                 IniSet("pw", pw);
                                 IniSet("cheked", cheked);
                             }
                             else
                             {
                                 string cheked = ChkSave.EditValue.ToString();
                                 IniSet("cheked", cheked);
                             }

                             Process process = Process.Start(startInfo);
                             //process.WaitForExit();
                             Application.Exit();


                         }
                         catch (Exception ex)
                         {

                         }
                     }
                     else
                     {
                         this.DialogResult = DialogResult.OK;
                         FmMainToolBar2 fm = new FmMainToolBar2();
                         FmMainToolBar2.drUser = drUserInfo;
                         this.Visible = false;
                         if (ChkSave.Checked == true)
                         {

                             string id = TxtId.Text;
                             string pw = TxtPw.Text;
                             if (!string.IsNullOrEmpty(pw))
                             {
                                 pw = ComnEtcFunc.Encrypt(pw, ComnEtcFunc._SECRET_KEY2);
                             }
                             string cheked = ChkSave.EditValue.ToString();

                             IniSet("id", id);
                             IniSet("pw", pw);
                             IniSet("cheked", cheked);
                             //ini.SetIniValue("LOGIN", "id", id);
                             //ini.SetIniValue("LOGIN", "pw", pw);
                             //ini.SetIniValue("LOGIN", "cheked", cheked);
                         }
                         else
                         {
                             string cheked = ChkSave.EditValue.ToString();
                             IniSet("cheked", cheked);
                         }
                         fm.Show();
                     }
                 }
                 else
                 {
                     switch (res)
                     {
                         case -1:
                             if (XtraMessageBox.Show("아이디를 잘못 입력하였습니다.", "사용자 로그인 오류", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning) == DialogResult.Retry)
                             {
                                 TxtId.Text = "";
                                 TxtPw.Text = "";
                             }
                             else
                             {
                                 this.DialogResult = DialogResult.Cancel;
                             }
                             break;
                         case -2:
                             if (XtraMessageBox.Show("암호를 잘못 입력하였습니다.", "사용자 로그인 오류", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning) == DialogResult.Retry)
                             {
                                 TxtPw.Text = "";
                             }
                             else
                                 this.DialogResult = DialogResult.Cancel;
                             break;

                         default:
                             XtraMessageBox.Show("ISM-DVS 시스템 오류입니다. 관리자에게 문의하세요.", "사용자 로그인 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                             this.DialogResult = DialogResult.Cancel;

                             break;
                     }
                 }
             }
            
        }

        private int Authenticate(string id, string pw)
        {
            string sid = TxtId.Text;
            string spw = TxtPw.Text;
            if (!string.IsNullOrEmpty(spw))
            {
                spw = ComnEtcFunc.Encrypt(spw, ComnEtcFunc._SECRET_KEY2);
            }

            StringBuilder strSql = new StringBuilder();


            strSql.AppendLine(" SELECT A.USRID             ");
			strSql.AppendLine("      , A.PASSWD            ");
			strSql.AppendLine("      , A.USRCD             ");
			strSql.AppendLine("      , A.USRNM             ");
			strSql.AppendLine("      , A.INSANO            ");
			strSql.AppendLine("      , B.DEPT_CD AS DEPTCD ");
			strSql.AppendLine("      , B.GRADE_CD AS JKWICD");
            strSql.AppendLine("   FROM ZUSRLST A           ");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B ");
            strSql.AppendLine("     ON A.INSANO = B.EMP_ID ");
             
            DataTable dtLogin = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dtLogin.Rows.Count > 0)
            {
                for (int i = 0; i < dtLogin.Rows.Count; i++)
                {
                    if (sid.Equals(dtLogin.Rows[i]["USRID"].ToString()))
                    {
                        if (spw.Equals(dtLogin.Rows[i]["PASSWD"]))
                        {
                            drUserInfo = dtLogin.Rows[i];
                            return 1;
                        }
                        else
                        {
                            return -2;
                        }
                    }
                    else
                    {
                        continue;
                    }

                }
                return -98;
            }
            else
            {
                return -99;
            }
        }

        private void IniSet(string key, string var)
        {
            string filePath = Application.StartupPath + @"\VersionCheck.ini";
            iniUtil ini = new iniUtil(filePath);
            ini.SetIniValue("LOGIN", key, var);
        }

        private string GetUpdateRemark(string sVersionId)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT CASE WHEN MAX(A.VERSION_ID) = '" + sVersionId + "' THEN '' ELSE '업데이트를 진행하세요.' END AS UPDATE_RMK ");
            strSql.AppendLine("   FROM ZSYS_VERSION A ");

            DataTable dtResult = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dtResult.Rows.Count > 0)
            {
                return dtResult.Rows?[0]["UPDATE_RMK"]?.ToString();
            }
            else
            {
                return null;
            }
        }

        private void BtnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void FmLogin_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void labelControl1_Click(object sender, EventArgs e)
        {

        }

        private void TxtPw_Enter(object sender, EventArgs e)
        {

        }

        private void TxtPw_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnLogin_Click(null, null);
        }
    }

}