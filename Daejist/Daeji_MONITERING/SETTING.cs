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
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Daeji_MONITERING.ComnLib;

namespace Daeji_MONITERING
{
    public partial class SETTING : DevExpress.XtraEditors.XtraForm
    {
        public SETTING()
        {
            InitializeComponent();
        }

        private string filePath = ComnString.INI_FILE_PATH;
        private IniUtil _ini = null;

        private void SETTING_Load(object sender, EventArgs e)
        {
            _ini = new IniUtil(filePath);

            TxtCam1.EditValue = _ini.GetIniValue("CAM", "CAM1");
            TxtCam2.EditValue = _ini.GetIniValue("CAM", "CAM2");

            TxtOilOndo.EditValue = _ini.GetIniValue("MAXNUM", "OILONDO");
            TxtApRyuk.EditValue = _ini.GetIniValue("MAXNUM", "APRYUK");
            TxtJunRu.EditValue = _ini.GetIniValue("MAXNUM", "JUNRU");
            TxtSecTime1.EditValue = _ini.GetIniValue("SECTIME", "TIME1");
        }

        private void SETTING_FormClosed(object sender, FormClosedEventArgs e)
        {
            SetFocuseForm();//경고창있을때 경고창에 포커스
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            #region 환경값 검사

            #region 경고설정값
            string sOilOndo = TxtOilOndo.EditValue?.ToString();

            if (string.IsNullOrEmpty(sOilOndo))
            {
                MessageBox.Show
                (
                        "오일온도를 설정하세요.  ",
                        "환경설정",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Stop
                );

                xtraTabControl1.SelectedTabPageIndex = 0;
                TxtOilOndo.Focus();

                return;
            }

            string sApRyuk = TxtApRyuk.EditValue?.ToString();

            if (string.IsNullOrEmpty(sApRyuk))
            {
                MessageBox.Show
                (
                        "압력을 설정하세요.  ",
                        "환경설정",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Stop
                );

                xtraTabControl1.SelectedTabPageIndex = 0;
                TxtApRyuk.Focus();

                return;
            }

            string sJunRu = TxtJunRu.EditValue?.ToString();

            if (string.IsNullOrEmpty(sJunRu))
            {
                MessageBox.Show
                (
                        "전류를 설정하세요.  ",
                        "환경설정",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Stop
                );

                xtraTabControl1.SelectedTabPageIndex = 0;
                TxtJunRu.Focus();

                return;
            }

            string sTime1 = TxtSecTime1.EditValue?.ToString();

            if (string.IsNullOrEmpty(sTime1))
            {
                MessageBox.Show
                (
                        "과전류 유지 초를 입력하세요.",
                        "환경설정",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Stop
                );

                xtraTabControl1.SelectedTabPageIndex = 0;
                TxtJunRu.Focus();

                return;
            }
            #endregion

            #region 카메라 설정
            string sCam1 = TxtCam1.EditValue?.ToString();

            if (string.IsNullOrEmpty(sCam1))
            {
                MessageBox.Show
                (
                        "카메라1 RTSP 주소를 입력하세요.",
                        "환경설정",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Stop
                );

                xtraTabControl1.SelectedTabPageIndex = 1;
                TxtCam1.Focus();

                return;
            }

            string sCam2 = TxtCam2.EditValue?.ToString();

            if (string.IsNullOrEmpty(sCam2))
            {
                MessageBox.Show
                (
                        "카메라2 RTSP 주소를 입력하세요.",
                        "환경설정",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Stop
                );

                xtraTabControl1.SelectedTabPageIndex = 1;
                TxtCam2.Focus();

                return;
            }
            #endregion

            #endregion 환경값 검사

            _ini.SetIniValue("MAXNUM", "OILONDO", sOilOndo);
            _ini.SetIniValue("MAXNUM", "APRYUK", sApRyuk);
            _ini.SetIniValue("MAXNUM", "JUNRU", sJunRu);
            _ini.SetIniValue("SECTIME", "TIME1", sTime1);

            _ini.SetIniValue("CAM", "CAM1", sCam1);
            _ini.SetIniValue("CAM", "CAM2", sCam2);

            MessageBox.Show("변경된 정보는 프로그램을 재시작하여야 적용됩니다.");

            SetFocuseForm();
            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            SetFocuseForm();
            this.Close();
        }

        private void SetFocuseForm()
        {
            Form fchk = Application.OpenForms["MS001F00"];

            if (fchk != null)
            {
                fchk.Focus();
            }
        }
    }
}