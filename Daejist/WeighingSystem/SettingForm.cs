using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

/*
 * RTSP로 변경(2024-03-13)
 * 테스트 시 변경 해야할 코드 _ 검색:테스트용 
 */


namespace WeighingSystem
{
    public partial class SettingForm : Form
    {
        private IniFile _ini = null;

        private bool _firstLoad = false;

        public SettingForm()
        {
            InitializeComponent();
            soundcheck(); //24.04.25
        }

        private void SettingForm_Load(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                _firstLoad = true;

                tabControl.SelectedIndex = 0;

                _ini = new IniFile(string.Concat(Application.StartupPath, @"\setting.ini"));

                textBoxStartWeight.Text = _ini.GetIniValue("ENV", "StartWeight");
                textBoxWeighingDay.Text = _ini.GetIniValue("ENV", "SecondWeighingDay");
                textBoxStableWeightCount.Text = _ini.GetIniValue("ENV", "StableWeightCount");
                textBoxFaxNumber.Text = _ini.GetIniValue("ENV", "FaxNumber");
                textBoxPointX.Text = _ini.GetIniValue("ENV", "PointX");
                textBoxPointY.Text = _ini.GetIniValue("ENV", "PointY");

                textBoxPortName.Text = _ini.GetIniValue("INDICATOR", "PortName");
                textBoxBaudRate.Text = _ini.GetIniValue("INDICATOR", "BaudRate");
                textBoxDataBit.Text = _ini.GetIniValue("INDICATOR", "DataBit");
                comboBoxStopBit.Text = _ini.GetIniValue("INDICATOR", "StopBit");
                comboBoxParityBit.Text = _ini.GetIniValue("INDICATOR", "Parity");

                textBoxDBAddress.Text = _ini.GetIniValue("DATABASE", "Address");
                textBoxDBName.Text = _ini.GetIniValue("DATABASE", "Database");
                textBoxDBUserID.Text = _ini.GetIniValue("DATABASE", "User");
                textBoxDBPassword.Text = _ini.GetIniValue("DATABASE", "Password", true);
                //textBoxDBPassword.Text = _ini.GetIniValue("DATABASE", "Password");   //로컬테스트용

                textBoxFtpAddress.Text = _ini.GetIniValue("FTP", "Address");
                textBoxFtpUserID.Text = _ini.GetIniValue("FTP", "User");
                textBoxFtpPassword.Text = _ini.GetIniValue("FTP", "Password", true);
                //textBoxFtpPassword.Text = _ini.GetIniValue("FTP", "Password");           //로컬테스트용
                textBoxFtpUploadPath.Text = _ini.GetIniValue("FTP", "UploadPath");

                TxtCam1.Text = _ini.GetIniValue("CAM", "CAM1");
                TxtCam2.Text = _ini.GetIniValue("CAM", "CAM2");
                TxtCam3.Text = _ini.GetIniValue("CAM", "CAM3");

                Combosound.Text = _ini.GetIniValue("SOUND", "SOUND"); //24.04,25
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            Regex regex_digit = new Regex(@"[0-9]");

            #region 환경값 검사

            #region 계량

            #region StartWeight

            string startWeight = textBoxStartWeight.Text.Trim();

            if (startWeight.Length == 0)
            {
                MessageBox.Show
                (
                     "계량시작중량을 설정하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(0);
                textBoxStartWeight.Focus();

                return;
            }

            if (!regex_digit.IsMatch(startWeight))
            {
                MessageBox.Show
                (
                     "숫자만 입력해 주세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(0);
                textBoxStartWeight.Focus();

                return;
            }

            #endregion

            #region WeighingDay

            string weighingDay = textBoxWeighingDay.Text.Trim();

            if (weighingDay.Length == 0)
            {
                MessageBox.Show
                (
                     "계량 가능 일 수를 설정하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(0);
                textBoxWeighingDay.Focus();

                return;
            }

            if (!regex_digit.IsMatch(weighingDay))
            {
                MessageBox.Show
                (
                     "숫자만 입력해 주세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(0);
                textBoxWeighingDay.Focus();

                return;
            }

            #endregion

            #region StableWeightCount

            string stableWeightCount = textBoxStableWeightCount.Text.Trim();

            if (stableWeightCount.Length == 0)
            {
                MessageBox.Show
                (
                     "계량 가능 일 수를 설정하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(0);
                textBoxStableWeightCount.Focus();

                return;
            }

            if (!regex_digit.IsMatch(stableWeightCount))
            {
                MessageBox.Show
                (
                     "숫자만 입력해 주세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(0);
                textBoxStableWeightCount.Focus();

                return;
            }

            #endregion

            #region FaxNumber

            string faxNumber = textBoxFaxNumber.Text.Trim();

            if (faxNumber.Length == 0)
            {
                MessageBox.Show
                (
                     "팩스번호를 입력하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(0);
                textBoxFaxNumber.Focus();

                return;
            }

            #endregion

            #region PointX

            string pointX = textBoxPointX.Text.Trim();

            if (pointX.Length == 0)
            {
                MessageBox.Show
                (
                     "프린터출력 기준포인트를 설정하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(0);
                textBoxPointX.Focus();

                return;
            }

            if (!regex_digit.IsMatch(pointX))
            {
                MessageBox.Show
                (
                     "숫자만 입력해 주세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(0);
                textBoxPointX.Focus();

                return;
            }

            #endregion

            #region PointY

            string pointY = textBoxPointY.Text.Trim();

            if (pointY.Length == 0)
            {
                MessageBox.Show
                (
                     "프린터출력 기준포인트를 설정하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(0);
                textBoxPointY.Focus();

                return;
            }

            if (!regex_digit.IsMatch(pointY))
            {
                MessageBox.Show
                (
                     "숫자만 입력해 주세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(0);
                textBoxPointY.Focus();

                return;
            }

            #endregion

            #endregion

            #region DB

            #region Address

            string dbAdress = textBoxDBAddress.Text.Trim();

            if (dbAdress.Length == 0)
            {
                MessageBox.Show
                (
                     "데이터베이스 서버주소를 입력하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(1);
                textBoxDBAddress.Focus();

                return;
            }

            #endregion

            #region Name

            string dbName = textBoxDBName.Text.Trim();

            if (dbName.Length == 0)
            {
                MessageBox.Show
                (
                     "데이터베이스명을 입력하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(1);
                textBoxDBName.Focus();

                return;
            }

            #endregion

            #region ID

            string dbUserID = textBoxDBUserID.Text.Trim();

            if (dbUserID.Length == 0)
            {
                MessageBox.Show
                (
                     "데이터베이스 접속계정을 입력하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(1);
                textBoxDBUserID.Focus();

                return;
            }

            #endregion

            #region password

            string dbPassword = textBoxDBPassword.Text.Trim();

            if (dbPassword.Length == 0)
            {
                MessageBox.Show
                (
                     "데이터베이스 접속 비밀번호를 입력세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(1);
                textBoxDBPassword.Focus();

                return;
            }

            #endregion

            #endregion
            
            #region Indicator

            #region PortName

            string portname = textBoxPortName.Text.Trim();

            if (portname.Length == 0)
            {
                MessageBox.Show
                (
                     "PortName 을 입력하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(2);
                textBoxPortName.Focus();

                return;
            }

            #endregion

            #region BaudRate

            string baudrate = textBoxBaudRate.Text.Trim();

            if (baudrate.Length == 0)
            {
                MessageBox.Show
                (
                     "BaudRate 를 설정하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(2);
                textBoxBaudRate.Focus();

                return;
            }

            if (!regex_digit.IsMatch(baudrate))
            {
                MessageBox.Show
                (
                     "숫자만 입력해 주세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(2);
                textBoxBaudRate.Focus();

                return;
            }

            #endregion

            #region DataBit

            string databit = textBoxDataBit.Text.Trim();

            if (databit.Length == 0)
            {
                MessageBox.Show
                (
                     "DataBit 를 설정하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(2);
                textBoxDataBit.Focus();

                return;
            }

            if (!regex_digit.IsMatch(databit))
            {
                MessageBox.Show
                (
                     "숫자만 입력해 주세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(2);
                textBoxDataBit.Focus();

                return;
            }

            #endregion

            #region StopBit

            string stopbit = comboBoxStopBit.Text.Trim();

            if (stopbit.Length == 0)
            {
                MessageBox.Show
                (
                     "StopBit 를 선택하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(2);
                comboBoxStopBit.Focus();

                return;
            }

            #endregion

            #region ParityBit

            string paritybit = comboBoxParityBit.Text.Trim();

            if (paritybit.Length == 0)
            {
                MessageBox.Show
                (
                     "ParutyBit 를 선택하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(2);
                comboBoxParityBit.Focus();

                return;
            }

            #endregion

            #endregion

            #region FTP

            #region Address

            string ftpAdress = textBoxFtpAddress.Text.Trim();

            if (ftpAdress.Length == 0)
            {
                MessageBox.Show
                (
                     "FTP 서버주소를 입력하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(3);
                textBoxFtpAddress.Focus();

                return;
            }

            #endregion
            
            #region ID

            string ftpUserID = textBoxFtpUserID.Text.Trim();

            if (ftpUserID.Length == 0)
            {
                MessageBox.Show
                (
                     "FTP 접속계정을 입력하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(3);
                textBoxFtpUserID.Focus();

                return;
            }

            #endregion

            #region password

            string ftpPassword = textBoxFtpPassword.Text.Trim();

            if (ftpPassword.Length == 0)
            {
                MessageBox.Show
                (
                     "FTP 접속 비밀번호를 입력세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(3);
                textBoxFtpPassword.Focus();

                return;
            }

            #region 기준경로

            string ftpUploadPath = textBoxFtpUploadPath.Text.Trim();

            if (ftpUploadPath.Length == 0)
            {
                MessageBox.Show
                (
                     "FTP 업로드 기준경로를 입력하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(3);
                textBoxFtpUploadPath.Focus();

                return;
            }

            #endregion

            #endregion

            #endregion

            #region CAM

            #region CAM1

            string sCAM1 = TxtCam1.Text.Trim();

            if (sCAM1.Length == 0)
            {
                MessageBox.Show
                (
                     "CAM1 주소를 입력하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(4);
                TxtCam1.Focus();

                return;
            }

            #endregion

            #region CAM2

            string sCAM2 = TxtCam2.Text.Trim();

            if (sCAM2.Length == 0)
            {
                MessageBox.Show
                (
                     "CAM2 주소를 입력하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(4);
                TxtCam2.Focus();

                return;
            }

            #endregion

            #region CAM3

            string sCAM3 = TxtCam3.Text.Trim();

            if (sCAM3.Length == 0)
            {
                MessageBox.Show
                (
                     "CAM3 주소를 입력하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(4);
                TxtCam3.Focus();

                return;
            }

            #endregion

            #endregion

            #region sSOUND 
            string sSOUND = Combosound.Text.Trim(); //24.04,25

            if (sSOUND.Length == 0)
            {
                MessageBox.Show
                (
                     "SOUND 주소를 입력하세요.  ",
                     "환경변수 검사",
                     MessageBoxButtons.OK,
                     MessageBoxIcon.Stop
                );

                tabControl.SelectTab(5);
                Combosound.Focus();

                return;
            }
            #endregion
            #endregion 환경값 검사

            _ini.SetIniValue("ENV", "StartWeight", startWeight);
            _ini.SetIniValue("ENV", "SecondWeighingDay", weighingDay);
            _ini.SetIniValue("ENV", "StableWeightCount", stableWeightCount);
            _ini.SetIniValue("ENV", "FaxNumber", faxNumber);
            _ini.SetIniValue("ENV", "PointX", pointX);
            _ini.SetIniValue("ENV", "PointY", pointY);

            _ini.SetIniValue("INDICATOR", "PortName", portname);
            _ini.SetIniValue("INDICATOR", "BaudRate", baudrate);
            _ini.SetIniValue("INDICATOR", "DataBit", databit);
            _ini.SetIniValue("INDICATOR", "StopBit", stopbit);
            _ini.SetIniValue("INDICATOR", "Parity", paritybit);

            _ini.SetIniValue("DATABASE", "Address", dbAdress);
            _ini.SetIniValue("DATABASE", "Database", dbName);
            _ini.SetIniValue("DATABASE", "User", dbUserID);
            _ini.SetIniValue("DATABASE", "Password", dbPassword, true);

            _ini.SetIniValue("FTP", "Address", ftpAdress);
            _ini.SetIniValue("FTP", "User", ftpUserID);
            _ini.SetIniValue("FTP", "Password", ftpPassword, true);
            _ini.SetIniValue("FTP", "UploadPath", ftpUploadPath);

            _ini.SetIniValue("CAM", "CAM1", sCAM1);
            _ini.SetIniValue("CAM", "CAM2", sCAM2);
            _ini.SetIniValue("CAM", "CAM3", sCAM3);

            _ini.SetIniValue("SOUND", "SOUND", sSOUND); 

            MessageBox.Show(new Form { TopMost = true }, "변경된 정보는 프로그램을 재시작하여야 적용됩니다.");

            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void soundcheck() //24.04.25
        {
            string selectedPath = string.Concat(Application.StartupPath, @"\sound");

            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(selectedPath);
            foreach (var fileInfo in di.GetFiles())
            {
                if (!Combosound.Items.Contains(fileInfo.Name))
                {
                    Combosound.Items.Add(fileInfo.Name);
                }
            }
        }
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "모든 파일 (*.*)|*.*|텍스트 파일 (*.txt)|*.txt";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFilePath = openFileDialog.FileName;

                    // 실행 파일이 위치한 폴더 경로에 'sound' 폴더를 만듭니다.
                    string destinationFolderPath = Path.Combine(Application.StartupPath, "sound");

                    // 만약 'sound' 폴더가 존재하지 않는다면 생성합니다.
                    if (!Directory.Exists(destinationFolderPath))
                    {
                        Directory.CreateDirectory(destinationFolderPath);
                    }

                    // 파일을 'sound' 폴더로 복사합니다.
                    string destinationFilePath = Path.Combine(destinationFolderPath, Path.GetFileName(selectedFilePath));
                    File.Copy(selectedFilePath, destinationFilePath, true);
                }
            }
            soundcheck();
        }
    }
}
