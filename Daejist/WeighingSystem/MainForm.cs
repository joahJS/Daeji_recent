using DevExpress.XtraEditors;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraWaitForm;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Media;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using Size = System.Drawing.Size;

/*
 * RTSP로 변경(2024-03-13)
 * 테스트 시 변경 해야할 코드 _ 검색:테스트용 
 * setting.ini 파일 개발에 맞게 수정 _ 위치 : debug파일
 */

namespace WeighingSystem
{
    public partial class MainForm : Form
    {
        #region Properties
        /// <summary>
        /// 계량 진행 상태
        /// </summary>
        public enum WeighingState { None, Standby, Start, Stable, Complete };
        public enum WeighingType { None, In, Out };
        public enum WeighingStep { None, First, Second }

        /*
        * 2020-12-10 계근미확정 3초 지속 시 상태값 변화를 위하여 변수 추가
        */
        public enum ConfirmYN { Confirm, NonConfirm };

        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

        [DllImport("VKB.dll", CharSet = CharSet.Auto)]
        static extern void InitHook(IntPtr hHandle);

        [DllImport("VKB.dll", CharSet = CharSet.Auto)]
        static extern void InstallHook();

        //[DllImport("user32.dll")]
        //static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        //[DllImport("user32.dll")]
        //static extern int SendMessage(IntPtr hWnd, uint message, int wParam, int lParam);



        delegate void SelectTabPageCallback(int no);
        delegate void SetCompleteButtonCallback(bool enabled);
        delegate void SetCompleteButtonCallback3(bool enabled);
        delegate void SetCompleteButtonCallback2(bool enabled, string sText);

        //2020-12-08 계근확정 시 아무런 액션이 없을 경우 해당 변수 사용하기 위하여 추가 / Ini파일 참조
        private int _autoCompleteSecond = 5;

        private ConfirmYN _confirm = ConfirmYN.NonConfirm;

        private readonly object _lockObject = new object();
        private readonly object _lockCamera1 = new object();
        private readonly object _lockCamera2 = new object();
        private readonly object _lockCamera3 = new object();

        private SerialPort _indicator = null;                               // INDICATOR
        private bool _indicatorPrevConnected = false;
        private bool _indicatorConnected = false;
        private DateTime _indicatorReceiveTime;

        private string _connectionString = string.Empty;                    // DB

        private SingleFtpClient _ftpClient = SingleFtpClient.Instance;      // FTP
        private string _currentFtpLocation = string.Empty;

        private bool _bVKLShiftButtonState = false;                          // true : SHIFT ON, false : SHIFT OFF
        private bool _bVKRShiftButtonState = false;                          // true : SHIFT ON, false : SHIFT OFF
        private bool _bVKKorEngButtonState = false;                         // 한영 전환용
        private bool _bVKCapsLock = false;                                  // Caps Lock

        private bool _firstLoad = false;
        private bool _shutdown = false;

        private System.Threading.Timer _indicatorTimer = null;
        private System.Threading.Timer _ftpTimer = null;

        private WeighingType _weighingType = WeighingType.None;             // 계량타입 : 입고, 출고
        private WeighingStep _weighingStep = WeighingStep.None;             // 계근선택 : 1차, 2차
        private string _vehicleNumber = string.Empty;                       // 차량번호   
        private bool _useWeighing = false;                                  // 계량진행여부 플레그     
        private bool _weightFlag = false;                                   // 계량시작 플레그
        private WeighingState _state = WeighingState.None;                  // 계량 진행 상태
        private int _stableCount = 0;                                       // 안정화된 인디케이터 누적 카운트
        private int _weight = 0;                                            // 계량중량

        // 디비 검색 정보들..
        // - acc_dealer_cd
        private string _selectedDealerCd = string.Empty;                    // 거래처코드
        private string _selectedDealerNm = string.Empty;                    // 거래처명
        private string _selectedFax = string.Empty;                         // 팩스번호
        private string _selectedWebFaxYN = string.Empty;                    // 웹팩스 사용여부
        // - mesuring
        private string _selectedJunpyoID = string.Empty;                    // 전표아이디
        private string _selectedJ_Date = string.Empty;                      // 일자
        private string _selectedSun = string.Empty;                         // 계근번호(순번)
        private string _selectedMaipCherID = string.Empty;                  // 매입처코드
        private string _selectedMaipCher = string.Empty;                    // 매입처
        private string _selectedJ_AssignID = string.Empty;                  // 매출처코드
        private string _selectedJ_Company = string.Empty;                   // 매출처
        private string _selectedFirstTime = string.Empty;                   // 1차계량 (입/출고에 따라 변경) - 시간
        private int _selectedFirstWeight = 0;                               //                             - 중량 
        private string _selectedSecondTime = string.Empty;                  // 2차계량 (입/출고에 따라 변경) - 시간
        private int _selectedSecondWeight = 0;                              //                             - 중량
        private string _selectedJ_Serial = string.Empty;                    // 검수정보 - 검수여부 ("0" 이 아니면 검수)
        private string _selectedGubun1 = string.Empty;                      //         - 등급
        private int _selectediChaGam = 0;                                   //         - 매입감량
        private int _selectedOChaGam = 0;                                   //         - 매출감량
        private string _selectedJ_State = string.Empty;                     //         - 감가/감량사유
        private string _selectedgumsubigo = string.Empty;                   //         - 검수비고
        private string _selectedgumsu_serial = string.Empty;                //         - 검수자 사원번호
        private string _selectedEMP_NM = string.Empty;                      //         - 검수자명

        public string _selected_JunpyoID = string.Empty;

        // - ticket image
        private string _selectedTicketImage_1_1 = string.Empty;
        private string _selectedTicketImage_1_2 = string.Empty;
        private string _selectedTicketImage_1_3 = string.Empty;
        private string _selectedTicketImage_2_1 = string.Empty;
        private string _selectedTicketImage_2_2 = string.Empty;
        private string _selectedTicketImage_2_3 = string.Empty;

        #region setting.ini 

        private IniFile _ini = null;

        private int _startWeight = 0;                               // 계량시작 중량
        private int _secondWeighingDay = 0;                         // 1차->2차계량 가능 일 수
        private int _stableWeightCount = 0;                         // 안정화된 중량 카운트
        private bool _writeIndicatorData = false;                   // 인디케이터 수신데이터 로그기록여부
        private bool _faxServiceIsTest = false;                      // 팩스서비스 연동환경 설정값, true - 개발용(테스트베드), false - 상업용(실서비스)
        private string _faxNumber = string.Empty;                   // 팩스발송번호
        private bool _saveImage = false;                            // 사진저장여부
        private bool _ticketPrint = false;                          // 전표출력여부
        private float _X = 0.0f;                                    // 전표출력 기준좌표 : X
        private float _Y = 0.0f;                                    //                 : Y
        private string _nvrAppPath = string.Empty;                  // NVR 영상저장 프로그램 경로
        private int _logDay = 0;                                    // 로그기록 일수

        private string _indicatorPortName = string.Empty;           // 인디케이터   - 포트명
        private string _indicatorBaudRate = string.Empty;           //             - BaudRate
        private string _indicatorDataBit = string.Empty;            //             - DataBit
        private string _indicatorStopBit = string.Empty;            //             - StopBit
        private string _indicatorParityBit = string.Empty;          //             - Parity

        private string _dbAddress = string.Empty;                   // 데이타베이스 - Address
        private string _dbName = string.Empty;                      //             - Database
        private string _dbUser = string.Empty;                      //             - User
        private string _dbPassword = string.Empty;                  //             - Password

        private string _ftpAddress = string.Empty;                  // FTP SERVER  - Address
        private string _ftpUser = string.Empty;                     //             - User
        private string _ftpPassword = string.Empty;                 //             - Password
        private string _ftpUploadPath = string.Empty;               //             - 업로드폴더 

        private string _cam1 = string.Empty;
        private string _cam2 = string.Empty;
        private string _cam3 = string.Empty;

        private string _sound = string.Empty;

        #region [ CCTV 영상 변수 ]
        Mat frame1;
        Mat frame2;
        Mat frame3;
        string RSTPaddr1 = string.Empty;
        string RSTPaddr2 = string.Empty;
        string RSTPaddr3 = string.Empty;
        #endregion
        #region [ 쓰레드 ]
        Thread t_Camera1;
        Thread t_Camera2;
        Thread t_Camera3;
        #endregion
        #region [ 녹화 및 저장 변수 ]
        VideoCapture capture1;
        VideoCapture capture2;
        VideoCapture capture3;
        #endregion
        #region [ 조건 변수 ]
        bool isConnected;
        #endregion
        #endregion
        #endregion

        
        public MainForm()
        {
            InitializeComponent();
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            if (!_firstLoad)
            {
                //SplashScreenManager.ShowForm(typeof(WaitForm1));
                //SplashScreenManager.CloseForm();
                
                pictureBoxA.BackgroundImage = Properties.Resources.A;
                Log.AddLog("프로그램이 실행되었습니다.");
                _firstLoad = true;
                InitUserControl();
                ReadSettingData();
                Initialize();
                
                // 인디게이터 시작
                Start();
                //_ftpTimer.Change(1000, System.Threading.Timeout.Infinite); //테스트용
                RTSP_Init();
                StartAll();

                // CCTV 모니터링 (참고용)
                //MONITERING();

                // 로드 시, FTP 임시폴더 (\image) 잔류 이미지 삭제
                InitFTPFile();
            }
        }
        
        #region[단계별 코드]

        #region 입출고 선택 (Page 0)
        private void Page0_Button_Click(object sender, EventArgs e)
        {
            Button bt = (Button)sender;

            // 입고
            if (bt.Name.Contains("In"))
            {
                Log.AddLog("입고 선택");
                _weighingType = WeighingType.In;
            }
            // 출고
            else if (bt.Name.Contains("Out"))
            {
                this.Focus();

                Log.AddLog("출고 선택");
                _weighingType = WeighingType.Out;
            }
            _weighingStep = WeighingStep.None;
            SelectTabPage(1);   // 계근선택
        }
        #endregion

        #region 계근 선택 (Page 1)
        private void Page1_Button_Click(object sender, EventArgs e)
        {
            Button bt = (Button)sender;
            // 1차계근
            if (bt.Name.Contains("1st"))
            {
                Log.AddLog("1차계근 선택");
                _weighingStep = WeighingStep.First;
                InitVehicleNumber();
                SelectTabPage(2);   // 차량번호입력
            }
            // 2차계근
            else if (bt.Name.Contains("2nd"))
            {
                Log.AddLog("2차계근 선택");
                _weighingStep = WeighingStep.Second;
                InitVehicleNumber();
                SelectTabPage(2);   // 차량번호입력
            }
            // 이전
            else if (bt.Name.Contains("Prev"))
            {
                this.Focus();

                Log.AddLog("이전단계이동");
                _weighingStep = WeighingStep.None;
                SelectTabPage(0);   // 입출고 선택
            }
        }
        #endregion

        #region 차량번호 입력 (Page 2)
        private void Page2_Button_Click(object sender, EventArgs e)
        {
            Button bt = (Button)sender;
            // 0 ~ 9 / ←
            if (bt.Name.Contains("Number") || bt.Name.Contains("Backspace")) { InputVehicleNumber((Button)sender); }
            // 초기화
            else if (bt.Name.Contains("Init")) { InitVehicleNumber();  }
            // 이전
            else if (bt.Name.Contains("Prev"))
            {
                Log.AddLog("이전단계이동");
                _weighingStep = WeighingStep.None;
                InitVehicleNumber();
                SelectTabPage(1);   // 계근선택
            }
            // 다음
            else if (bt.Name.Contains("Next")) { CheckValues(); }
        }

        // 다음 페이지를 위한 값 체크 메서드
        private void CheckValues()
        {
            // 0. 차량번호 입력값 검사
            _vehicleNumber = textBoxVehicleNumber.Text?.ToString();
            if (_vehicleNumber.Length < 4)
            {
                using (MessageForm form = new MessageForm())
                {
                    form.FormTitle = "차량번호오류";
                    form.Message_1 = "차량번호를 4자리 이상으로";
                    form.Message_2 = "입력하시기 바랍니다.";
                    form.ShowDialog();
                }
                return;
            }

            Log.AddLog(string.Concat("차량번호 입력 : ", _vehicleNumber));
            StringBuilder sb = new StringBuilder();
            string error = string.Empty;

            // 1. 1차계근 상황일 때,
            if (_weighingStep == WeighingStep.First)
            {
                _selectedDealerCd = string.Empty;
                _selectedDealerNm = string.Empty;
                _selectedFax = string.Empty;
                _selectedWebFaxYN = string.Empty;

                SelectTabPage(3);   // 거래처 선택
                buttonTab4Next.Visible = false;
                listBoxCustomers.Visible = false;
                textBoxCustomer.Text = string.Empty;
                textBoxCustomer.Focus();
            }
            // 2. 2차계근 상황일 때,
            if (_weighingStep == WeighingStep.Second)
            {
                // 2-1. 차량번호에 대한 1차계근 정보 조회
                string jDate = DateTime.Now.AddDays(-_secondWeighingDay).ToString("yyyy-MM-dd");

                // 2차 계근으로 접근한 경우 Mesuring 에서 계근일(J_Date), 차량번호(J_BNUM) 체크 필요
                // - 계근일자는 검사하지 않아도 된다고 함..
                sb.AppendFormat("SELECT * ");
                sb.AppendFormat("  FROM mesuring M ");
                sb.AppendFormat("  LEFT OUTER JOIN hr_emp_basis H ");
                sb.AppendFormat("       ON M.gumsu_serial = H.EMP_ID ");
                sb.AppendFormat(" WHERE J_BNum = '{0}' ", _vehicleNumber);
                sb.AppendFormat("   AND J_Date >= '{0}' ", jDate);
                if (_weighingType == WeighingType.In)
                    sb.AppendFormat("   AND KeraType = '입고' ");
                else if (_weighingType == WeighingType.Out)
                    sb.AppendFormat("   AND KeraType = '출고' ");
                sb.AppendFormat("   AND Weight = 0 ");

                object scalar = ExecuteScalar(sb.ToString(), out error);
                QuitwithLog(error);
                DataSet ds = OpenSnapshot(sb.ToString(), out error);
                QuitwithLog(error);

                if (ds != null)
                {
                    int rowCount = ds.Tables[0].Rows.Count;

                    if (rowCount == 1)
                    {
                        DataTable table = ds.Tables[0];
                        // 1차실적 정보 바인딩
                        _selectedJunpyoID = table.Rows[0]["JunpyoID"].ToString();                              // 전표아이디
                        _selectedJ_Date = Convert.ToDateTime(table.Rows[0]["J_Date"]).ToString("yyyy-MM-dd");  // 일자
                        _selectedSun = table.Rows[0]["Sun"].ToString();                                        // 계근번호(순번)
                        _selectedMaipCherID = table.Rows[0]["MaipCherID"].ToString();                          // 매입처코드
                        _selectedMaipCher = table.Rows[0]["MaipCher"].ToString();                              // 매입처
                        _selectedJ_AssignID = table.Rows[0]["J_AssignID"].ToString();                          // 매출처코드
                        _selectedJ_Company = table.Rows[0]["J_Company"].ToString();                            // 매출처

                        //2020-10-29 임시 수정
                        string sFirstTime = table.Rows[0]["FirstTime"]?.ToString().Trim();
                        if (!string.IsNullOrEmpty(sFirstTime))
                            _selectedFirstTime = Convert.ToDateTime(table.Rows[0]["FirstTime"]).ToString("yyyy-MM-dd HH:mm:ss");   // 1차계량 (입/출고에 따라 변경) - 시간
                        else
                            _selectedFirstTime = sFirstTime;
                        _selectedFirstWeight = Convert.ToInt32(table.Rows[0]["FirstWeight"]);                                  //                             - 중량 

                        string sSecondTime = table.Rows[0]["SecondTime"]?.ToString().Trim();
                        if (!string.IsNullOrEmpty(sSecondTime))
                            _selectedSecondTime = Convert.ToDateTime(table.Rows[0]["SecondTime"]).ToString("yyyy-MM-dd HH:mm:ss");   // 1차계량 (입/출고에 따라 변경) - 시간
                        else
                            _selectedSecondTime = sSecondTime;

                        _selectedSecondWeight = Convert.ToInt32(table.Rows[0]["SecondWeight"]);                //          - 중량
                        _selectedJ_Serial = table.Rows[0]["J_Serial"].ToString();                              // 검수정보 - 검수여부 ("0" 이 아니면 검수)
                        //_selectedJ_Serial = "1";    // 테스트용 (검수완료)
                        _selectedGubun1 = table.Rows[0]["Gubun1"].ToString();                                  //          - 등급
                        _selectediChaGam = Convert.ToInt32(table.Rows[0]["iChaGam"]);                          //          - 매입감량
                        _selectedOChaGam = Convert.ToInt32(table.Rows[0]["OChaGam"]);                          //          - 매출감량
                        _selectedJ_State = table.Rows[0]["J_State"].ToString();                                //          - 감가/감량사유
                        _selectedgumsubigo = table.Rows[0]["gumsubigo"].ToString();                            //          - 검수비고
                        _selectedgumsu_serial = table.Rows[0]["gumsu_serial"].ToString();                      //          - 검수자 사원번호
                        _selectedEMP_NM = table.Rows[0]["EMP_NM"].ToString();                                  //          - 검수자명         
                    }
                    else if (rowCount == 0)
                    {
                        using (MessageForm form = new MessageForm())
                        {
                            form.FormTitle = "실적없음";

                            form.Message_1 = "입력하신 차량정보가 없습니다.";
                            form.Message_2 = "(1차계량 실적없음)";
                            form.Message_3 = "- 차량번호를 다시 입력하세요.";

                            form.ShowDialog();
                        }

                        Log.AddLog(string.Concat(_vehicleNumber, " 차량의 1차 계량실적이 존재하지 않습니다."));
                        return;
                    }
                    else
                    {
                        using (MessageForm form = new MessageForm())
                        {
                            form.FormTitle = "실적중복";

                            form.Message_1 = "차량 계량실적이 중복됩니다.";
                            form.Message_2 = string.Concat("(1차계량 실적 : ", rowCount.ToString(), "건)");
                            form.Message_3 = "- 사무실로 문의하시기 바랍니다.";

                            form.ShowDialog();
                        }

                        Log.AddLog(string.Concat(_vehicleNumber, " 차량의 1차 계량실적이 중복됩니다."));
                        return;
                    }
                }

                Log.AddLog(string.Concat("검수확인 - J_Serial : ", _selectedJ_Serial, ", Gubun1 : ", _selectedGubun1, ", gumsubigo : ", _selectedgumsubigo, ", gumsu_serial : ", _selectedgumsu_serial, ", EMP_NM : ", _selectedEMP_NM));

                // 검수 완료가 된 1차 계근데이터만 2차 계근 수행
                // -> Mesuring의 J_Serial, Gubun1, Gumsubigo, Gumsu_Serial 컬럼 조회
                if (_selectedJ_Serial == "0")
                {
                    using (MessageForm form = new MessageForm())
                    {
                        form.FormTitle = "검수필요";

                        form.Message_1 = "미검수 차량입니다.";
                        form.Message_2 = "- 검수 후에 계근하세요.";

                        form.ShowDialog();
                    }

                    Log.AddLog("미검수차량");
                    return;
                }

                // 2차계량은 거래처입력 없음.. 완료화면으로 진행..
                if (_weighingType == WeighingType.In)
                {
                    _selectedDealerCd = _selectedMaipCherID;
                    _selectedDealerNm = _selectedMaipCher;
                }
                if (_weighingType == WeighingType.Out)
                {
                    _selectedDealerCd = _selectedJ_AssignID;
                    _selectedDealerNm = _selectedJ_Company;
                }
                DataSet fax = OpenSnapshot(string.Concat("SELECT *, CONCAT(ISNULL(CHRG_RGN_NO, ''), REPLACE(FAX, '-', '')) AS FAX_NO_RESULT  FROM acc_dealer_cd WHERE DEALER_NM = '", _selectedDealerNm, "'"), out error);
                QuitwithLog(error);

                if (ds != null)
                {
                    int rowCount = fax.Tables[0].Rows.Count;

                    if (fax.Tables[0].Rows.Count > 0)
                    {
                        DataTable table = fax.Tables[0];
                        _selectedFax = table.Rows[0]["FAX_NO_RESULT"].ToString();   // 팩스번호
                        _selectedWebFaxYN = table.Rows[0]["WEB_FAX_YN"].ToString(); // 웹팩스 여부
                    }
                }

                // 계량관련 데이터 초기화
                _weightFlag = false;
                _weight = 0;
                _stableCount = 0;

                // 인디케이터 중량안정화 false
                indicator.Stable = false;

                // 계량진행 플레그 SET
                _useWeighing = true;
                //indicator.Weight = 5000; //테스트용
                SelectTabPage(4);   // 중량확정 & 계량완료

                buttonTab5Complete.Visible = true;
                buttonTab5Complete.BackColor = Color.LightGray;
                buttonTab5Complete.Enabled = true;

                //buttonTab5Complete.Enabled = true;
                labelWeighingInfo.Text = string.Concat(_vehicleNumber, " 차량중량 확인 후 완료");
                labelCustomerInfo.Text = _selectedDealerNm;
            }
        }
        #endregion

        #region 거래처 입력 (Page 3)
        private void Page3_Button_Click(object sender, EventArgs e)
        {
            Button bt = (Button)sender;
            // 검색
            if (bt.Name.Contains("Search")) { CvSearch(); }
            // 초기화
            else if (bt.Name.Contains("Clear"))
            {
                listBoxCustomers.Items.Clear();
                listBoxCustomers.Visible = false;
                buttonTab4Next.Visible = false;
                textBoxCustomer.Text = string.Empty;
                textBoxCustomer.Focus();
            }
            // 이전
            else if (bt.Name.Contains("Prev"))
            {
                Log.AddLog("이전단계이동");
                if (_T_CONFIRM != null && _T_CONFIRM.IsAlive)
                    //_T_CONFIRM.Abort();
                    is_AutoComplete = false;
                SelectTabPage(2);   // 차량번호 입력
            }
            // 다음
            else if (bt.Name.Contains("Next")) { InitValues(); }
        }

        // 거래처 검색 메서드 (1차 계량만 수행)
        private void CvSearch()
        {
            string keyword = textBoxCustomer.Text.Trim();
            if (string.IsNullOrEmpty(keyword))
                return;

            StringBuilder sb = new StringBuilder();
            string error = string.Empty;

            sb.AppendFormat("SELECT * ");
            sb.AppendFormat("  FROM acc_dealer_cd ");
            sb.AppendFormat(" WHERE (DEALER_GB = '매입' OR DEALER_GB = '매출' OR DEALER_GB = '입출') ");
            sb.AppendFormat("   AND (DEALER_NM LIKE '%{0}%' OR INITIAL_NM LIKE '%{1}%') ", keyword, keyword);
            sb.AppendFormat("   AND EOB_YN <> 'Y'");

            DataSet ds = OpenSnapshot(sb.ToString(), out error);
            QuitwithLog(error);

            if (ds != null)
            {
                int rowCount = ds.Tables[0].Rows.Count;
                if (rowCount > 0)
                {
                    listBoxCustomers.Items.Clear();
                    listBoxCustomers.Visible = true;
                    DataTable table = ds.Tables[0];
                    DataRowCollection row = table.Rows;

                    foreach (DataRow dr in row)
                        listBoxCustomers.Items.Add(dr["DEALER_NM"].ToString());
                }
            }
            textBoxCustomer.Text = string.Empty;
            textBoxCustomer.Focus();
        }

        // 거래처 선택 시, 정보 바인딩 이벤트
        private void listBoxCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 리스트박스 빈칸 선택하면 오류떠서 추가 2022-06-24 정은영
            if (listBoxCustomers.SelectedItem == null)
                return;

            // 리스트박스에서 거래처를 선택함
            string keyword = listBoxCustomers.SelectedItem.ToString();
            if (string.IsNullOrEmpty(keyword))
                return;

            listBoxCustomers.Visible = false;

            StringBuilder sb = new StringBuilder();
            string error = string.Empty;

            sb.AppendFormat("SELECT * ");
            sb.AppendFormat("     , CONCAT(ISNULL(CHRG_RGN_NO, ''), REPLACE(FAX, '-', '')) AS FAX_NO_RESULT  "); //2021-01-07 추가 (지역번호 + 팩스번호)
            sb.AppendFormat("  FROM acc_dealer_cd ");
            sb.AppendFormat(" WHERE (DEALER_GB = '매입' OR DEALER_GB = '매출') ");
            sb.AppendFormat("   AND DEALER_NM = '{0}' ", keyword);

            DataSet ds = OpenSnapshot(sb.ToString(), out error);
            if (!string.IsNullOrEmpty(error))
            {
                Log.AddLog(error);
                MessageBox.Show(new Form { TopMost = true }, error);
                return;
            }

            // 거래처관련 변수들..
            _selectedDealerCd = string.Empty;
            _selectedDealerNm = string.Empty;
            _selectedFax = string.Empty;
            _selectedWebFaxYN = string.Empty;

            if (ds != null)
            {
                int rowCount = ds.Tables[0].Rows.Count;

                if (rowCount > 0)
                {
                    DataTable table = ds.Tables[0];
                    _selectedDealerCd = table.Rows[0]["DEALER_CD"].ToString();
                    _selectedDealerNm = table.Rows[0]["DEALER_NM"].ToString();
                    _selectedFax = table.Rows[0]["FAX_NO_RESULT"].ToString();
                    _selectedWebFaxYN = table.Rows[0]["WEB_FAX_YN"].ToString();

                    Log.AddLog(string.Concat("거래처 선택결과 : ", _selectedDealerNm, "(", _selectedDealerCd, ")"));
                    textBoxCustomer.Text = _selectedDealerNm;
                    textBoxCustomer.Focus();
                    buttonTab4Next.Visible = true;

                    return;
                }
            }

            textBoxCustomer.Text = "";
            textBoxCustomer.Focus();
            buttonTab4Next.Visible = false;
        }

        // 다음 페이지 관련 변수 초기화 메서드
        private void InitValues()
        {
            Log.AddLog("거래처입력");

            // 계량관련 데이터 초기화
            _weightFlag = false;
            _weight = 0;
            _stableCount = 0;

            // 인디케이터 중령안정화 false
            indicator.Stable = false;

            // 계량진행 플레그 SET
            _useWeighing = true;

            SelectTabPage(4);   // 중량확정 & 계량완료
            //indicator.Weight = 10000; //테스트용
            labelWeighingInfo.Text = string.Concat(_vehicleNumber, " 차량중량 확인 후 완료");
            labelCustomerInfo.Text = _selectedDealerNm;



            buttonTab5Complete.Visible = false;
            buttonTab5Complete.BackColor = Color.LightGray;
            buttonTab5Complete.Enabled = false;
            
            //buttonTab5Complete.Visible = true;   //테스트용
            //buttonTab5Complete.BackColor = Color.Orange;
            //buttonTab5Complete.Enabled = true;
        }
        #endregion

        #region 중량확정 & 계량완료 (Page 4)
        private void Page4_Button_Click(object sender, EventArgs e)
        {
            Button bt = (Button)sender;
            // 완료
            if (bt.Name.Contains("Complete")) { SaveProdf(); }
            // 이전
            else if (bt.Name.Contains("Prev"))
            {
                Log.AddLog("이전단계이동");

                _confirm = ConfirmYN.NonConfirm;
                _state = WeighingState.None;
                _useWeighing = false;
                if (_T_CONFIRM != null && _T_CONFIRM.IsAlive)
                    //_T_CONFIRM.Abort();
                    is_AutoComplete = false;

                // 1차 계근
                if (_weighingStep == WeighingStep.First)
                {
                    _selectedDealerCd = string.Empty;
                    _selectedDealerNm = string.Empty;
                    _selectedFax = string.Empty;
                    _selectedWebFaxYN = string.Empty;

                    SelectTabPage(3);   // 거래처 입력
                    buttonTab4Next.Visible = false;
                    textBoxCustomer.Focus();
                }

                // 2차 계근
                if (_weighingStep == WeighingStep.Second)
                    SelectTabPage(2);   // 차량번호 입력
            }
        }

        // 계량 완료 후, 데이터 저장 메서드
        private void SaveProdf()
        {
            //2020-12-10 추가
            //자동등록 방지를 위하여 Confirm처리
            _confirm = ConfirmYN.Confirm;

            // 완료처리
            _state = WeighingState.Complete;
            Log.AddLog("계량상태 변경 : Stable -> Complete");

            // 버튼 비활성화
            SetCompleteButton2(false);
            SetPrvButton2(false);

            string imageFile = string.Empty;
            string appPath = Application.StartupPath;
            DateTime now = DateTime.Now;
            StringBuilder sb = new StringBuilder();
            int weight = 0;
            int iWeight = 0;
            int oWeight = 0;
            int nRet = 0;
            string error = string.Empty;

            try
            {
                #region 1차계량
                if (_weighingStep == WeighingStep.First)
                {
                    string sJunpyoID = string.Empty;
                    string sJ_DATE = string.Empty;
                    // 1. 1차 계량실적 저장
                    // - 순번을 구한다
                    object scalar = ExecuteScalar(string.Concat("SELECT MAX(Sun) FROM mesuring WHERE J_Date = '", now.ToString("yyyy-MM-dd"), "'"), out error);
                    QuitwithLog_Page4(error);

                    // 순번 세팅 (일자 기준)
                    int sun = 0;
                    if (string.IsNullOrEmpty(scalar.ToString())) { sun = 1; }   // 빈 값이면 1부터 시작
                    else { sun = Convert.ToInt32(scalar) + 1; }                 // 값이 있다면, +1

                    // 1. 실적등록 (입고 : 2차실적 컬럼 / 출고 1차실적 컬럼에 값넣기)
                    /*
                     * 2021-02-22
                     * MESURING에 업체담당자 컬럼 추가됨에 따라 아래 쿼리에 해당 컬럼에 INSERT
                     */
                    sb.AppendFormat("INSERT INTO mesuring ");
                    sb.AppendFormat("       ( ");
                    sb.AppendFormat("            KeraType, MaipCherID, MaipCher, J_AssignID, J_Company, ");
                    sb.AppendFormat("            Sun, J_Date, FirstTime, SecondTime, FirstWeight, SecondWeight, Weight, ");
                    sb.AppendFormat("            J_Serial, J_BNum, K_Name, ");
                    sb.AppendFormat("            iWeight, OWeight, P_ID, J_Garage, customweight, ");
                    sb.AppendFormat("            J_Check, MesuNo, Gubun1, Jajae, ");                // 기존테이블에 데이터 맞춤
                    sb.AppendFormat("            FUserCode, UserCode, FBuseoCode, BuseoCode, ");    // 기존테이블에 데이터 맞춤
                    sb.AppendFormat("            U_Date, J_State, hms1, hms2, gumsubigo, kyeryang12 ");         // 기존테이블에 데이터 맞춤
                    //sb.AppendFormat("            CHRG_ID "); //CHRG_ID 추가
                    sb.AppendFormat("       ) ");
                    sb.AppendFormat("VALUES ");
                    sb.AppendFormat("       ( ");
                    if (_weighingType == WeighingType.In)
                    {
                        sb.AppendFormat("            '{0}', ", "입고");                             // KeraType
                        sb.AppendFormat("            {0}, ", _selectedDealerCd);                    // MaipCherID
                        sb.AppendFormat("            '{0}', ", _selectedDealerNm);                  // MaipCher
                        sb.AppendFormat("            {0}, ", 0);                                    // J_AssignID
                        sb.AppendFormat("            '{0}', ", "");                                 // J_Company
                    }
                    else if (_weighingType == WeighingType.Out)
                    {
                        sb.AppendFormat("            '{0}', ", "출고");                             // KeraType
                        sb.AppendFormat("            {0}, ", 0);                                    // MaipCherID
                        sb.AppendFormat("            '{0}', ", "");                                 // MaipCher
                        sb.AppendFormat("            {0}, ", _selectedDealerCd);                    // J_AssignID
                        sb.AppendFormat("            '{0}', ", _selectedDealerNm);                  // J_Company
                    }
                    sb.AppendFormat("            {0}, ", sun);                                      // Sun
                    sb.AppendFormat("            '{0}', ", now.ToString("yyyy-MM-dd"));             // J_Date
                    sb.AppendFormat("            '{0}', ", now.ToString("yyyy-MM-dd HH:mm:ss"));    // FirstTime
                    sb.AppendFormat("            '{0}', ", now.ToString("yyyy-MM-dd HH:mm:ss"));    // SecondTime
                    if (_weighingType == WeighingType.In)
                    {
                        sb.AppendFormat("            {0}, ", 0);                                    // FirstWeight
                        sb.AppendFormat("            {0}, ", _weight);                              // SecondWeight
                    }
                    else if (_weighingType == WeighingType.Out)
                    {
                        sb.AppendFormat("            {0}, ", _weight);                              // FirstWeight
                        sb.AppendFormat("            {0}, ", 0);                                    // SecondWeight
                    }
                    sb.AppendFormat("            {0}, ", 0);                                        // Weight
                    sb.AppendFormat("            {0}, ", 0);                                        // J_Serial
                    sb.AppendFormat("            '{0}', ", _vehicleNumber);                         // J_BNum
                    if (_weighingType == WeighingType.In)
                    { sb.AppendFormat("            '{0}', ", "원재료"); }                           // K_Name
                    else if (_weighingType == WeighingType.Out)
                    { sb.AppendFormat("            '{0}', ", "상품"); }                             // K_Name
                    sb.AppendFormat("            {0}, ", 0);                                        // iWeight
                    sb.AppendFormat("            {0}, ", 0);                                        // OWeight
                    sb.AppendFormat("            {0}, ", 100);                                      // P_ID
                    sb.AppendFormat("            {0}, ", 1);                                        // J_Garage ?
                    sb.AppendFormat("            {0}, ", 0);                                        // customweight
                    sb.AppendFormat("            '{0}', ", "");                                     // J_Check
                    sb.AppendFormat("            '{0}', ", "");                                     // MesuNo
                    sb.AppendFormat("            '{0}', ", "");                                     // Gubun1
                    sb.AppendFormat("            '{0}', ", "");                                     // Jajae
                    sb.AppendFormat("            {0}, ", 0);                                        // FUserCode
                    sb.AppendFormat("            {0}, ", 0);                                        // UserCode
                    sb.AppendFormat("            {0}, ", 0);                                        // FBuseoCode
                    sb.AppendFormat("            {0}, ", 0);                                        // BuseoCode
                    sb.AppendFormat("            '{0}', ", "0000-00-00");                           // U_Date
                    sb.AppendFormat("            '{0}', ", "");                                     // J_State
                    sb.AppendFormat("            {0}, ", 0);                                        // hms1
                    sb.AppendFormat("            {0}, ", 0);                                        // hms2
                    sb.AppendFormat("            '{0}', ", "");                                      // gumsubigo
                    sb.AppendFormat("            {0}", 1);                  //2021-01-12 계량구분 1차 시 1, 2차시 2
                    //sb.AppendFormat("            (SELECT CHRG_ID FROM ACC_DEALER_CD WHERE DEALER_CD = '{0}')", _selectedDealerCd); // CHRG_ID
                    sb.AppendFormat("       ) ");

                    nRet = ExecuteNonQuery(sb.ToString(), out error);
                    if (nRet == 1) { Log.AddLog(string.Concat("1차 계량완료 - ", _weight.ToString(), "kg")); }
                    else { QuitwithLog_Page4(error); }

                    // 이미지 저장을 위하여 전표번호조회
                    sb.Clear();
                    sb.AppendFormat(" SELECT A.JUNPYOID ");
                    sb.AppendFormat("   FROM MESURING A  ");
                    sb.AppendFormat("  WHERE A.SUN = '{0}' ", sun);
                    if (_weighingType == WeighingType.In)
                    { sb.AppendFormat("    AND A.KERATYPE = '입고' "); }
                    else if (_weighingType == WeighingType.Out)
                    { sb.AppendFormat("    AND A.KERATYPE = '출고' "); }
                    sb.AppendFormat("    AND A.J_DATE = '{0}'  ", now.ToString("yyyy-MM-dd"));
                    if (_weighingType == WeighingType.In)
                    { sb.AppendFormat("    AND A.MAIPCHERID = {0} ", _selectedDealerCd); }
                    else if (_weighingType == WeighingType.Out)
                    { sb.AppendFormat("    AND A.J_ASSIGNID = {0} ", _selectedDealerCd); }
                    sb.AppendFormat("    AND A.J_BNUM = '{0}' ", _vehicleNumber);
                    if (_weighingType == WeighingType.In)
                    { sb.AppendFormat("   AND A.SECONDWEIGHT = {0}; ", _weight); }
                    else if (_weighingType == WeighingType.Out)
                    { sb.AppendFormat("   AND A.FIRSTWEIGHT = {0}; ", _weight); }

                    scalar = ExecuteScalar(sb.ToString(), out error);
                    QuitwithLog_Page4(error);

                    sJunpyoID = scalar.ToString();

                    // 이미지 저장을 위하여 날짜조회
                    sb.Clear();
                    sb.AppendFormat(" SELECT A.J_DATE  ");
                    sb.AppendFormat("   FROM MESURING A ");
                    sb.AppendFormat("  WHERE A.JUNPYOID = {0} ", sJunpyoID);

                    scalar = ExecuteScalar(sb.ToString(), out error);
                    QuitwithLog_Page4(error);

                    sJ_DATE = scalar.ToString();

                    // 2. 1차실적 차량이미지 저장
                    if (_saveImage)
                    {
                        // 2-1. 이미지 저장경로 지정
                        string saveFolder = string.Concat(appPath, @"\image\");
                        // 2-2. 이미지 저장경로 확인 및 생성
                        if (!System.IO.Directory.Exists(saveFolder))
                            System.IO.Directory.CreateDirectory(saveFolder);
                        // 2-3. RTSP 이미지 저장(디버그 image폴더에 저장)
                        RTSP_IMAGE_SAVE(saveFolder, sJunpyoID, "1"); 

                        // 3. 2-3번 과정 성공 시, FTP 파일 전송
                        // 3-1. 파일이 생성되어 있는지 검사
                        int cnt = 0;
                        bool result = false;
                        string file_1 = string.Concat(saveFolder, sJunpyoID, "_1_1.jpg");     // 1차 - 1번
                        string file_2 = string.Concat(saveFolder, sJunpyoID, "_1_2.jpg");     //     - 2번
                        string file_3 = string.Concat(saveFolder, sJunpyoID, "_1_3.jpg");     //     - 3번
                        while (cnt < 13)
                        {
                            cnt++;
                            System.Threading.Thread.Sleep(500);

                            if (File.Exists(file_1) && File.Exists(file_2) && File.Exists(file_3))
                            {
                                result = true;
                                break;
                            }
                        }

                        // 3-2. 이미지 복사경로 지정
                        string uploadFolder = string.Concat(appPath, @"\image\upload\");
                        // 3-3. 이미지 복사경로 확인 및 생성
                        if (!System.IO.Directory.Exists(uploadFolder))
                            System.IO.Directory.CreateDirectory(uploadFolder);

                        // 3-4. 2-3번에서 생성된 이미지 경로 담기
                        string[] file = new string[3] { file_1, file_2, file_3 };
                        //// 3-5. 디렉토리 이미지 복사(MES에 보여주기위해 image/upload 파일에 복사)
                        RTSP_IMAGE_COPY(uploadFolder, sJunpyoID, sJ_DATE, "1", file);

                        // 4. FTP 업로드 (1차 계근 이미지)
                        //FTP_FILE_UPLOAD(file);
                    }
                }
                #endregion

                #region 2차계량
                if (_weighingStep == WeighingStep.Second)
                {
                    // 2회 계량실적 중량이 같을 경우
                    if (_selectedFirstWeight == _weight || _selectedSecondWeight == _weight)
                    {
                        using (MessageForm form = new MessageForm())  //테스트용(실제사용시 주석해제)
                        {
                            form.FormTitle = "계량오류";
                        
                            form.Message_1 = "계량 중량이 잘못되었습니다.";
                            form.Message_2 = "(1,2차 계량중량이 동일합니다)";
                            form.Message_3 = "- 사무실로 문의하시기 바랍니다.";
                        
                            if (form.ShowDialog() == DialogResult.Cancel)
                            {
                                //SelectTabPage(3);   //메세지 확인 이후 거래처선택 으로 이동
                                _state = WeighingState.Stable;
                                _useWeighing = false;
                                InitWeighingValue();
                                //InitVehicleNumber();
                                SetCompleteButton2(true);
                                SetPrvButton2(true);
                                return;
                            }
                        }
                    }

                    // 1. 1차 계량실적에 2차정보 기록
                    // - 입고실적 (1차실적 컬럼에 값을 넣는다)
                    if (_weighingType == WeighingType.In)
                    {
                        /*
                         * 2021-01-11 (현업요청)
                         * 2차계근 시 1차계근(공차중량) 값과 2차계근값(대지중량)의 값이 뒤바뀌는 경우가 존재
                         * 마지막 저장시 계근값은 공차중량과 대지중량을 비교하여 무조건 큰것이 대지중량, 작은것이 공차중량으로 바꾸어 수정
                         * 5000 / 10000
                         */
                        if (_selectedSecondWeight < _weight)
                        {
                            int iTemp = _selectedSecondWeight;
                            _selectedSecondWeight = _weight;
                            _weight = iTemp;
                        }

                        weight = _selectedSecondWeight - _weight;
                        iWeight = weight - _selectediChaGam;
                    }
                    // - 출고실적 (정상적으로 2차실적 컬럼에 값을 넣는다)
                    else if (_weighingType == WeighingType.Out)
                    {
                        /*
                         * 2021-01-11(현업요청)
                         * 출고 시 1차계근(공차중량)값이 2차계근(대지중량)값 보다 큰 경우가 발생한다고함
                         * 따라서 아래 계량순서 오류는 주석처리
                         * 출고는 계근값 체크하지 않는 것으로 수정하고
                         * 공차중량이 대지중량보다 큰경우에는 아래 로직에 따라
                         * 공자중량에 대지중량을 세팅 / 대지중량에 공차중량 세팅하여
                         * 쿼리날림
                         */
                        

                        /*
                         * 2021-01-11 (현업요청)
                         * 2차계근 시 1차계근(공차중량) 값과 2차계근값(대지중량)의 값이 뒤바뀌는 경우가 존재
                         * 마지막 저장시 계근값은 공차중량과 대지중량을 비교하여 무조건 큰것이 대지중량, 작은것이 공차중량으로 바꾸어 수정
                         * 예) 1차계근(공차중량) : 3000 / 2차계근(대지중량) : 9000 -> 정상수행
                         * 1차계근(공차중량) : 9000 / 2차계근(대지중량) : 3000 -> 두 값을 뒤바꾸어 자동으로 수행(아래 if문 참고)
                         */
                        if (_selectedFirstWeight > _weight)
                        {
                            int iTemp = _selectedFirstWeight;
                            _selectedFirstWeight = _weight;
                            _weight = iTemp;
                        }

                        weight = _weight - _selectedFirstWeight;
                        oWeight = weight - _selectedOChaGam;
                    }

                    /*
                     * 2021-02-22
                     * MESURING에 업체담당자 컬럼 추가됨에 따라 아래 쿼리에 해당 컬럼에 INSERT
                     */
                    sb.AppendFormat("UPDATE mesuring ");
                    if (_weighingType == WeighingType.In)
                    {
                        sb.AppendFormat("   SET FirstTime = '{0}', ", now.ToString("yyyy-MM-dd HH:mm:ss"));
                        sb.AppendFormat("       FirstWeight = {0}, ", _weight);
                        sb.AppendFormat("       SecondWeight = {0}, ", _selectedSecondWeight);
                        sb.AppendFormat("       iWeight = {0}, ", iWeight);
                    }
                    else if (_weighingType == WeighingType.Out)
                    {
                        sb.AppendFormat("   SET SecondTime = '{0}', ", now.ToString("yyyy-MM-dd HH:mm:ss"));
                        sb.AppendFormat("       FirstWeight = {0}, ", _selectedFirstWeight);
                        sb.AppendFormat("       SecondWeight = {0}, ", _weight);
                        sb.AppendFormat("       OWeight = {0}, ", oWeight);
                    }
                    sb.AppendFormat("       Weight = {0}, ", weight);
                    sb.AppendFormat("       customweight = {0}, ", weight);
                    sb.AppendFormat("       kyeryang12 = {0} ", 2);
                    //sb.AppendFormat("       CHRG_ID = ( SELECT CHRG_ID FROM ACC_DEALER_CD WHERE DEALER_CD = {0} )", _selectedDealerCd);
                    sb.AppendFormat(" WHERE JunpyoID = {0} ", _selectedJunpyoID);
                    if (_weighingType == WeighingType.In)
                    {
                        // 전표작성을 위한 값들을 기록
                        _selectedFirstTime = now.ToString("yyyy-MM-dd HH:mm:ss");
                        _selectedFirstWeight = _weight;
                    }
                    else if (_weighingType == WeighingType.Out)
                    {
                        // 전표작성을 위한 값들을 기록
                        _selectedSecondTime = now.ToString("yyyy-MM-dd HH:mm:ss");
                        _selectedSecondWeight = _weight;
                    }
                    
                    nRet = ExecuteNonQuery(sb.ToString(), out error);
                    if (nRet == 1) { Log.AddLog(string.Concat("2차 계량완료 - ", _weight.ToString(), "kg")); }
                    else { QuitwithLog_Page4(error); }

                    // 2. 2차실적 차량이미지 저장
                    if (_saveImage)
                    {
                        imageFile = string.Concat
                        (
                            "!",
                            string.Concat(Application.StartupPath, @"\image\", _selectedJunpyoID, "_2_").PadRight(126),
                            //string.Concat(@"D:\", _vehicleNumber, "_2_").PadRight(126),
                            "@"
                        );

                        // 2-1. 이미지 저장경로 지정
                        string saveFolder = string.Concat(appPath, @"\image\");
                        // 2-2. 이미지 저장경로 확인 및 생성
                        if (!System.IO.Directory.Exists(saveFolder))
                            System.IO.Directory.CreateDirectory(saveFolder);

                        RTSP_IMAGE_SAVE(saveFolder, _selectedJunpyoID, "2");

                        // 3. 2-3번 과정 성공 시, FTP 파일 전송
                        // 3-1. 파일이 생성되어 있는지 검사
                        int cnt = 0;
                        bool result = false;
                        string file_1 = string.Concat(appPath, @"\image\", _selectedJunpyoID, "_2_1.jpg");     // 2차 - 1번
                        string file_2 = string.Concat(appPath, @"\image\", _selectedJunpyoID, "_2_2.jpg");     //     - 2번
                        string file_3 = string.Concat(appPath, @"\image\", _selectedJunpyoID, "_2_3.jpg");     //     - 3번
                        while (cnt < 16)
                        {
                            cnt++;
                            System.Threading.Thread.Sleep(500);

                            if (File.Exists(file_1) && File.Exists(file_2) && File.Exists(file_3))
                            {
                                result = true;
                                break;
                            }
                        }
                        // 3-2. 이미지 복사경로 지정
                        string deskFolder = string.Concat(appPath, @"\image\upload\");
                        // 3-3. 이미지 복사경로 확인 및 생성
                        if (!System.IO.Directory.Exists(deskFolder))
                            System.IO.Directory.CreateDirectory(deskFolder);
                        // 3-4. 2-3번에서 생성된 이미지 경로 담기
                        string[] file = new string[3] { file_1, file_2, file_3 };
                        // 3-5. 디렉토리 이미지 복사
                        RTSP_IMAGE_COPY(deskFolder, _selectedJunpyoID, _selectedJ_Date, "2", file);

                        // 4. FTP 업로드 (2차 계근 이미지)
                        //FTP_FILE_UPLOAD(file);
                    }
                    
                    // 4. 이미지 파일관련 정리
                    // 4-1. 이미지파일 경로 할당
                    _selectedTicketImage_1_1 = string.Concat(appPath, @"\image\", _selectedJunpyoID, "_1_1.jpg");
                    _selectedTicketImage_1_2 = string.Concat(appPath, @"\image\", _selectedJunpyoID, "_1_2.jpg");
                    _selectedTicketImage_1_3 = string.Concat(appPath, @"\image\", _selectedJunpyoID, "_1_3.jpg");
                    _selectedTicketImage_2_1 = string.Concat(appPath, @"\image\", _selectedJunpyoID, "_2_1.jpg");
                    _selectedTicketImage_2_2 = string.Concat(appPath, @"\image\", _selectedJunpyoID, "_2_2.jpg");
                    _selectedTicketImage_2_3 = string.Concat(appPath, @"\image\", _selectedJunpyoID, "_2_3.jpg");
                    // 4-2. 전표출력용 임시파일 경로 할당
                    string tmpImage_1_1 = string.Concat(appPath, @"\image\print\1_1.jpg");
                    string tmpImage_1_2 = string.Concat(appPath, @"\image\print\1_2.jpg");
                    string tmpImage_1_3 = string.Concat(appPath, @"\image\print\1_3.jpg");
                    string tmpImage_2_1 = string.Concat(appPath, @"\image\print\2_1.jpg");
                    string tmpImage_2_2 = string.Concat(appPath, @"\image\print\2_2.jpg");
                    string tmpImage_2_3 = string.Concat(appPath, @"\image\print\2_3.jpg");
                    // 4-3. 이전 전표출력용 파일이 있을 경우 삭제한다
                    if (File.Exists(tmpImage_1_1)) File.Delete(tmpImage_1_1);
                    if (File.Exists(tmpImage_1_2)) File.Delete(tmpImage_1_2);
                    if (File.Exists(tmpImage_1_3)) File.Delete(tmpImage_1_3);
                    if (File.Exists(tmpImage_2_1)) File.Delete(tmpImage_2_1);
                    if (File.Exists(tmpImage_2_2)) File.Delete(tmpImage_2_2);
                    if (File.Exists(tmpImage_2_3)) File.Delete(tmpImage_2_3);
                    // 4-4. 출력용 임시파일 복사
                    if (File.Exists(_selectedTicketImage_1_1)) File.Copy(_selectedTicketImage_1_1, tmpImage_1_1);
                    if (File.Exists(_selectedTicketImage_1_2)) File.Copy(_selectedTicketImage_1_2, tmpImage_1_2);
                    if (File.Exists(_selectedTicketImage_1_3)) File.Copy(_selectedTicketImage_1_3, tmpImage_1_3);
                    if (File.Exists(_selectedTicketImage_2_1)) File.Copy(_selectedTicketImage_2_1, tmpImage_2_1);
                    if (File.Exists(_selectedTicketImage_2_2)) File.Copy(_selectedTicketImage_2_2, tmpImage_2_2);
                    if (File.Exists(_selectedTicketImage_2_3)) File.Copy(_selectedTicketImage_2_3, tmpImage_2_3);
                    // 4-5. 출력 이미지 바인딩 (이미지가 없으면 No Image 표시)
                    BINDING_IMAGE(tmpImage_1_1, pictureBoxIn1);
                    BINDING_IMAGE(tmpImage_1_2, pictureBoxIn2);
                    BINDING_IMAGE(tmpImage_1_3, pictureBoxIn3);
                    BINDING_IMAGE(tmpImage_2_1, pictureBoxOut1);
                    BINDING_IMAGE(tmpImage_2_2, pictureBoxOut2);
                    BINDING_IMAGE(tmpImage_2_3, pictureBoxOut3);
                    File.Delete(_selectedTicketImage_1_1);
                    File.Delete(_selectedTicketImage_1_2);
                    File.Delete(_selectedTicketImage_1_3);
                    File.Delete(_selectedTicketImage_2_1);
                    File.Delete(_selectedTicketImage_2_2);
                    File.Delete(_selectedTicketImage_2_3);
                    /*
                     *  2021-01-28 (현업요청)
                     *  전표출력은 모든 로직 처리 후 마지막에 출력하도록 수정
                     *  
                     * 2021-02-01
                     *  전표출력 로직순서 변경 후 이미지가 정상적으로 하나씩 빠지는 현상으로
                     * 다시 원상복구
                     */
                    // 5. 전표(계량증명서) 출력
                    if (_ticketPrint)                                                       //테스트용(테스트 시 전표, 팩스 주석처리)
                    {
                        using (PrintForm form = new PrintForm())
                        {
                            DialogResult result = form.ShowDialog();
                            if (result == DialogResult.Yes)
                            {
                                Log.AddLog("계량전표출력");
                                printDocument.Print();
                            }
                        }
                    }
                    
                    // 6. WebFax 전송대상일 경우 해당 전표이미지 발송
                    if (_selectedWebFaxYN.Trim().ToUpper() == "Y")
                    {
                        string weighingType = string.Empty;
                        string fullWeight = string.Empty;
                        string emptyWeight = string.Empty;
                        int netWeight = 0;
                        int chaGam = 0;

                        if (_weighingType == WeighingType.In)
                        {
                            weighingType = "입고";
                            //fullWeight = string.Concat((_selectedFirstWeight.ToString() + "KG").PadRight(15), _selectedFirstTime.Substring(11, 5));
                            //emptyWeight = string.Concat((_selectedSecondWeight.ToString() + "KG").PadRight(15), _selectedSecondTime.Substring(11, 5));

                            string secTime = string.Empty;
                            string firTime = string.Empty;

                            if (_selectedSecondTime.Length > 11)
                            {
                                secTime = _selectedSecondTime.Substring(11, 5);
                            }

                            if (_selectedFirstTime.Length > 11)
                            {
                                firTime = _selectedFirstTime.Substring(11, 5);
                            }

                            fullWeight = string.Concat((_selectedSecondWeight.ToString() + "KG").PadRight(15), secTime);
                            emptyWeight = string.Concat((_selectedFirstWeight.ToString() + "KG").PadRight(15), firTime);
                            netWeight = _selectedSecondWeight - _selectedFirstWeight - _selectediChaGam;
                            chaGam = _selectediChaGam;
                        }
                        if (_weighingType == WeighingType.Out)
                        {
                            string secTime = string.Empty;
                            string firTime = string.Empty;

                            if (_selectedSecondTime.Length > 11)
                            {
                                secTime = _selectedSecondTime.Substring(11, 5);
                            }

                            if (_selectedFirstTime.Length > 11)
                            {
                                firTime = _selectedFirstTime.Substring(11, 5);
                            }

                            weighingType = "출고";
                            fullWeight = string.Concat((_selectedSecondWeight.ToString() + "KG").PadRight(15), secTime);
                            emptyWeight = string.Concat((_selectedFirstWeight.ToString() + "KG").PadRight(15), firTime);
                            netWeight = _selectedSecondWeight - _selectedFirstWeight - _selectedOChaGam;
                            chaGam = _selectedOChaGam;
                        }

                        // - 결과 텍스트 기록
                        label1.Text = _selectedDealerNm;                        // 거래처명
                        label2.Text = _selectedJ_Date;                          // 일자
                        label3.Text = weighingType;                             // 입출고
                        label4.Text = _selectedSun;                             // 계근번호
                        label5.Text = _selectedGubun1;                          // 품명 ?
                        label6.Text = _vehicleNumber;                           // 차량번호
                        label7.Text = _selectedEMP_NM;                          // 검수자
                        label8.Text = fullWeight;                               // 총중량
                        label9.Text = emptyWeight;                              // 공차중량
                        label10.Text = string.Format("{0}KG", netWeight.ToString());                    // 실중량
                        label11.Text = _selectedGubun1;                         // 등급명
                        label12.Text = string.Format("{0}KG", chaGam.ToString());                       // 감량
                        label13.Text = _selectedJ_State;                        // 감량사유
                    
                        //#003 이전위치
                    
                        SelectTabPage(5);
                        //화면이 깨지는거 처럼 이상해서 없앰 2022-06-24 정은영
                        //SelectTabPage(4);
                    
                        // - jpg 저장 
                        Bitmap bmp = new Bitmap(this.panelTicket.Width, this.panelTicket.Height);
                        this.panelTicket.DrawToBitmap(bmp, new Rectangle(0, 0, this.panelTicket.Width, this.panelTicket.Height));
                        string ticketImageFile = string.Concat(appPath, @"\ticket\", DateTime.Now.ToString("yyyyMMdd_HHmmss_"), _selectedJunpyoID, ".jpg");
                        bmp.Save(ticketImageFile, System.Drawing.Imaging.ImageFormat.Jpeg);
                        Thread.Sleep(300);
                        // - WebFax 발송
                        WebFax fax = new WebFax(_faxServiceIsTest);
                        string faxResultNumber = string.Empty;
                        bool bRet = fax.SendFax
                        (
                            _faxNumber,                                 // 발신번호
                            _selectedFax,                               // 수신번호
                            _selectedDealerNm,                          // 수신자명
                            ticketImageFile,                            // 파일명
                            "WeighingTicket",                           // 팩스제목
                            out faxResultNumber                         // out result : 접수번호
                        );
                    
                        Log.AddLog(string.Concat("팩스발송 - 발신번호 : ", _faxNumber, ", 수신번호 : ", _selectedFax, ", 수신자명 : ", _selectedDealerNm, ", 파일명 : ", ticketImageFile));
                    
                        /*
                         * #002
                         */
                        if (bRet)
                        {
                            Log.AddLog(string.Concat("팩스 접수번호 : ", faxResultNumber));
                    
                            StringBuilder strSql = new StringBuilder();
                    
                            strSql.Clear();
                            strSql.AppendFormat(" ");
                            strSql.AppendFormat(" INSERT INTO WEB_FAX_LOG ");
                            strSql.AppendFormat("           ( JUNPYOID ");
                            strSql.AppendFormat("           , WEB_FAX_NO ");
                            strSql.AppendFormat("           , FAX_SND_NO ");
                            strSql.AppendFormat("           , PGM_ID ");
                            strSql.AppendFormat("           , REG_ID ");
                            strSql.AppendFormat("           , REG_DT ) ");
                            strSql.AppendFormat("     VALUES( {0} ", _selectedJunpyoID);
                            strSql.AppendFormat("           , '{0}' ", faxResultNumber);
                            strSql.AppendFormat("           , '{0}' ", _selectedFax);
                            strSql.AppendFormat("           , 'SYS999F01' "); //HARDCODING
                            strSql.AppendFormat("           , '9999' "); //HARDCODING
                            strSql.AppendFormat("           , CONVERT(VARCHAR(19),GETDATE(),20) ) ");
                    
                            nRet = ExecuteNonQuery(strSql.ToString(), out error);
                        }
                        else
                        {
                            MessageBox.Show(new Form { TopMost = true }, faxResultNumber);
                        }
                    
                        if (nRet == 1)
                        {
                    
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(error))
                            {
                                Log.AddLog(error);
                                MessageBox.Show(new Form { TopMost = true }, error);
                            }
                        }
                    
                    }
                    
                }
                #endregion
                //using (MessageForm form = new MessageForm())
                //{
                //    form.FormTitle = "계근등록";
                //    form.Message_1 = "처리되었습니다.";
                //    form.Message_2 = "- 초기 페이지로 이동합니다.";
                //    form.ShowDialog();
                //}
                Log.AddLog(string.Concat(_vehicleNumber, " 차량의 1차 계량실적이 존재하지 않습니다."));
                // 완료 후 인디케이터 데이터를 처리하지 않음
                _useWeighing = false;
                InitWeighingValue();

                SetInitVehicleNumber(true);
                //InitVehicleNumber();
                SetCompleteButton2(true);
                SetCompleteButtonVisible(false);
                SetPrvButton2(true);

                // 전표(계량증명서) 바인딩 값 초기화
                InitTicketImage();
                SelectTabPage(0);

                compleAlarm();
                //#001 소리재생 제거 2022-12-08
                //Thread alarmThread = new Thread(new ThreadStart(compleAlarm));
            }
            catch (Exception ex)
            {
                SelectTabPage(0);
                Log.AddLog(ex.Message);
                MessageBox.Show(new Form { TopMost = true }, ex.Message);
            }
            finally
            {

            }
        }

        #region 소리재생  #001
        private void compleAlarm()
        {
            string fileName = _sound;
            playMusic(fileName);
            Thread.Sleep(10000);
            stopMusic(fileName);
        }

        //private string musicPath = Application.StartupPath + @"\music\classical-waltz-loop.wav";
        //private void playMusic()
        //{
        //    SoundPlayer sp = new SoundPlayer(musicPath);
        //    sp.Play();
        //}

        private string musicPath = Path.Combine(Application.StartupPath, "sound");

        private void playMusic(string fileName)
        {
            string filePath = Path.Combine(musicPath, fileName);
            if (File.Exists(filePath))
            {
                SoundPlayer sp = new SoundPlayer(filePath);
                sp.Play();
            }
            else
            {

            }
        }

        private void stopMusic(string fileName)
        {
            string filePath = Path.Combine(musicPath, fileName);
            if (File.Exists(filePath))
            {
                SoundPlayer sp = new SoundPlayer(filePath);
                sp.Stop();
            }
            else
            {

            }
        }
        #endregion

        // RTSP 영상 이미지 저장
        private void RTSP_IMAGE_SAVE(string path, string slino, string GB)
        {
            //frame1 = new Mat();
            //frame2 = new Mat();
            //frame3 = new Mat();
            //if (!capture1.Read(frame1)) { }
            //if (!capture2.Read(frame2)) {  }
            //if (!capture3.Read(frame3)) { }
            
            string imageFile = string.Empty;

            // 2-3. 이미지 저장
            for(int i = 1; i <= 3; i++)
            {
                /*
                Bitmap CpBitmap = null;
                if (i == 1) { CpBitmap = BitmapConverter.ToBitmap(frame1); }
                else if (i == 2) { CpBitmap = BitmapConverter.ToBitmap(frame2); }
                else if (i == 3) { CpBitmap = BitmapConverter.ToBitmap(frame3); }
                imageFile = string.Concat(path, slino, "_", GB, "_", i.ToString(), ".jpg");
                CpBitmap.Save(imageFile, ImageFormat.Jpeg);
                //Delay(800);
                */
                PictureBox pb = new PictureBox();
                Image CpBitmap = null;
                if (i == 1) { CSafeSetImageBox(pb, pictureBox14.Image); CpBitmap = pb.Image; }
                else if (i == 2) { CSafeSetImageBox(pb, pictureBox13.Image); CpBitmap = pb.Image; }
                else if (i == 3) { CSafeSetImageBox(pb, pictureBox12.Image); CpBitmap = pb.Image; }
                imageFile = string.Concat(path, slino, "_", GB, "_", i.ToString(), ".jpg");
                CpBitmap.Save(imageFile, ImageFormat.Jpeg);
            }

            // 2-4. CCTV 영상 메모리 해제 (중요★, 한번 끊었다가 스레드 재수행)
            isConnected = false;

            //capture1.Release();
            //capture2.Release();
            //capture3.Release();

            //capture1.Dispose();
            //capture2.Dispose();
            //capture3.Dispose();

            //t_Camera1?.Abort();
            //t_Camera1 = null;

            //t_Camera2?.Abort();
            //t_Camera2 = null;

            //t_Camera3?.Abort();
            //t_Camera3 = null;

            //영상쓰레드 재시작
            //StartAll();
        }
        
        // 저장된 이미지 파일 복사
        private void RTSP_IMAGE_COPY(string path, string slino, string jdate, string GB, string[] file)
        {
            string destFile = string.Empty;
            DateTime createdTime = Convert.ToDateTime(jdate);
            if (_saveImage)
            {
                for(int i = 0; i < 3; i++) 
                    {
                    string filepath = file[i];
                    destFile = string.Concat(path, slino.ToString(), "_", GB, "_", (i+1).ToString(), ".jpg");
                    if (File.Exists(filepath))
                    {
                        File.Copy(filepath, destFile);
                        File.SetCreationTime(destFile, createdTime);
                    }
                }
            }
        }
        
        // FTP 파일 업로드
        private void FTP_FILE_UPLOAD(string[] file)
        {
            try
            {
                foreach (string filepath in file)
                {
                    string tempname = filepath.Substring(filepath.LastIndexOf('\\'));
                    string filename = tempname.Replace("\\", "");

                    Image image = Image.FromFile(filepath);
                    byte[] temp = ImageToByteArray(image);

                    if (temp != null)
                    {
                        var ftpUpfile = new WebClient();
                        ftpUpfile.Credentials = new NetworkCredential(_ftpUser, _ftpPassword);
                        ftpUpfile.UploadFile(string.Concat("ftp://" + _ftpAddress + "/" + _ftpUploadPath + "/" + filename), filepath);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(new Form { TopMost = true }, ex.Message);
                return;
            }
        }

        // 1차/2차 계근 이미지 출력물에 바인딩
        private void BINDING_IMAGE(string image, PictureBox pb)
        {
            if (File.Exists(image))
            {
                StreamReader sr = new StreamReader(image);
                pb.BackgroundImage = Image.FromStream(sr.BaseStream);
                sr.Dispose();
            }
            //2020-12-08 계근이미지가 존재하지 않을 경우 이미지 없음으로 출력
            else
                pb.BackgroundImage = Properties.Resources.No_Img;
        }

        private void buttonTab5Prev_Click(object sender, EventArgs e)
        {
            // 이전단계로..
            Log.AddLog("이전단계이동");

            _confirm = ConfirmYN.NonConfirm;
            _state = WeighingState.None;
            _useWeighing = false;
            is_AutoComplete = false;
            //if (_T_CONFIRM != null)
            //{
            //    if (_T_CONFIRM.IsAlive)
            //    {
            //        _T_CONFIRM.Abort();
            //    }
            //}

            if (_weighingStep == WeighingStep.First)
            {
                // 거래처 입력
                _selectedDealerCd = string.Empty;
                _selectedDealerNm = string.Empty;
                _selectedFax = string.Empty;
                _selectedWebFaxYN = string.Empty;

                SelectTabPage(3);
                buttonTab4Next.Visible = false;
                textBoxCustomer.Focus();
            }

            if (_weighingStep == WeighingStep.Second)
            {
                // 차량번호 입력

                // 이전단계로 넘어가도 차량정보는 그대로 둔다..
                // InitVehicleNumber();

                SelectTabPage(2);
            }
        }

        #endregion
        
        #endregion

        #region[인디게이터]
        void OnIndicatorTimer(object state)
        {
            _indicatorTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite);

            // 2초이상 수신없을 시 연결해제 알람
            TimeSpan sp = DateTime.Now - _indicatorReceiveTime;
            _indicatorConnected = sp.TotalSeconds >= 2 ? false : true;

            if (_indicatorPrevConnected != _indicatorConnected)
            {
                Log.AddLog(string.Concat("인디케이터 연결상태 변경 : ", _indicatorPrevConnected.ToString(), " -> ", _indicatorConnected.ToString()));

                indicator.Connect = _indicatorConnected;
            }

            _indicatorPrevConnected = _indicatorConnected;

            if (!_shutdown)
                _indicatorTimer.Change(1000, System.Threading.Timeout.Infinite);
        }



        private void _indicator_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            // 데이터 수신 (데이터 포멧 확인 할 것)
            // - 기본 Data Format : AD4329A 18byte Format
            // - 'H1,H2,Data(8) Kg  '
            //                    Cr
            //                    Lf
            // - '123456789012345678'
            // - 'ST,GS,+00000.0kg  '   <- ex

            // H1 : OL - Over Load
            //      ST - 표시기 안정
            //      US - 표시기 비 안정
            // H2 : NT - Net-Weight (NET)
            //      GS - 실 중량 (GROSS)
            string readData = string.Empty;

            lock (_lockObject)
            {
                try
                {
                    if (!_shutdown)
                    {
                        readData = _indicator.ReadLine();

                        _indicatorReceiveTime = DateTime.Now;

                        if (_writeIndicatorData)
                            Log.AddLog(string.Concat("Indicator data - ", readData));

                        if (readData.Length == 17)
                        {
                            // 계량상태에서만 인디케이터 데이터 처리
                            if (_useWeighing)
                            {
                                // 'US,GS,+12345.0kg' - 비안정
                                // 'ST,GS,+12345.0kg' - 안정

                                // 실데이터 
                                // S??GS?+0000000???
                                //TxtBuffer.Text = readData;
                                string state = readData.Substring(0, 2);
                                string kg = readData.Substring(14, 2).ToLower();
                                int weight = 0;

                                if (int.TryParse(readData.Substring(7, 7), out weight) && kg == "kg")
                                {
                                    // 계량중량표시..
                                    indicator.Weight = weight;

                                    //ERP 요청값 처리를 위하여 추가세팅
                                    Indicator_Thread.Weight = weight;
                                    //_iWeight = weight;//소나무 정보기술 로직 추가 하단 코드 참조

                                    // 계량시작 인식중량보다 크면 계량시작!!
                                    if (weight > _startWeight)
                                    {
                                        if (!_weightFlag)
                                        {
                                            // 계량이 시작됨
                                            _weightFlag = true;

                                            Log.AddLog("계량상태 변경 : Standby -> Start");
                                            _state = WeighingState.Start;
                                            SetCompleteButton(false);
                                            // 계량관련 데이터 초기화
                                            _weight = 0;
                                            _stableCount = 0;
                                        }

                                        if (state == "ST")
                                            _stableCount++;           // 인디케이터 값이 안정적임

                                        if (state == "US")
                                        {
                                            _stableCount = 0;

                                            if (_state == WeighingState.Stable)
                                            {
                                                Log.AddLog("계량상태 변경 : Stable -> Start");

                                                _state = WeighingState.Start;
                                                SetCompleteButton(false);
                                            }

                                            indicator.Stable = false;
                                        }

                                        if (_stableCount >= _stableWeightCount)
                                        {
                                            // 계량 중량 확정    
                                            _weight = weight;

                                            if (_state == WeighingState.Start)
                                            {
                                                Log.AddLog("계량상태 변경 : Start -> Stable");

                                                _state = WeighingState.Stable;

                                                indicator.Stable = true;

                                                //SetCompleteButton(true);
                                                Page4_Button_Click(buttonTab5Complete, null);   //24.04.25
                                                /*
                                                 * 2020-12-08 
                                                 * 중량확정 시 5초간 아무런 액션이 없을 시 자동으로 완료처리되도록 로직추가
                                                 */

                                                //is_AutoComplete = true;
                                                
                                                //if (_T_CONFIRM != null)
                                                //{
                                                //    if (!_T_CONFIRM.IsAlive)
                                                //    {
                                                //        is_AutoComplete = true;
                                                //        _T_CONFIRM = new Thread(AutoComplete);
                                                //        _T_CONFIRM.IsBackground = true;
                                                //        _T_CONFIRM.Start();
                                                //    }
                                                //}
                                                //else
                                                //{
                                                //    is_AutoComplete = true;
                                                //    _T_CONFIRM = new Thread(AutoComplete);
                                                //    _T_CONFIRM.IsBackground = true;
                                                //    _T_CONFIRM.Start();
                                                //}
                                            }
                                        }
                                        else
                                        {
                                            //is_AutoComplete = false;
                                            //if (_T_CONFIRM.IsAlive)
                                            //{
                                            //    _T_CONFIRM.Abort();
                                            //}
                                        }
                                    }
                                    else
                                    {
                                        // 계량상태가 아님..
                                        _weightFlag = false;

                                        if (_state != WeighingState.Standby)
                                        {
                                            Log.AddLog(string.Concat("계량상태 변경 : ", _state.ToString(), " -> Standby"));

                                            indicator.Stable = false;
                                            _state = WeighingState.Standby;
                                        }
                                    }
                                }
                                else
                                {
                                    Log.AddLog(string.Concat("Indicator 데이터 오류 - ", readData));
                                }
                            }
                            /*
                             * 2020-12-07 
                             * 계근값 전송을 위하여 지속적으로 변수에 값 세팅
                             */
                            else
                            {
                                //TxtBuffer.Text = readData;

                                string kg = readData.Substring(14, 2).ToLower();
                                int weight = 0;
                                if (int.TryParse(readData.Substring(7, 7), out weight) && kg == "kg")
                                {
                                    Indicator_Thread.Weight = weight;
                                    //_iWeight = _weight;
                                }
                            }
                        }
                        else
                        {
                            Log.AddLog(string.Concat("Indicator 수신데이터 길이(", readData.Length.ToString(), ") 오류 - ", readData));
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.AddLog(ex.Message);
                }
            }
        }

        private void _indicator_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            Log.AddLog(string.Concat("_indicator_ErrorReceived : ", e.EventType.ToString(), " - ", e.ToString()));
        }
        #endregion

        #region[프린터]
        private void printDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            Graphics g = e.Graphics;
            //display = string.Format("{0,6:N0}{1}{2}", _weight, _space ? " " : "", _unit);

            using (Font font = new Font("굴림", 11))
            {
                using (SolidBrush drawBrush = new SolidBrush(Color.Black))
                {
                    int netWeight = 0;

                    string fullWeight = string.Empty;
                    string emptyWeight = string.Empty;
                    if (_weighingType == WeighingType.In)
                    {
                        fullWeight = string.Concat(string.Format("{0,6:N0} KG", _selectedSecondWeight).PadRight(15), _selectedSecondTime.Substring(11, 5).Replace('-', '/'));
                        emptyWeight = string.Concat(string.Format("{0,6:N0} KG", _selectedFirstWeight).PadRight(15), _selectedFirstTime.Substring(11, 5).Replace('-', '/'));
                        netWeight = _selectedSecondWeight - _selectedFirstWeight - _selectediChaGam;
                    }
                    if (_weighingType == WeighingType.Out)
                    {
                        fullWeight = string.Concat(string.Format("{0,6:N0} KG", _selectedSecondWeight).PadRight(15), _selectedSecondTime.Substring(11, 5).Replace('-', '/'));
                        emptyWeight = string.Concat(string.Format("{0,6:N0} KG", _selectedFirstWeight).PadRight(15), _selectedFirstTime.Substring(11, 5).Replace('-', '/'));
                        netWeight = _selectedSecondWeight - _selectedFirstWeight - _selectedOChaGam;
                    }

                    g.DrawString(_selectedDealerNm, font, drawBrush, new PointF(_X + 130.0F, _Y + 3.0F));                       // 거래처명
                    g.DrawString(_selectedJ_Date.Replace('-', '/'), font, drawBrush, new PointF(_X + 128.0F, _Y + 33.0F));                        // 일자
                    if (_weighingType == WeighingType.In)                                                                       // 입출고
                        g.DrawString("입고", font, drawBrush, new PointF(_X + 315.0F, _Y + 33.0F));
                    if (_weighingType == WeighingType.Out)
                        g.DrawString("출고", font, drawBrush, new PointF(_X + 315.0F, _Y + 33.0F));
                    g.DrawString(_selectedSun, font, drawBrush, new PointF(_X + 150.0F, _Y + 62.0F));                           // 계근번호
                    g.DrawString(_selectedGubun1, font, drawBrush, new PointF(_X + 305.0F, _Y + 62.0F));                        // 품명 ?
                    g.DrawString(_vehicleNumber, font, drawBrush, new PointF(_X + 143.0F, _Y + 92.0F));                         // 차량번호
                    g.DrawString(_selectedEMP_NM, font, drawBrush, new PointF(_X + 300.0F, _Y + 89.0F));                        // 검수자
                    g.DrawString(fullWeight, font, drawBrush, new PointF(_X + 140.0F, _Y + 118.0F));                            // 총중량
                    g.DrawString(emptyWeight.ToString(), font, drawBrush, new PointF(_X + 140.0F, _Y + 147.0F));                // 공차중량

                    // 2020.03.25
                    // - 실중량은 폰트 크게..
                    using (Font bigFont = new Font("굴림", 12, FontStyle.Bold))
                        g.DrawString(string.Format("{0,6:N0} KG", netWeight), bigFont, drawBrush, new PointF(_X + 140.0F, _Y + 176.0F));

                    g.DrawString(_selectedGubun1, font, drawBrush, new PointF(_X + 105.0F, _Y + 232.0F));                        // 등급명
                    if (_weighingType == WeighingType.In)                                                                       // 감량
                        g.DrawString(string.Format("{0} KG", _selectediChaGam.ToString()), font, drawBrush, new PointF(_X + 200.0F, _Y + 232.0F));
                    if (_weighingType == WeighingType.Out)
                        g.DrawString(string.Format("{0} KG", _selectedOChaGam.ToString()), font, drawBrush, new PointF(_X + 200.0F, _Y + 232.0F));

                    using (Font smallFont = new Font("굴림", 8, FontStyle.Bold))
                        g.DrawString(_selectedJ_State, smallFont, drawBrush, new PointF(_X + 273.0F, _Y + 232.0F));                      // 감가.감량사유

                    // 이미지 출력
                    Size size = new Size(160, 110);
                    string basePath = string.Concat(Application.StartupPath, @"\image\print\");
                    string printFile = string.Empty;

                    Bitmap NoImage = Properties.Resources.No_Img;

                    // - 1_1.jpg
                    printFile = string.Concat(basePath, "1_1.jpg");
                    if (File.Exists(printFile))
                    {
                        Bitmap bmp_1_1 = new Bitmap(new Bitmap(printFile), size);
                        g.DrawImage(bmp_1_1, new PointF(_X + 50.0F, _Y + 282.0F));
                    }
                    else
                    {
                        g.DrawImage(NoImage, new PointF(_X + 50.0F, _Y + 282.0F));
                    }
                    // - 1_2.jpg
                    printFile = string.Concat(basePath, "1_2.jpg");
                    if (File.Exists(printFile))
                    {
                        Bitmap bmp_1_2 = new Bitmap(new Bitmap(printFile), size);
                        g.DrawImage(bmp_1_2, new PointF(_X + 50.0F, _Y + 399.0F));
                    }
                    else
                    {
                        g.DrawImage(NoImage, new PointF(_X + 50.0F, _Y + 399.0F));
                    }
                    // - 1_3.jpg
                    printFile = string.Concat(basePath, "1_3.jpg");
                    if (File.Exists(printFile))
                    {
                        Bitmap bmp_1_3 = new Bitmap(new Bitmap(printFile), size);
                        g.DrawImage(bmp_1_3, new PointF(_X + 50.0F, _Y + 516.0F));
                    }
                    else
                    {
                        g.DrawImage(NoImage, new PointF(_X + 50.0F, _Y + 516.0F));
                    }
                    // - 2_1.jpg
                    printFile = string.Concat(basePath, "2_1.jpg");
                    if (File.Exists(printFile))
                    {
                        Bitmap bmp_2_1 = new Bitmap(new Bitmap(printFile), size);
                        g.DrawImage(bmp_2_1, new PointF(_X + 224.0F, _Y + 283.0F));
                    }
                    else
                    {
                        g.DrawImage(NoImage, new PointF(_X + 224.0F, _Y + 283.0F));
                    }
                    // - 2_2.jpg
                    printFile = string.Concat(basePath, "2_2.jpg");
                    if (File.Exists(printFile))
                    {
                        Bitmap bmp_2_2 = new Bitmap(new Bitmap(printFile), size);
                        g.DrawImage(bmp_2_2, new PointF(_X + 224.0F, _Y + 400.0F));
                    }
                    else
                    {
                        g.DrawImage(NoImage, new PointF(_X + 224.0F, _Y + 400.0F));
                    }
                    // - 2_3.jpg
                    printFile = string.Concat(basePath, "2_3.jpg");
                    if (File.Exists(printFile))
                    {
                        Bitmap bmp_2_3 = new Bitmap(new Bitmap(printFile), size);
                        g.DrawImage(bmp_2_3, new PointF(_X + 224.0F, _Y + 517.0F));
                    }
                    else
                    {
                        g.DrawImage(NoImage, new PointF(_X + 224.0F, _Y + 517.0F));
                    }
                }
            }
        }
        #endregion

        #region [키보드 클릭]


        #region[2021-01-14 기준 사용하지 않는 배열]
        /*
         * 2021-01-14 현업요청
         * 현장 운전자들이 거래처 입력 시 레이아웃의 항목들이 많고
         * 특수문자 같은 배열들은 사용하지 않는다고 하여 문자배열 및 SPACE 등고 같이 필수요소들만 남기고 나머지는 주석처리
         */

        //"~"
        private void pictureBoxOem1_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.Oem3, 0, 0, 0);      //키다운
            keybd_event((byte)Keys.Oem3, 0, 0x02, 0);   //키업

            ShiftClear();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.D1, 0, 0, 0);   //키다운
            keybd_event((byte)Keys.D1, 0, 0x02, 0);//키업

            ShiftClear();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.D2, 0, 0, 0);   //키다운
            keybd_event((byte)Keys.D2, 0, 0x02, 0);//키업

            ShiftClear();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.D3, 0, 0, 0);   //키다운
            keybd_event((byte)Keys.D3, 0, 0x02, 0);//키업

            ShiftClear();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.D4, 0, 0, 0);
            keybd_event((byte)Keys.D4, 0, 0x02, 0);

            ShiftClear();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.D5, 0, 0, 0);
            keybd_event((byte)Keys.D5, 0, 0x02, 0);

            ShiftClear();
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.D6, 0, 0, 0);
            keybd_event((byte)Keys.D6, 0, 0x02, 0);

            ShiftClear();
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.D7, 0, 0, 0);
            keybd_event((byte)Keys.D7, 0, 0x02, 0);

            ShiftClear();
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.D8, 0, 0, 0);
            keybd_event((byte)Keys.D8, 0, 0x02, 0);

            ShiftClear();
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.D9, 0, 0, 0);
            keybd_event((byte)Keys.D9, 0, 0x02, 0);

            ShiftClear();
        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.D0, 0, 0, 0);
            keybd_event((byte)Keys.D0, 0, 0x02, 0);

            ShiftClear();
        }

        //"-"
        private void pictureBoxOem2_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.OemMinus, 0, 0, 0);
            keybd_event((byte)Keys.OemMinus, 0, 0x02, 0);

            ShiftClear();
        }

        //"="
        private void pictureBoxOem3_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.Oemplus, 0, 0, 0);
            keybd_event((byte)Keys.Oemplus, 0, 0x02, 0);

            ShiftClear();
        }

        private void pictureBoxTab_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.Tab, 0, 0, 0);
            keybd_event((byte)Keys.Tab, 0, 0x02, 0);
        }

        private void pictureBoxCapsLock_Click(object sender, EventArgs e)
        {
            return;

            if (_bVKCapsLock == true)
                pictureBoxCapsLock.BackgroundImage = Properties.Resources.CapsLock;
            else
                pictureBoxCapsLock.BackgroundImage = Properties.Resources.CapsLock_on;

            _bVKCapsLock = !_bVKCapsLock;

            keybd_event((byte)Keys.CapsLock, 0, 0, 0);
            keybd_event((byte)Keys.CapsLock, 0, 0x02, 0);
        }

        //"["
        private void pictureBoxOem4_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.Oem4, 0, 0, 0);
            keybd_event((byte)Keys.Oem4, 0, 0x02, 0);

            ShiftClear();
        }

        //"]"
        private void pictureBoxOem5_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.Oem6, 0, 0, 0);
            keybd_event((byte)Keys.Oem6, 0, 0x02, 0);

            ShiftClear();
        }

        //"\"
        private void pictureBoxOem6_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.Oem5, 0, 0, 0);
            keybd_event((byte)Keys.Oem5, 0, 0x02, 0);

            ShiftClear();
        }

        //";"
        private void pictureBoxOem7_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.OemSemicolon, 0, 0, 0);
            keybd_event((byte)Keys.OemSemicolon, 0, 0x02, 0);

            ShiftClear();
        }

        //" ' "
        private void pictureBoxOem8_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.OemQuotes, 0, 0, 0);
            keybd_event((byte)Keys.OemQuotes, 0, 0x02, 0);

            ShiftClear();
        }

        //"<"
        private void pictureBoxOem9_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.Oemcomma, 0, 0, 0);
            keybd_event((byte)Keys.Oemcomma, 0, 0x02, 0);

            ShiftClear();
        }

        //">"
        private void pictureBoxOem10_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.OemPeriod, 0, 0, 0);
            keybd_event((byte)Keys.OemPeriod, 0, 0x02, 0);

            ShiftClear();
        }

        //"?"
        private void pictureBox11_Click(object sender, EventArgs e)
        {
            return;

            keybd_event((byte)Keys.OemQuestion, 0, 0, 0);
            keybd_event((byte)Keys.OemQuestion, 0, 0x02, 0);

            ShiftClear();
        }

        #endregion[2021-01-14 기준 사용하지 않는 배열]


        private void pictureBoxBackspace_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            keybd_event((byte)Keys.Back, 0, 0, 0);
            keybd_event((byte)Keys.Back, 0, 0x02, 0);
        }
        

        private void pictureBoxQ_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.F, 0, 0, 0);
                keybd_event((byte)Keys.F, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.Q, 0, 0, 0);
                keybd_event((byte)Keys.Q, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"ㅃ"
        private void pictureBoxQ_1_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (!_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.LShiftKey, 0, 0, 0);
                keybd_event((byte)Keys.Q, 0, 0, 0);
                keybd_event((byte)Keys.Q, 0, 0x02, 0);
                keybd_event((byte)Keys.LShiftKey, 0, 0x02, 0);
            }
            ShiftClear();
        }

        //"I", "ㅈ"
        private void pictureBoxW_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.I, 0, 0, 0);
                keybd_event((byte)Keys.I, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.W, 0, 0, 0);
                keybd_event((byte)Keys.W, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"ㅉ"
        private void pictureBoxW_1_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (!_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.LShiftKey, 0, 0, 0);
                keybd_event((byte)Keys.W, 0, 0, 0);
                keybd_event((byte)Keys.W, 0, 0x02, 0);
                keybd_event((byte)Keys.LShiftKey, 0, 0x02, 0);
            }
            ShiftClear();
        }

        //"C", "ㄷ"
        private void pictureBoxE_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.C, 0, 0, 0);
                keybd_event((byte)Keys.C, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.E, 0, 0, 0);
                keybd_event((byte)Keys.E, 0, 0x02, 0);
            }
                

            ShiftClear();
        }

        //"ㄸ"
        private void pictureBoxE_1_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (!_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.LShiftKey, 0, 0, 0);
                keybd_event((byte)Keys.E, 0, 0, 0);
                keybd_event((byte)Keys.E, 0, 0x02, 0);
                keybd_event((byte)Keys.LShiftKey, 0, 0x02, 0);
            }
            ShiftClear();
        }

        //"A", "ㄱ"
        private void pictureBoxR_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.A, 0, 0, 0);
                keybd_event((byte)Keys.A, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.R, 0, 0, 0);
                keybd_event((byte)Keys.R, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"ㄲ"
        private void pictureBoxR_1_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (!_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.LShiftKey, 0, 0, 0);
                keybd_event((byte)Keys.R, 0, 0, 0);
                keybd_event((byte)Keys.R, 0, 0x02, 0);
                keybd_event((byte)Keys.LShiftKey, 0, 0x02, 0);
            }
            ShiftClear();
        }

        //"G", "ㅅ"
        private void pictureBoxT_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.G, 0, 0, 0);
                keybd_event((byte)Keys.G, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.T, 0, 0, 0);
                keybd_event((byte)Keys.T, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"ㅆ"
        private void pictureBoxT_1_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (!_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.LShiftKey, 0, 0, 0);
                keybd_event((byte)Keys.T, 0, 0, 0);
                keybd_event((byte)Keys.T, 0, 0x02, 0);
                keybd_event((byte)Keys.LShiftKey, 0, 0x02, 0);
            }
            ShiftClear();
        }
        
        //"T", "ㅛ"
        private void pictureBoxY_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.T, 0, 0, 0);
                keybd_event((byte)Keys.T, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.Y, 0, 0, 0);
                keybd_event((byte)Keys.Y, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"R", "ㅕ"
        private void pictureBoxU_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.R, 0, 0, 0);
                keybd_event((byte)Keys.R, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.U, 0, 0, 0);
                keybd_event((byte)Keys.U, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"P", "ㅑ"
        private void pictureBoxI_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.P, 0, 0, 0);
                keybd_event((byte)Keys.P, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.I, 0, 0, 0);
                keybd_event((byte)Keys.I, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"ㅐ"
        private void pictureBoxO_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (!_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.O, 0, 0, 0);
                keybd_event((byte)Keys.O, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"ㅒ"
        private void pictureBoxO_1_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (!_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.LShiftKey, 0, 0, 0);
                keybd_event((byte)Keys.O, 0, 0, 0);
                keybd_event((byte)Keys.O, 0, 0x02, 0);
                keybd_event((byte)Keys.LShiftKey, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"Y", "ㅔ"
        private void pictureBoxP_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.Y, 0, 0, 0);
                keybd_event((byte)Keys.Y, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.P, 0, 0, 0);
                keybd_event((byte)Keys.P, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"Z", "ㅖ"
        private void pictureBoxP_1_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.Z, 0, 0, 0);
                keybd_event((byte)Keys.Z, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.LShiftKey, 0, 0, 0);
                keybd_event((byte)Keys.P, 0, 0, 0);
                keybd_event((byte)Keys.P, 0, 0x02, 0);
                keybd_event((byte)Keys.LShiftKey, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"E", "ㅁ"
        private void pictureBoxA_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.E, 0, 0, 0);
                keybd_event((byte)Keys.E, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.A, 0, 0, 0);
                keybd_event((byte)Keys.A, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"B" , "ㄴ"
        private void pictureBoxS_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.B, 0, 0, 0);
                keybd_event((byte)Keys.B, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.S, 0, 0, 0);
                keybd_event((byte)Keys.S, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"H", "ㅇ"
        private void pictureBoxD_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.H, 0, 0, 0);
                keybd_event((byte)Keys.H, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.D, 0, 0, 0);
                keybd_event((byte)Keys.D, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"D", "ㄹ"
        private void pictureBoxF_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.D, 0, 0, 0);
                keybd_event((byte)Keys.D, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.F, 0, 0, 0);
                keybd_event((byte)Keys.F, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"N", "ㅎ"
        private void pictureBoxG_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.N, 0, 0, 0);
                keybd_event((byte)Keys.N, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.G, 0, 0, 0);
                keybd_event((byte)Keys.G, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"S", "ㅗ"
        private void pictureBoxH_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.S, 0, 0, 0);
                keybd_event((byte)Keys.S, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.H, 0, 0, 0);
                keybd_event((byte)Keys.H, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"Q", "ㅓ"
        private void pictureBoxJ_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.Q, 0, 0, 0);
                keybd_event((byte)Keys.Q, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.J, 0, 0, 0);
                keybd_event((byte)Keys.J, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"O", "ㅏ"
        private void pictureBoxK_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.O, 0, 0, 0);
                keybd_event((byte)Keys.O, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.K, 0, 0, 0);
                keybd_event((byte)Keys.K, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"X", "ㅣ"
        private void pictureBoxL_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.X, 0, 0, 0);
                keybd_event((byte)Keys.X, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.L, 0, 0, 0);
                keybd_event((byte)Keys.L, 0, 0x02, 0);
            }

            ShiftClear();
        }

        private void pictureBoxEnter_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            keybd_event((byte)Keys.Enter, 0, 0, 0);
            keybd_event((byte)Keys.Enter, 0, 0x02, 0);
        }

        private void pictureBoxShiftL_Click(object sender, EventArgs e)
        {
            // 이미지 변경
            // VK 애니메이션 제어.
            if (_bVKLShiftButtonState == false)
            {
                _bVKLShiftButtonState = true;

                // 피쳐박스에뿌려주기
                pictureBoxShiftL.BackgroundImage = Properties.Resources.Shift_L_on;

                keybd_event((byte)Keys.LShiftKey, 0, 0, 0);

            }
            else
            {
                _bVKLShiftButtonState = false;

                // 피쳐박스에뿌려주기
                pictureBoxShiftL.BackgroundImage = Properties.Resources.Shift_L_off;

                keybd_event((byte)Keys.LShiftKey, 0, 0x02, 0);
            }
        }

        //"K", "ㅋ"
        private void pictureBoxZ_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.K, 0, 0, 0);
                keybd_event((byte)Keys.K, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.Z, 0, 0, 0);
                keybd_event((byte)Keys.Z, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"L", "ㅌ"
        private void pictureBoxX_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.L, 0, 0, 0);
                keybd_event((byte)Keys.L, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.X, 0, 0, 0);
                keybd_event((byte)Keys.X, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"J", "ㅊ"
        private void pictureBoxC_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.J, 0, 0, 0);
                keybd_event((byte)Keys.J, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.C, 0, 0, 0);
                keybd_event((byte)Keys.C, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"M", "ㅍ"
        private void pictureBoxV_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.M, 0, 0, 0);
                keybd_event((byte)Keys.M, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.V, 0, 0, 0);
                keybd_event((byte)Keys.V, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"V", "ㅠ"
        private void pictureBoxB_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.V, 0, 0, 0);
                keybd_event((byte)Keys.V, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.B, 0, 0, 0);
                keybd_event((byte)Keys.B, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"U", "ㅜ"
        private void pictureBoxN_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.U, 0, 0, 0);
                keybd_event((byte)Keys.U, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.N, 0, 0, 0);
                keybd_event((byte)Keys.N, 0, 0x02, 0);
            }

            ShiftClear();
        }

        //"W", "ㅡ"
        private void pictureBoxM_Click(object sender, EventArgs e)
        {
            SetDealerInfo();
            if (_bVKKorEngButtonState)
            {
                keybd_event((byte)Keys.W, 0, 0, 0);
                keybd_event((byte)Keys.W, 0, 0x02, 0);
            }
            else
            {
                keybd_event((byte)Keys.M, 0, 0, 0);
                keybd_event((byte)Keys.M, 0, 0x02, 0);
            }

            ShiftClear();
        }
        
        /*
         * 2021-01-14 현업요청
         * 거래처 검색 후 검색창에서 다른영역 선택 시 검색창은 사라지도록 선택하도록 수정
         */
        private void SetDealerInfo()
        {
            if(listBoxCustomers.Visible == true)
            {
                listBoxCustomers.Visible = false;
                // 거래처관련 변수들..
                _selectedDealerCd = string.Empty;
                _selectedDealerNm = string.Empty;
                _selectedFax = string.Empty;
                _selectedWebFaxYN = string.Empty;
                buttonCustomerClear.PerformClick();
            }
            else
            {
                _selectedDealerCd = string.Empty;
                _selectedDealerNm = string.Empty;
                _selectedFax = string.Empty;
                _selectedWebFaxYN = string.Empty;
                buttonTab4Next.Visible = false;
            }
        }

        private void pictureBoxShiftR_Click(object sender, EventArgs e)
        {
            // 이미지 변경
            // VK 애니메이션 제어.
            if (_bVKRShiftButtonState == false)
            {
                _bVKRShiftButtonState = true;

                // 피쳐박스에뿌려주기
                pictureBoxShiftR.BackgroundImage = Properties.Resources.Shift_R_on;

                keybd_event((byte)Keys.RShiftKey, 0, 0, 0);

            }
            else
            {
                _bVKRShiftButtonState = false;

                // 피쳐박스에뿌려주기
                pictureBoxShiftR.BackgroundImage = Properties.Resources.Shift_R_off;

                keybd_event((byte)Keys.RShiftKey, 0, 0x02, 0);
            }
        }

        private void pictureBoxSpace_Click(object sender, EventArgs e)
        {
            keybd_event((byte)Keys.Space, 0, 0, 0);
            keybd_event((byte)Keys.Space, 0, 0x02, 0);
        }

        private void pictureBoxHangul_Click(object sender, EventArgs e)
        {
            // 이미지 변경
            // VK 애니메이션 제어.
            if (_bVKKorEngButtonState == false)
            {
                _bVKKorEngButtonState = true;

                // 피쳐박스에뿌려주기
                pictureBoxHangul.BackgroundImage = Properties.Resources.한영_영문;
            }
            else
            {
                _bVKKorEngButtonState = false;

                // 피쳐박스에뿌려주기
                pictureBoxHangul.BackgroundImage = Properties.Resources.한영_한글;
            }

            keybd_event((byte)Keys.HangulMode, 0, 0, 0);
        }

        private void ShiftClear()
        {
            if (_bVKLShiftButtonState)
            {
                _bVKLShiftButtonState = false;

                // 피쳐박스에뿌려주기
                pictureBoxShiftL.BackgroundImage = Properties.Resources.Shift_L_off;

                keybd_event((byte)Keys.LShiftKey, 0, 0x02, 0);
            }

            if (_bVKRShiftButtonState)
            {
                _bVKRShiftButtonState = false;

                // 피쳐박스에뿌려주기
                pictureBoxShiftR.BackgroundImage = Properties.Resources.Shift_R_off;

                keybd_event((byte)Keys.RShiftKey, 0, 0x02, 0);
            }
        }

        #endregion

        #region[기초세팅]

        private void InitWeighingValue()
        {
            _weighingType = WeighingType.None;
            _weighingStep = WeighingStep.None;
            _vehicleNumber = string.Empty;

            _state = WeighingState.None;
            _weightFlag = false;
            _stableCount = 0;
            _weight = 0;

            //2020-12-16
            //자동등록을 위하여 변수추가
            _confirm = ConfirmYN.NonConfirm;

            // 2020.03.17
            // - 거래처정보 추가
            _selectedDealerCd = string.Empty;
            _selectedDealerNm = string.Empty;
            _selectedFax = string.Empty;
            _selectedWebFaxYN = string.Empty;

            // 2020.03.18
            // - 실적정보 추가
            _selectedJunpyoID = string.Empty;                   // 전표아이디
            _selectedJ_Date = string.Empty;                     // 일자
            _selectedSun = string.Empty;                        // 계근번호(순번)
            _selectedMaipCherID = string.Empty;                 // 매입처코드
            _selectedMaipCher = string.Empty;                   // 매입처
            _selectedJ_AssignID = string.Empty;                 // 매출처코드
            _selectedJ_Company = string.Empty;                  // 매출처
            _selectedFirstTime = string.Empty;                  // 1차계량 (입/출고에 따라 변경) - 시간
            _selectedFirstWeight = 0;                           //                             - 중량 
            _selectedSecondTime = string.Empty;                 // 2차계량 (입/출고에 따라 변경) - 시간
            _selectedSecondWeight = 0;                          //                             - 중량
            _selectedJ_Serial = string.Empty;                   // 검수정보 - 검수여부 ("0" 이 아니면 검수)
            _selectedGubun1 = string.Empty;                     //         - 등급
            _selectediChaGam = 0;                               //         - 매입감량
            _selectedOChaGam = 0;                               //         - 매출감량
            _selectedJ_State = string.Empty;                    //         - 감가/감량사유
            _selectedgumsubigo = string.Empty;                  //         - 검수비고
            _selectedgumsu_serial = string.Empty;               //         - 검수자 사원번호
            _selectedEMP_NM = string.Empty;                     //         - 검수자명

            // 2020.03.20
            // - ticket image 추가
            _selectedTicketImage_1_1 = string.Empty;
            _selectedTicketImage_1_2 = string.Empty;
            _selectedTicketImage_1_3 = string.Empty;
            _selectedTicketImage_2_1 = string.Empty;
            _selectedTicketImage_2_2 = string.Empty;
            _selectedTicketImage_2_3 = string.Empty;
        }

        // 차량번호 초기화
        private void InitVehicleNumber()
        {
            _vehicleNumber = string.Empty;
            textBoxVehicleNumber.Text = string.Empty;
        }

        private void InitUserControl()
        {
            tabControl1.SelectedIndex = 0;

            tabControl1.Appearance = TabAppearance.FlatButtons;
            tabControl1.ItemSize = new Size(0, 1);
            tabControl1.SizeMode = TabSizeMode.Fixed;

            foreach (TabPage tab in tabControl1.TabPages)
            {
                tab.Text = "";
            }
        }

        private void Initialize()
        {
            //화상키보드를 위한 후킹 함수 등록
            InitHook(panelKeyboard.Handle);                     // 후킹 컨트롤 등록
            InstallHook();

            // 인디케이터 설정
            _indicator = new SerialPort();
            _indicator.DataReceived += _indicator_DataReceived;
            _indicator.ErrorReceived += _indicator_ErrorReceived;

            int indicatorBaudRate = 0;
            int.TryParse(_indicatorBaudRate, out indicatorBaudRate);
            int indicatorDataBit = 0;
            int.TryParse(_indicatorDataBit, out indicatorDataBit);

            StopBits indicatorStopBits;
            switch (_indicatorStopBit.ToLower())
            {
                case "one": indicatorStopBits = StopBits.One; break;
                case "two": indicatorStopBits = StopBits.Two; break;
                case "onepointfive": indicatorStopBits = StopBits.OnePointFive; break;
                default: indicatorStopBits = StopBits.None; break;
            }
            Parity indicatorParity;
            switch (_indicatorParityBit.ToLower())
            {
                case "none": indicatorParity = Parity.None; break;
                case "odd": indicatorParity = Parity.Odd; break;
                case "even": indicatorParity = Parity.Even; break;
                case "mark": indicatorParity = Parity.Mark; break;
                case "space": indicatorParity = Parity.Space; break;
                default: indicatorParity = Parity.None; break;
            }

            _indicator.PortName = _indicatorPortName;
            _indicator.BaudRate = indicatorBaudRate;
            _indicator.DataBits = indicatorDataBit;
            _indicator.StopBits = indicatorStopBits;
            _indicator.Parity = indicatorParity;

            _indicatorReceiveTime = DateTime.MinValue;

            // DB 설정 - 커넥션스트링만 만들어 놓는다..
            _connectionString = string.Concat
            (
                "SERVER = ", _dbAddress, "; ",
                "DATABASE = ", _dbName, "; ",
                "UID= ", _dbUser, "; ",
                "PASSWORD = ", _dbPassword, "; "
               // "Integrated Security = SSPI;"           // 윈도우 인증 시 연결스트링에 추가
               //"Allow Zero Datetime = True;"

                //로컬 테스트용
                //"SERVER = ", _dbAddress, "; ",
                //"DATABASE = ", _dbName, "; ",
                //"UID= ", _dbUser, "; ",
                //"PASSWORD = ", _dbPassword, "; "
                //"Integrated Security = SSPI;"           // 윈도우 인증 시 연결스트링에 추가
            
            );

            // FTP 설정
            _ftpClient.Initialize(_ftpAddress, _ftpUser, _ftpPassword);
            _ftpClient.FtpDirectioryCheck(_ftpUploadPath);

            // 기타..

            _indicatorTimer = new System.Threading.Timer
            (
                new System.Threading.TimerCallback(OnIndicatorTimer),
                this,
                System.Threading.Timeout.Infinite,
                System.Threading.Timeout.Infinite
            );
            _ftpTimer = new System.Threading.Timer
            (
                new System.Threading.TimerCallback(OnFtpTimer),
                this,
                System.Threading.Timeout.Infinite,
                System.Threading.Timeout.Infinite
            );
            // 전표저장용 디렉토리 생성
            string ticketDir = string.Concat(Application.StartupPath, @"\ticket");
            if (!Directory.Exists(ticketDir))
            {
                Directory.CreateDirectory(ticketDir);
                Log.AddLog(string.Concat("Create Directory..", ticketDir));
            }
            // 이미지 디렉토리 생성
            string imageDir = string.Concat(Application.StartupPath, @"\image");
            if (!Directory.Exists(imageDir))
            {
                Directory.CreateDirectory(imageDir);
                Log.AddLog(string.Concat("Create Directory..", imageDir));
            }
            imageDir = string.Concat(Application.StartupPath, @"\image\upload");
            if (!Directory.Exists(imageDir))
            {
                Directory.CreateDirectory(imageDir);
                Log.AddLog(string.Concat("Create Directory..", imageDir));
            }
            imageDir = string.Concat(Application.StartupPath, @"\image\print");
            if (!Directory.Exists(imageDir))
            {
                Directory.CreateDirectory(imageDir);
                Log.AddLog(string.Concat("Create Directory..", imageDir));
            }
            
            Thread.Sleep(1000);
            this.BringToFront();

            // 2020.04.02
            // - 로그삭제
            Log.DeleteLogFile(_logDay);
        }

        // 오류 메시지 및 로그저장과 함께 로직 종료 메서드
        private void QuitwithLog(string msg)
        {
            if (!string.IsNullOrEmpty(msg))
            {
                Log.AddLog(msg);
                MessageBox.Show(new Form { TopMost = true }, msg);
                return;
            }
        }

        // 오류 메시지 및 로그저장과 함께 로직 종료 메서드 (중량확정 & 계량완료 페이지 (Page 4))
        private void QuitwithLog_Page4(string msg)
        {
            if (!string.IsNullOrEmpty(msg))
            {
                Log.AddLog(msg);
                MessageBox.Show(new Form { TopMost = true }, msg);
                // 버튼 활성화
                SetCompleteButton2(true);
                SetPrvButton2(true);
                return;
            }
        }
        /*
         * 2020-12-14 3초간 계근값 유지 후 5초간 아무런 액션이 없을 시 완료처리 하기위한 쓰레드 추가
         */
        Thread _T_CONFIRM;
        DateTime _CurTime;
        private bool is_AutoComplete = false;
        private void AutoComplete()
        {
            _CurTime = DateTime.Now;
            while (is_AutoComplete)
            {
                if (_confirm == ConfirmYN.NonConfirm && _state == WeighingState.Stable)
                {
                    //5초가 지나도록 완료버튼 처리가 되지 않을 시 저장처리
                    TimeSpan dateDiff = DateTime.Now - _CurTime;
                    if (dateDiff.Seconds > 3)
                    {
                        if (buttonTab5Complete.InvokeRequired)
                        {
                            buttonTab5Complete.BeginInvoke(new MethodInvoker(delegate
                            {
                                buttonTab5Complete.Enabled = false;
                                buttonTab5Complete.Text = "완료";
                            }));
                        }
                        else
                        {
                            buttonTab5Complete.Enabled = false;
                        }

                        if (buttonTab5Prev.InvokeRequired)
                        {
                            buttonTab5Prev.BeginInvoke(new MethodInvoker(delegate
                            {
                                buttonTab5Prev.Enabled = false;
                            }));
                        }
                        else
                        {
                            buttonTab5Prev.Enabled = false;
                        }

                        Page4_Button_Click(buttonTab5Complete, null);

                        if (buttonTab5Complete.InvokeRequired)
                        {
                            buttonTab5Complete.BeginInvoke(new MethodInvoker(delegate
                            {
                                buttonTab5Complete.Enabled = true;
                                buttonTab5Complete.Text = "완료";
                            }));
                        }
                        else
                        {
                            buttonTab5Complete.Enabled = true;
                        }

                        if (buttonTab5Prev.InvokeRequired)
                        {
                            buttonTab5Prev.BeginInvoke(new MethodInvoker(delegate
                            {
                                buttonTab5Prev.Enabled = true;
                            }));
                        }
                        else
                        {
                            buttonTab5Prev.Enabled = true;
                        }

                        _CurTime = DateTime.Now;
                        is_AutoComplete = false;
                        //if (_T_CONFIRM != null)
                        //{
                        //    if (_T_CONFIRM.IsAlive)
                        //    {
                        //        _T_CONFIRM.Abort();
                        //        break;
                        //    }
                        //    else
                        //    {

                        //    }
                        //}
                    }
                }
                else
                {
                    is_AutoComplete = false;
                    //if (_T_CONFIRM != null)
                    //{
                    //    if (_T_CONFIRM.IsAlive)
                    //    {
                    //        _T_CONFIRM.Abort();
                    //        break;
                    //    }
                    //}
                    //else
                    //{
                    //    _T_CONFIRM.Abort();
                    //    break;
                    //}
                }

                Thread.Sleep(100);
            }
        }

        private void ReadSettingData()
        {
            _ini = new IniFile(string.Concat(Application.StartupPath, @"\setting.ini"));

            try
            {
                 _startWeight = Convert.ToInt32(_ini.GetIniValue("ENV", "StartWeight"));
                _secondWeighingDay = Convert.ToInt32(_ini.GetIniValue("ENV", "SecondWeighingDay"));
                _stableWeightCount = Convert.ToInt32(_ini.GetIniValue("ENV", "StableWeightCount"));
                _writeIndicatorData = Convert.ToBoolean(_ini.GetIniValue("ENV", "WriteIndicatorData"));
                _faxServiceIsTest = Convert.ToBoolean(_ini.GetIniValue("ENV", "FaxServiceIsTest"));
                _faxNumber = _ini.GetIniValue("ENV", "FaxNumber");
                _saveImage = Convert.ToBoolean(_ini.GetIniValue("ENV", "SaveImage"));
                _ticketPrint = Convert.ToBoolean(_ini.GetIniValue("ENV", "TicketPrint"));
                _X = Convert.ToSingle(_ini.GetIniValue("ENV", "PointX"));
                _Y = Convert.ToSingle(_ini.GetIniValue("ENV", "PointY"));
                _logDay = Convert.ToInt32(_ini.GetIniValue("ENV", "LogDay"));

                _indicatorPortName = _ini.GetIniValue("INDICATOR", "PortName");
                _indicatorBaudRate = _ini.GetIniValue("INDICATOR", "BaudRate");
                _indicatorDataBit = _ini.GetIniValue("INDICATOR", "DataBit");
                _indicatorStopBit = _ini.GetIniValue("INDICATOR", "StopBit");
                _indicatorParityBit = _ini.GetIniValue("INDICATOR", "Parity");

                _dbAddress = _ini.GetIniValue("DATABASE", "Address");
                _dbName = _ini.GetIniValue("DATABASE", "Database");
                _dbUser = _ini.GetIniValue("DATABASE", "User");
                _dbPassword = _ini.GetIniValue("DATABASE", "Password", true);     // 데이터의 암호화 상태여부
                //_dbPassword = _ini.GetIniValue("DATABASE", "Password"); //로컬테스트용

                _ftpAddress = _ini.GetIniValue("FTP", "Address");
                _ftpUser = _ini.GetIniValue("FTP", "User");
                _ftpPassword = _ini.GetIniValue("FTP", "Password", true); // 데이터의 암호화 상태여부
                //_ftpPassword = _ini.GetIniValue("FTP", "Password");  //로컬테스트용
                _ftpUploadPath = _ini.GetIniValue("FTP", "UploadPath");

                _sound = _ini.GetIniValue("SOUND", "SOUND");

                //2020-12-08 추가
                int.TryParse(_ini.GetIniValue("ETC", "AutoCompleteTime"), out _autoCompleteSecond);
            }
            catch (Exception ex)
            {
                Log.AddLog(ex.Message);
                MessageBox.Show(new Form { TopMost = true }, ex.Message);
            }
        }

        private void Start()
        {
            try
            {
                _indicator.Open();
                
                _indicatorTimer.Change(1000, System.Threading.Timeout.Infinite);
                _ftpTimer.Change(1000, System.Threading.Timeout.Infinite);
            }
            catch (Exception ex)
            {
                Log.AddLog(ex.Message);
                MessageBox.Show(new Form { TopMost = true }, ex.Message);
            }
        }

        private void Stop()
        {
            try
            {
                _shutdown = true;

                _indicatorTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite);
                _ftpTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite);

                if (_indicator.IsOpen)
                {
                    _indicator.DataReceived -= _indicator_DataReceived;
                    _indicator.ErrorReceived -= _indicator_ErrorReceived;
                }

                _indicator.Close();
                _indicator = null;
            }
            catch (Exception ex)
            {
                Log.AddLog(ex.Message);
            }
        }
        
        private void InputVehicleNumber(Button button)
        {
            string inputText = textBoxVehicleNumber.Text.Trim();
            string buttonText = button.Text.Trim();

            if (buttonText == "←")
            {
                if (inputText.Length > 0)
                    inputText = inputText.Substring(0, inputText.Length - 1);
            }
            else
            {
                if (inputText.Length == 4)
                    return;

                inputText += buttonText;
            }

            _vehicleNumber = inputText;
            textBoxVehicleNumber.Text = inputText;
        }

        // 키 입력 유효성 검사
        private void TextVehicleNumber_KeyDown(object sender, KeyEventArgs e)
        {
            // 숫자 0 ~ 9 허용
            if (e.KeyCode >= Keys.D0 && e.KeyCode <= Keys.D9) { e.SuppressKeyPress = false; }
            // 숫자패드 (텐키) 0 ~ 9 허용
            else if (e.KeyCode >= Keys.NumPad0 && e.KeyCode <= Keys.NumPad9) { e.SuppressKeyPress = false; }
            // ESC / Delete / 백스페이스 / 탭키 허용
            else if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back || e.KeyCode == Keys.Tab || e.KeyCode == Keys.Escape) { e.SuppressKeyPress = false; }
            // 상하좌우 방향키 / 엔터키 허용
            else if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right || e.KeyCode == Keys.Up || e.KeyCode == Keys.Down || e.KeyCode == Keys.Enter) { e.SuppressKeyPress = false; }
            // 그 외 나머지 키 입력 불허
            else { e.SuppressKeyPress = true; }
        }

        private int ExecuteNonQuery(string query, out string error)
        {
            error = string.Empty;
            
            try
            {
                using (SqlConnection conn = new SqlConnection(_connectionString))
                {
                    conn.Open();

                    SqlCommand cmd = new SqlCommand(query, conn);
                    return cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;

                return -1;
            }
        }

        public object ExecuteScalar(string query, out string error)
        {
            error = string.Empty;
            
            try
            {
                using (SqlConnection conn = new SqlConnection(_connectionString))
                {
                    conn.Open();

                    SqlCommand cmd = new SqlCommand(query, conn);
                    return cmd.ExecuteScalar();
                    
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }

            return null;
        }

        private DataSet OpenSnapshot(string query, out string error)
        {
            error = string.Empty;

            DataSet ds = new DataSet();
            
            try
            {
                using (SqlConnection conn = new SqlConnection(_connectionString))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter(query, conn);
                    adapter.Fill(ds);

                    return ds;
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;

                return ds;
            }
        }

        private void SelectTabPage(int no)
        {
            if (tabControl1.InvokeRequired)
            {
                SelectTabPageCallback d = new SelectTabPageCallback(SelectTabPage);
                this.Invoke(d, new object[] { no });
            }
            else
            {
                tabControl1.SelectedIndex = no;
            }
        }

        private void SetCompleteButton(bool enabled)
        {
            if (buttonTab5Complete.InvokeRequired)
            {
                SetCompleteButtonCallback d = new SetCompleteButtonCallback(SetCompleteButton);
                this.Invoke(d, new object[] { enabled });
            }
            else
            {
                if (enabled)
                {
                    buttonTab5Complete.BackColor = Color.Orange;
                }
                else
                {
                    buttonTab5Complete.BackColor = Color.LightGray;
                }
                /*
                 * 2021-01-22
                 * 현업요청
                 * 계근확정 전까지는 Visibla false 처리
                 */
                buttonTab5Complete.Visible = enabled;
                buttonTab5Complete.Enabled = enabled;
            }
        }

        private void SetCompleteButtonText(bool enabled, string sText)
        {
            if (buttonTab5Complete.InvokeRequired)
            {
                SetCompleteButtonCallback2 d = new SetCompleteButtonCallback2(SetCompleteButtonText);
                this.Invoke(d, new object[] { enabled, sText });
            }
            else
            {
                buttonTab5Complete.Visible = enabled;
                buttonTab5Complete.Text = sText;
                buttonTab5Complete.Enabled = enabled;
            }
        }

        private void SetCompleteButton2(bool enabled)
        {
            if (buttonTab5Complete.InvokeRequired)
            {
                SetCompleteButtonCallback3 d = new SetCompleteButtonCallback3(SetCompleteButton2);
                this.Invoke(d, new object[] { enabled });
            }
            else
            {
                buttonTab5Complete.Enabled = enabled;
            }
        }

        private void SetPrvButton2(bool enabled)
        {
            if (buttonTab5Prev.InvokeRequired)
            {
                SetCompleteButtonCallback3 d = new SetCompleteButtonCallback3(SetPrvButton2);
                this.Invoke(d, new object[] { enabled });
            }
            else
            {
                buttonTab5Prev.Enabled = enabled;
            }
        }

        private void SetCompleteButtonVisible(bool enabled)
        {
            if (buttonTab5Complete.InvokeRequired)
            {
                SetCompleteButtonCallback3 d = new SetCompleteButtonCallback3(SetCompleteButtonVisible);
                this.Invoke(d, new object[] { enabled });
            }
            else
            {
                buttonTab5Complete.Visible = enabled;
            }
        }

        private void SetInitVehicleNumber(bool enabled)
        {
            if (textBoxVehicleNumber.InvokeRequired)
            {
                SetCompleteButtonCallback3 d = new SetCompleteButtonCallback3(SetInitVehicleNumber);
                this.Invoke(d, new object[] { enabled });
            }
            else
            {
                textBoxVehicleNumber.Text = "";
            }
        }    

        private void InitTicketImage()
        {
            label1.Text = "";                           // 거래처명
            label2.Text = "";                           // 일자
            label3.Text = "";                           // 입출고
            label4.Text = "";                           // 계근번호
            label5.Text = "";                           // 품명
            label6.Text = "";                           // 차량번호
            label7.Text = "";                           // 검수자
            label8.Text = "";                           // 총중량
            label9.Text = "";                           // 공차중량
            label10.Text = "";                          // 실중량
            label11.Text = "";                          // 등급명
            label12.Text = "";                          // 감량
            label13.Text = "";                          // 감량사유
            pictureBoxIn1.BackgroundImage = null;       // 입차이미지 - 1
            pictureBoxIn2.BackgroundImage = null;       //           - 2
            pictureBoxIn3.BackgroundImage = null;       //           - 3
            pictureBoxOut1.BackgroundImage = null;      // 출차이미지 - 1
            pictureBoxOut2.BackgroundImage = null;      //           - 2
            pictureBoxOut3.BackgroundImage = null;      //           - 3
        }
        #endregion
        
        #region[FTP 관련 외]
        
        //FTP 
        void OnFtpTimer(object state)
        {
            _ftpTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite);

            bool error = false;

            string[] uploadFiles = Directory.GetFiles(string.Concat(Application.StartupPath, @"\image\upload"), "*.jpg");
            if (uploadFiles.Length > 0)
            {
                foreach (string file in uploadFiles)
                {

                    DateTime fileDt = File.GetCreationTime(file);

                    if (!CreateFtpUploadDirectory(fileDt))
                    {
                        error = true;
                        break;
                    }

                    if (!_ftpClient.FtpUpload(_ftpClient.Host, _currentFtpLocation, file))
                    {
                        error = true;
                        break;
                    }

                    Log.AddLog(string.Concat("FTP Upload.. ", Path.GetFileName(file)));

                    // 업로드 완료되었으므로 해당파일 삭제
                    File.Delete(file);
                }
            }

            if (!_shutdown)
            {
                if (error)
                    _ftpTimer.Change(30000, System.Threading.Timeout.Infinite);
                else
                    _ftpTimer.Change(5000, System.Threading.Timeout.Infinite);
            }
        }
        
        private bool CreateFtpUploadDirectory(DateTime dt)
        {
            if (!_ftpClient.FtpDirectioryCheck(string.Concat(_ftpUploadPath, "/", dt.ToString("yyyy"))))
                return false;

            if (!_ftpClient.FtpDirectioryCheck(string.Concat(_ftpUploadPath, "/", dt.ToString("yyyy"), "/", dt.ToString("MM"))))
                return false;

            if (!_ftpClient.FtpDirectioryCheck(string.Concat(_ftpUploadPath, "/", dt.ToString("yyyy"), "/", dt.ToString("MM"), "/", dt.ToString("yyyy-MM-dd"))))
                return false;

            _currentFtpLocation = string.Concat(_ftpUploadPath, "/", dt.ToString("yyyy"), "/", dt.ToString("MM"), "/", dt.ToString("yyyy-MM-dd"));

            return true;
        }

        public byte[] ImageToByteArray(System.Drawing.Image imageIn)
        {
            using (var ms = new MemoryStream())
            {
                imageIn.Save(ms, imageIn.RawFormat);
                return ms.ToArray();
            }
        }
        /// <summary>
        /// FTP 경로의 디렉토리를 점검하고 없으면 생성
        /// </summary>
        /// <param name="directoryPath">디렉터리 경로 입니다.</param>
        public void FTPDirectioryCheck(string directoryPath, string _FTPuserID, string _FTPpassword)
        {
            string[] directoryPaths = directoryPath.Split(new string[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            string[] result = new string[directoryPaths.Length - 1];
            for (int i = 0; i < result.Length; i++)
            {
                if (i == 0)
                {
                    result[i] = directoryPaths[i] + "//" + directoryPaths[i + 1];
                }
                else
                {
                    result[i] = directoryPaths[i + 1];
                }
            }

            string currentDirectory = string.Empty;
            foreach (string directory in result)
            {
                currentDirectory += string.Format("{0}/", directory);
                if (!IsExistDirectory(currentDirectory, _FTPuserID, _FTPpassword))
                {
                    MakeDirectory(currentDirectory, _FTPuserID, _FTPpassword);
                }
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

        private bool MakeDirectory(string Directory, string _FTPuserID, string _FTPpassword)
        {
            string URI = Directory;

            System.Net.FtpWebRequest ftp = WebRequest.Create(new Uri(URI)) as FtpWebRequest;
            ftp.Credentials = new NetworkCredential(_FTPuserID, _FTPpassword);
            ftp.UseBinary = true;
            ftp.UsePassive = true;
            ftp.Timeout = 10000;
            ftp.Method = System.Net.WebRequestMethods.Ftp.MakeDirectory;

            try
            {
                string str = GetStringResponse(ftp);
            }
            catch (Exception ex)
            {
                return false;
            }
            return true;
        }

        private string GetStringResponse(FtpWebRequest ftp)
        {
            string result = "";
            using (FtpWebResponse response = (FtpWebResponse)ftp.GetResponse())
            {
                long size = response.ContentLength;
                using (Stream datastream = response.GetResponseStream())
                {
                    if (datastream != null)
                    {
                        using (StreamReader sr = new StreamReader(datastream))
                        {
                            result = sr.ReadToEnd();
                            sr.Close();
                        }

                        datastream.Close();
                    }
                }

                response.Close();
            }

            return result;
        }
        /// <summary>
        /// FTP 경로에 Upload
        /// </summary>
        /// <param name="directoryPath">디렉터리 경로 입니다.</param>
        public void FTPUpload(string directoryPath, string _FTPuserID, string _FTPpassword, byte[] data)
        {
            //업로드 위한 설정
            FtpWebRequest req = (FtpWebRequest)WebRequest.Create(directoryPath);
            req.Method = WebRequestMethods.Ftp.UploadFile;
            req.Credentials = new NetworkCredential(_FTPuserID, _FTPpassword);
            req.UsePassive = false;
            // RequestStream에 데이터를 쓴다
            req.ContentLength = data.Length;
            Stream reqStream = req.GetRequestStream();
            reqStream.Write(data, 0, data.Length);
            reqStream.Close();

            FtpWebResponse response = (FtpWebResponse)req.GetResponse();
            response.Close();
        }
        #endregion

        #region[그 외]
        private void pictureBoxDJ_DoubleClick(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                using (SettingForm form = new SettingForm())
                    form.ShowDialog();
            }
        }
        private void panelExit_DoubleClick(object sender, EventArgs e)
        {
            // 대지에스텍 요청으로 특정화면 특정위치를 더블클릭하면 종료진행
            // - 입고 -> 2차계량버튼 우측하단부분
            if (_weighingType == WeighingType.In)
            {
                this.Close();
            }
        }
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = DialogResult.None;
            result = MessageBox.Show
            (new Form { TopMost = true },
                "프로그램을 종료하시겠습니까?",
                "대지에스텍 계량 시스템",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2
            );

            if (result != DialogResult.Yes)
            {
                e.Cancel = true;
                return;
            }
            Log.AddLog("프로그램이 종료됩니다 - 사용자 종료");
        }
        private void c1DockingTabPage4_Click(object sender, EventArgs e)
        {
            PictureBox pic = new PictureBox();
            Button btn = new Button();
            if (!e.Equals(btn) || !e.Equals(pic))
            {
                textBoxCustomer.Focus();
                //keybd_event(null, 0, 0, 0);
                //keybd_event(null, 0, 0x02, 0);
                ShiftClear();
            }
        }
        private static DateTime Delay(int MS)
        {
            DateTime ThisMoment = DateTime.Now;
            TimeSpan duration = new TimeSpan(0, 0, 0, 0, MS);
            DateTime AfterWards = ThisMoment.Add(duration);
            while (AfterWards >= ThisMoment)
            {
                System.Windows.Forms.Application.DoEvents(); ThisMoment = DateTime.Now;
            }
            return DateTime.Now;
        }
        #endregion

        #region[CCTV관련]
        //CCTV
        private volatile bool _shouldStop;
        private void MONITERING()
        {
            RTSP_Init();
            StartAll();
        }

        // 카메라 RTSP 정보 불러오기 메서드
        private void RTSP_Init()
        {
            RSTPaddr1 = "rtsp://fa356df5:caps%40112@192.168.1.241/h264";
            RSTPaddr2 = "rtsp://fa356df5:caps%40112@192.168.1.242/h264";
            RSTPaddr3 = "rtsp://fa356df5:caps%40112@192.168.1.244/h264";
        }

        #region [CCTV 영상 가져오기 쓰레드]
        // 쓰레드 시작 메서드
        private void StartAll()
        {
            if (t_Camera1 == null || !t_Camera1.IsAlive)
            {
                t_Camera1 = new Thread(new ThreadStart(StartCamera1));
                t_Camera1.IsBackground = true;
                t_Camera1.Start();
            }
            if (t_Camera2 == null || !t_Camera2.IsAlive)
            {
                t_Camera2 = new Thread(new ThreadStart(StartCamera2));
                t_Camera2.IsBackground = true;
                t_Camera2.Start();
            }
            if (t_Camera3 == null || !t_Camera3.IsAlive)
            {
                t_Camera3 = new Thread(new ThreadStart(StartCamera3));
                t_Camera3.IsBackground = true;
                t_Camera3.Start();
            }
            //if (_T_CONFIRM == null || !_T_CONFIRM.IsAlive)
            //{
            //    _T_CONFIRM = new Thread(new ThreadStart(AutoComplete));
            //    _T_CONFIRM.IsBackground = true;
            //    _T_CONFIRM.Start();
            //}
        }

        // 1번 카메라 쓰레드 (반복)
        private void StartCamera1()
        {
            try
            {
                while (!_shouldStop)
                {
                    try
                    {
                        lock(_lockCamera1)
                        {
                            if (!isConnected)
                            {
                                string RSTPaddr = RSTPaddr1;
                                capture1 = new VideoCapture(RSTPaddr);
                                frame1 = new Mat();
                                isConnected = true;
                            }

                            if (!capture1.Read(frame1))
                            {
                                Cv2.WaitKey();
                            }
                            pictureBox14.Image = BitmapConverter.ToBitmap(frame1);
                            Process pro = Process.GetCurrentProcess(); // 현재 프로세스 사용량
                            long memory = pro.VirtualMemorySize64;

                            // 500mb
                            if (memory > 1000000000)
                            {
                                //isConnected = false;
                                //capture1.Release();
                            }

                            if (Cv2.WaitKey(1) >= 0)
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        isConnected = false;
                    }
                    Thread.Sleep(1);
                }
            }
            catch
            {
            }
        }
        public Bitmap ResizeBitmap(Bitmap bmp, int width, int height)
        {
            Bitmap result = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(result))
            {
                g.DrawImage(bmp, 0, 0, width, height);
            }

            return result;
        }
        // 2번 카메라 쓰레드 (반복)
        private void StartCamera2()
        {
            try
            {
                while (!_shouldStop)
                {
                    try
                    {
                        lock (_lockCamera2)
                        {
                            if (!isConnected)
                            {
                                string RSTPaddr = RSTPaddr2;
                                capture2 = new VideoCapture(RSTPaddr);
                                frame2 = new Mat();
                                isConnected = true;
                            }

                            if (!capture2.Read(frame2))
                            {
                                Cv2.WaitKey();
                            }
                            pictureBox13.Image = BitmapConverter.ToBitmap(frame2);
                            Process pro = Process.GetCurrentProcess(); // 현재 프로세스 사용량
                            long memory = pro.VirtualMemorySize64;

                            // 500mb
                            if (memory > 1000000000)
                            {
                                //isConnected = false;
                                // capture2.Release();
                            }
                            if (Cv2.WaitKey(1) >= 0)
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        isConnected = false;
                    }
                    Thread.Sleep(1);
                }
            }
            catch
            {
            }
        }
        private void StartCamera3()
        {
            try
            {
                while (!_shouldStop)
                {
                    try
                    {
                        lock (_lockCamera3)
                        {
                            if (!isConnected)
                            {
                                string RSTPaddr = RSTPaddr3;
                                capture3 = new VideoCapture(RSTPaddr);
                                frame3 = new Mat();
                                isConnected = true;
                            }

                            if (!capture3.Read(frame3))
                            {
                                Cv2.WaitKey();
                            }
                            pictureBox12.Image = BitmapConverter.ToBitmap(frame3);
                            Process pro = Process.GetCurrentProcess(); // 현재 프로세스 사용량
                            long memory = pro.VirtualMemorySize64;

                            // 500mb
                            if (memory > 1000000000)
                            {
                                //isConnected = false;
                                //capture3.Release();
                            }
                            if (Cv2.WaitKey(1) >= 0)
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        isConnected = false;
                    }
                    Thread.Sleep(1);
                }
            }
            catch
            {
            }
        }
        #endregion
        
        #region [크로스 스레드]
        delegate void CrossThreadSafetySetText(Control ctl, object obj);
        private void CSafeSetText(Control ctl, object obj)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetText(CSafeSetText), ctl, obj);
            else
                ctl.Text = obj.ToString();
        }

        delegate void CrossThreadSafetySetForeColor(Control ctl, Color color);
        private void CSafeSetForeColor(Control ctl, Color color)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetForeColor(CSafeSetForeColor), ctl, color);
            else
                ctl.ForeColor = color;
        }

        delegate void CrossThreadSafetySetBackgroundImage(Control ctl, Image img);
        private void CSafeSetBackgroundImage(Control ctl, Image img)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetBackgroundImage(CSafeSetBackgroundImage), ctl, img);
            else
                ctl.BackgroundImage = img;
        }

        delegate void CrossThreadSafetySetLocation(Control ctl, System.Drawing.Point pt);
        private void CSafeSetLocation(Control ctl, System.Drawing.Point pt)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetLocation(CSafeSetLocation), ctl, pt);
            else
                ctl.Location = pt;
        }

        delegate void CrossThreadSafetySetImage(PictureEdit ctl, Image img);
        private void CSafeSetImage(PictureEdit ctl, Image img)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetImage(CSafeSetImage), ctl, img);
            else
                ctl.Image = img;
        }

        delegate void CrossThreadSafetySetImageBox(PictureBox ctl, Image img);
        private void CSafeSetImageBox(PictureBox ctl, Image img)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetImageBox(CSafeSetImageBox), ctl, img);
            else
                ctl.Image = img;
        }
        delegate void CrossThreadSafetySetTextMemo(MemoEdit ctl, string str);
        private void CSafeSetmemoEdit(MemoEdit ctl, string str)
        {
            if (ctl.InvokeRequired)
            {
                ctl.Invoke(new CrossThreadSafetySetTextMemo(CSafeSetmemoEdit), ctl, str);
            }
            else
            {
                ctl.Text = str;
            }
        }
        #endregion

        #region [프로그램 종료]
        private void label1_Click(object sender, EventArgs e)
        {
            _shouldStop = true;
            is_AutoComplete = false;
            Dispose();
        }
        #endregion

        #region [버튼이벤트]
        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnCapture_Click_1(object sender, EventArgs e)
        {
            SelectTabPage(0);
        }

        #endregion

        private void button1_Click_1(object sender, EventArgs e)
        {
            SelectTabPage(6);
        }
        #endregion

        // 프로그램 실행 시, FTP 임시폴더 검사후, 잔류파일 삭제 (현재기준 : -30일)
        private void InitFTPFile()
        {
            string appPath = Application.StartupPath;
            string ImagePath = string.Concat(appPath, @"\image\");             // 이미지파일 경로
            string PastDate = DateTime.Now.AddDays(-30).ToString("yyyy-MM-dd"); // 프로그램 실행일자 -30일
            string error = string.Empty;

            try
            {
                // 계근일자가 한 달 전인 전표번호의 최댓값 조회
                StringBuilder strSql = new StringBuilder();
                strSql.Clear();
                strSql.AppendFormat(" SELECT MAX(JUNPYOID) AS JUNPYOID                      ");
                strSql.AppendFormat(" FROM   MESURING                                       ");
                strSql.AppendFormat(" WHERE  J_DATE = '" + PastDate + "'                    ");
                DataSet ds = OpenSnapshot(strSql.ToString(), out error);
                QuitwithLog(error);
                if (ds != null)
                {
                    int rowCount = ds.Tables[0].Rows.Count;
                    if (rowCount > 0)
                    {
                        DataTable table = ds.Tables[0];
                        DataRowCollection row = table.Rows;
                        if (table.Rows.Count > 0)
                        {
                            // -30일 기준 전표번호 최댓값
                            string MaxID = table.Rows[0]["JUNPYOID"]?.ToString();

                            // 빈 값이면 아무 작업없이 로직 종료
                            if (!string.IsNullOrEmpty(MaxID))
                            {
                                // 폴더경로에 전표번호 최댓값보다 작은 전표번호의 이미지 비교 후 삭제
                                DirectoryInfo di = new DirectoryInfo(ImagePath);
                                // 폴더경로의 각각의 파일 검사
                                foreach (FileInfo File in di.GetFiles())
                                {
                                    int junpyo_id = Convert.ToInt32(File.Name.Substring(0, 6));

                                    if (Convert.ToInt32(MaxID) >= junpyo_id)
                                        File.Delete();
                                }
                            }
                        }
                    }
                }
            }
            catch
            {

            }
        }
    }
}
