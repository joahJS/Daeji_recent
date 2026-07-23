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
using System.Threading;
using OpenCvSharp;
using OpenCvSharp.Extensions;
using Daeji_MONITERING.ComnLib;
using DevExpress.XtraLayout;
/*
 * 수정일자 : 2023-02-06
 * 수정자   : 정은영
 * ID       : #0001
 * 수정내용 : (현업요청)
 *            1. 품목 기본값 G S로 변경
 *            2. 모니터링 프로그램 뻗음 현상 해결
 *            
 */
namespace Daeji_MONITERING
{
    public partial class DJMONITERING : DevExpress.XtraEditors.XtraForm
    {
        public DJMONITERING()
        {
            InitializeComponent();
        }

        private volatile bool _shouldStop;
        private string PROCEDURE_ID = "DP_MONITERING";
        private string PROCEDURE_AL = "DP_MONIALARM";
        #region [ CCTV 영상 변수 ]
        Mat frame1;
        Mat frame2;
        #endregion
        #region [ 쓰레드 ]
        //cctv
        Thread t_Camera1;
        Thread t_Camera2;

        //차지카운트 쓰레드
        Thread t_Charg;

        //1공정진행도
        Thread t_Progress;

        //오일량
        Thread t_oil;

        //알람
        Thread t_Alram;
        #endregion
        #region [ 녹화 및 저장 변수 ]
        VideoCapture capture1;
        VideoCapture capture2;
        #endregion
        #region [ 조건 변수 ]
        //카메라관련
        bool isFirstCamera;
        bool isConnected;

        //버튼선택용
        string _EMPID = string.Empty;
        string _GUBUN = string.Empty;

        //설비별 과전류 유지초
        Dictionary<Label, DateTime> JunList = new Dictionary<Label, DateTime>();
        Dictionary<Label, string> JunMSGList = new Dictionary<Label, string>();

        //설비별 압력 유지초
        Dictionary<Label, DateTime> ApList = new Dictionary<Label, DateTime>();

        #region [ CCTV 연결 시작/종료 구분용 변수 ]
        private bool _isWaitStart_1 = true;
        private bool _isWaitStart_2 = true;
        #endregion
        #endregion
        #region [ 알람메세지 변수 ]
        //알람창 display
        public enum ALMDISPALY { ON, OFF }
        public ALMDISPALY _ALMDISPALY = ALMDISPALY.ON;

        //경고창 전체알람
        public string _WARNINGMSG = string.Empty;
        #endregion
        #region [ 화면전환변수 ]
        public enum CCTV_GB { CCTV, PRINT}
        public CCTV_GB _CCTV_GB = CCTV_GB.CCTV;
        #endregion
        #region [ ini 값 변수 ]
        private string filePath = ComnString.INI_FILE_PATH;
        private IniUtil _ini = null;

        private string _CAM1 = string.Empty; //1카메라 RTSP  //rtsp://caps1:caps%40112@192.168.1.198/stream1
        private string _CAM2 = string.Empty; //2카메라 RTSP  //rtsp://caps1:caps%40112@192.168.1.221/h264


        private double _OilOndo = 0; //오일온도 경고 설정값
        private double _ApRyuk = 0; //압력 경고 설정값
        private double _JunRu = 0; //전류 경고 설정값
        private double _Time1 = 0; //전류 유지 초 설정값
        #endregion

        #region [ timer ]
        System.Windows.Forms.Timer timerCharge = new System.Windows.Forms.Timer();
        System.Windows.Forms.Timer timeProgress = new System.Windows.Forms.Timer();
        System.Windows.Forms.Timer timerOil = new System.Windows.Forms.Timer();
        #endregion

        private void MONITERING_Shown(object sender, EventArgs e)
        {
            StartAll();
        }

        private void MONITERING_Load(object sender, EventArgs e)
        {
            ReadSettingData();//ini값 세팅
            SetJunMSGList();//과전류 메세지 설정

            //작업자세팅
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "MESWORKER");
            DataTable dt2 = DBConn.GetDataTable(this.PROCEDURE_ID, dicParams);

            if (dt2.Rows.Count > 0)
            {
                for (int i = 0; i< dt2.Rows.Count; i++)
                {
                    // 버튼의 텍스트 설정
                    switch (i)
                    {
                        case 0:
                            Btnworker1.Text = dt2.Rows[i]["EMP_NM"]?.ToString();
                            break;
                        case 1:
                            Btnworker2.Text = dt2.Rows[i]["EMP_NM"]?.ToString();
                            break;
                        case 2:
                            Btnworker3.Text = dt2.Rows[i]["EMP_NM"]?.ToString();
                            break;
                        case 3:
                            Btnworker4.Text = dt2.Rows[i]["EMP_NM"]?.ToString();
                            break;
                        case 4:
                            Btnworker5.Text = dt2.Rows[i]["EMP_NM"]?.ToString();
                            break;
                        case 5:
                            Btnworker6.Text = dt2.Rows[i]["EMP_NM"]?.ToString();
                            break;
                        default:
                            break;
                    }
                }
            }

            if (string.IsNullOrEmpty(Btnworker1.Text))
            {
                Btnworker1.Enabled = false;
            }
            if (string.IsNullOrEmpty(Btnworker2.Text))
            {
                Btnworker2.Enabled = false;
            }
            if (string.IsNullOrEmpty(Btnworker3.Text))
            {
                Btnworker3.Enabled = false;
            }
            if (string.IsNullOrEmpty(Btnworker4.Text))
            {
                Btnworker4.Enabled = false;
            }
            if (string.IsNullOrEmpty(Btnworker5.Text))
            {
                Btnworker5.Enabled = false;
            }
            if (string.IsNullOrEmpty(Btnworker6.Text))
            {
                Btnworker6.Enabled = false;
            }

            
            dicParams.Clear();
            dicParams.Add("CMD", "GETPERSON");

            DataTable dt = DBConn.GetDataTable(this.PROCEDURE_ID, dicParams);

            if(dt != null && dt.Rows.Count > 0)
            {
                _EMPID = dt.Rows[0]["MPERSN"]?.ToString();
                _GUBUN = dt.Rows[0]["MGUBUN"]?.ToString();

                SetWorkerBtn(_EMPID);
                SetGubBtn(_GUBUN);
            }
            
            
            
        }

        private void SetJunMSGList()
        {
            //설비별 메세지 설정
            JunMSGList.Clear();
            JunMSGList.Add(La_J1, "설비 1 과전류");
            JunMSGList.Add(La_J2, "설비 2 과전류");
            JunMSGList.Add(La_J3, "설비 3 과전류");
            JunMSGList.Add(La_J4, "설비 4 과전류");
            JunMSGList.Add(La_J5, "설비 5 과전류");
            JunMSGList.Add(La_J6, "설비 6 과전류");
            JunMSGList.Add(La_J7, "설비 7 과전류");
            JunMSGList.Add(La_J8, "설비 8 과전류");
            JunMSGList.Add(La_J9, "설비 9 과전류");
            JunMSGList.Add(La_J10, "설비 10 과전류");

            DateTime now = DateTime.Now;

            //설비별 경고값 초과 시간설정
            JunList.Clear();
            JunList.Add(La_J1, now);
            JunList.Add(La_J2, now);
            JunList.Add(La_J3, now);
            JunList.Add(La_J4, now);
            JunList.Add(La_J5, now);
            JunList.Add(La_J6, now);
            JunList.Add(La_J7, now);
            JunList.Add(La_J8, now);
            JunList.Add(La_J9, now);
            JunList.Add(La_J10, now);

            //설비별 압력 경고값 초과 시간설정
            ApList.Clear();
            ApList.Add(La_P1, now);
            ApList.Add(La_P2, now);
            ApList.Add(La_P3, now);
            ApList.Add(La_P4, now);
            ApList.Add(La_P5, now);
            ApList.Add(La_P6, now);
            ApList.Add(La_P7, now);
            ApList.Add(La_P8, now);
            ApList.Add(La_P9, now);
            ApList.Add(La_P10, now);
        }

        //ini값 세팅
        private void ReadSettingData()
        {
            _ini = new IniUtil(filePath);

            _CAM1 = _ini.GetIniValue("CAM", "CAM1");
            _CAM2 = _ini.GetIniValue("CAM", "CAM2");

            double.TryParse(_ini.GetIniValue("MAXNUM", "OILONDO"), out _OilOndo);
            double.TryParse(_ini.GetIniValue("MAXNUM", "APRYUK"), out _ApRyuk);
            double.TryParse(_ini.GetIniValue("MAXNUM", "JUNRU"), out _JunRu);
            double.TryParse(_ini.GetIniValue("SECTIME", "TIME1"), out _Time1);
        }

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
            //if (t_Progress == null || !t_Progress.IsAlive)
            //{
            //    t_Progress = new Thread(new ThreadStart(StartProgress));
            //    t_Progress.IsBackground = true;
            //    t_Progress.Start();
            //}
            //if (t_oil == null || !t_oil.IsAlive)
            //{
            //    t_oil = new Thread(new ThreadStart(StratOil));
            //    t_oil.IsBackground = true;
            //    t_oil.Start();
            //}
            //알람메세지 제거
            //if(t_Alram == null || !t_Alram.IsAlive)
            //{
            //    t_Alram = new Thread(new ThreadStart(StartAlram));
            //    t_Alram.IsBackground = true;
            //    t_Alram.Start();
            //}
            //if(t_Charg == null || !t_Charg.IsAlive)
            //{
            //    t_Charg = new Thread(new ThreadStart(StartCharg));
            //    t_Charg.IsBackground = true;
            //    t_Charg.Start();
            //}

            //#0001 타이머로 변경
            timerCharge.Interval = 100;
            timerCharge.Tick += new EventHandler(StartCharg);
            timerCharge.Start();

            timeProgress.Interval = 100;
            timeProgress.Tick += new EventHandler(StartProgress);
            timeProgress.Start();

            timerOil.Interval = 100;
            timerOil.Tick += new EventHandler(StratOil);
            timerOil.Start();
        }

        #region [ CCTV 영상 가져오기 쓰레드 ]
        // 1번 카메라 쓰레드 (반복)
        private void StartCamera1()
        {
            try
            {
                while (!_shouldStop)
                {
                    if(_CCTV_GB == CCTV_GB.CCTV)
                    {
                        try
                        {
                            if (!isConnected)
                            {
                                capture1 = new VideoCapture(_CAM1);
                                frame1 = new Mat();
                                isConnected = true;
                            }

                            if (!isFirstCamera)
                            {
                                isFirstCamera = true;
                            }

                            if (!capture1.Read(frame1))
                            {
                                Cv2.WaitKey();
                            }

                            Bitmap bitmap2 = BitmapConverter.ToBitmap(frame1);
                            PicCCTV1.Image = bitmap2;

                            if (_isWaitStart_1)
                            {
                                //ComnMethod.SetLogInfo(Name, Text, "1번 카메라 연결");
                                _isWaitStart_1 = false;
                            }

                            if (Cv2.WaitKey(1) >= 0)
                                break;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            isConnected = false;
                            if (!_isWaitStart_1)
                            {
                                //ComnMethod.SetLogInfo(Name, Text, "1번 카메라 단절");
                                _isWaitStart_1 = true;
                            }
                        }
                        Thread.Sleep(1);
                    }
                }
            }
            catch(Exception ex)
            {
                //CSafeSetmemoEdit(memoEdit1, "1번카메라 쓰레드 오류 : " + ex.Message);
            }
        }

        // 2번 카메라 쓰레드 (반복)
        private void StartCamera2()
        {
            try
            {
                while (!_shouldStop)
                {
                    if(_CCTV_GB == CCTV_GB.CCTV)
                    {
                        try
                        {
                            if (!isConnected)
                            {
                                capture2 = new VideoCapture(_CAM2);
                                frame2 = new Mat();
                                isConnected = true;
                            }

                            if (!isFirstCamera)
                            {
                                isFirstCamera = true;
                            }

                            if (!capture2.Read(frame2))
                            {
                                Cv2.WaitKey();
                            }

                            Bitmap bitmap2 = BitmapConverter.ToBitmap(frame2);
                            PicCCTV2.Image = bitmap2;

                            if (_isWaitStart_2)
                            {
                                //ComnMethod.SetLogInfo(Name, Text, "1번 카메라 연결");
                                _isWaitStart_2 = false;
                            }

                            if (Cv2.WaitKey(1) >= 0)
                                break;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            isConnected = false;
                            if (!_isWaitStart_2)
                            {
                                //ComnMethod.SetLogInfo(Name, Text, "1번 카메라 단절");
                                _isWaitStart_2 = true;
                            }
                        }
                        Thread.Sleep(1);
                    }
                }
            }
            catch(Exception ex)
            {
                //CSafeSetmemoEdit(memoEdit1, "2번카메라 쓰레드 오류 : " + ex.Message);
            }
        }
        #endregion

        #region [ 차지카운트 쓰레드 ]
        private void StartCharg(object sender, System.EventArgs e)
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST4");

                //while (true)
                //{
                    DataTable dt = DBConn.GetDataTable(this.PROCEDURE_ID, dicParams);

                    if (dt != null && dt.Rows.Count > 0)
                    {
                        string sCHARG = dt.Rows[0]["CHARG"]?.ToString();

                        CSafeSetLabel(La_Charg, sCHARG);
                    }
                    //Thread.Sleep(500);
                //}
            }
            catch (Exception ex)
            {
            }
        }
        #endregion

        #region [ 1공정 진행도 쓰레드 ]
        double dCOVER_OP = 0;//커버열림
        double dCOVER_CL = 0;//커버닫힘
        double dP_CLIMB = 0;//압축상승
        double dP_DESCENT = 0;//압축하강
        double dCUT_CLIMB = 0;//절단상승
        double dCUT_DESCENT = 0;//절단하강
        double dFOR_TRANS = 0;//이송전진
        double dBACK_TRANS = 0;//이송후진
        double dFOR_PRESS = 0;//횡압전진
        double dBACK_PRESS = 0;//횡압후진

        private void StartProgress(object sender, System.EventArgs e)
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST3");

                //while (true)
                //{
                    DataTable dt = DBConn.GetDataTable(this.PROCEDURE_ID, dicParams);

                    if(dt != null && dt.Rows.Count > 0)
                    {
                        string sWORK_PRC = dt.Rows[0]["WORK_PRC"]?.ToString();
                        int.TryParse(sWORK_PRC, out int iVal);
                        CSafeSetProgressControl(Pgb1, iVal);

                        string sCOVER_OP     = dt.Rows[0]["COVER_OP"]?.ToString();
                        string sCOVER_CL     = dt.Rows[0]["COVER_CL"]?.ToString();
                        string sP_CLIMB      = dt.Rows[0]["P_CLIMB"]?.ToString();
                        string sP_DESCENT    = dt.Rows[0]["P_DESCENT"]?.ToString();
                        string sCUT_CLIMB    = dt.Rows[0]["CUT_CLIMB"]?.ToString();
                        string sCUT_DESCENT  = dt.Rows[0]["CUT_DESCENT"]?.ToString();
                        string sFOR_TRANS    = dt.Rows[0]["FOR_TRANS"]?.ToString();
                        string sBACK_TRANS   = dt.Rows[0]["BACK_TRANS"]?.ToString();
                        string sFOR_PRESS    = dt.Rows[0]["FOR_PRESS"]?.ToString();
                        string sBACK_PRESS   = dt.Rows[0]["BACK_PRESS"]?.ToString();

                        double dCOVER_OP_1 = 0;
                        double dCOVER_CL_1 = 0;
                        double dP_CLIMB_1 = 0;
                        double dP_DESCENT_1 = 0;
                        double dCUT_CLIMB_1 = 0;
                        double dCUT_DESCENT_1 = 0;
                        double dFOR_TRANS_1 = 0;
                        double dBACK_TRANS_1 = 0;
                        double dFOR_PRESS_1 = 0;
                        double dBACK_PRESS_1 = 0;

                        double.TryParse(sCOVER_OP   , out dCOVER_OP_1);
                        double.TryParse(sCOVER_CL   , out dCOVER_CL_1);
                        double.TryParse(sP_CLIMB    , out dP_CLIMB_1);
                        double.TryParse(sP_DESCENT  , out dP_DESCENT_1);
                        double.TryParse(sCUT_CLIMB  , out dCUT_CLIMB_1);
                        double.TryParse(sCUT_DESCENT, out dCUT_DESCENT_1);
                        double.TryParse(sFOR_TRANS  , out dFOR_TRANS_1);
                        double.TryParse(sBACK_TRANS , out dBACK_TRANS_1);
                        double.TryParse(sFOR_PRESS  , out dFOR_PRESS_1);
                        double.TryParse(sBACK_PRESS, out dBACK_PRESS_1);

                        //커버열림
                        if (dCOVER_OP != dCOVER_OP_1) //값 변경됨 -> 실행
                        {
                            CSafeSetImage(PicCoverOp, Properties.Resources.열기on);
                        }
                        else //값 고정 -> 종료
                        {
                            CSafeSetImage(PicCoverOp, Properties.Resources.열기off);
                        }
                        //커버닫힘
                        if (dCOVER_CL != dCOVER_CL_1) //값 변경됨 -> 실행
                        {
                            CSafeSetImage(PicCoverCl, Properties.Resources.닫기on);
                        }
                        else //값 고정 -> 종료
                        {
                            CSafeSetImage(PicCoverCl, Properties.Resources.닫기off);
                        }

                        //압축상승
                        if (dP_CLIMB != dP_CLIMB_1) //값 변경됨 -> 실행
                        {
                            CSafeSetImage(PicAUp, Properties.Resources.상승on);
                        }
                        else //값 고정 -> 종료
                        {
                            CSafeSetImage(PicAUp, Properties.Resources.상승off);
                        }
                        //압축하강
                        if (dP_DESCENT != dP_DESCENT_1) //값 변경됨 -> 실행
                        {
                            CSafeSetImage(PicADown, Properties.Resources.하강on);
                        }
                        else //값 고정 -> 종료
                        {
                            CSafeSetImage(PicADown, Properties.Resources.하강off);
                        }

                        //절단상승
                        if (dCUT_CLIMB != dCUT_CLIMB_1) //값 변경됨 -> 실행
                        {
                            CSafeSetImage(PicCUp, Properties.Resources.상승on);
                        }
                        else //값 고정 -> 종료
                        {
                            CSafeSetImage(PicCUp, Properties.Resources.상승off);
                        }
                        //절단하강
                        if (dCUT_DESCENT != dCUT_DESCENT_1) //값 변경됨 -> 실행
                        {
                            CSafeSetImage(PicCDown, Properties.Resources.하강on);
                        }
                        else //값 고정 -> 종료
                        {
                            CSafeSetImage(PicCDown, Properties.Resources.하강off);
                        }

                        //이송전진
                        if (dFOR_TRANS != dFOR_TRANS_1) //값 변경됨 -> 실행
                        {
                            CSafeSetImage(PicGo, Properties.Resources.전진on);
                            
                        }
                        else //값 고정 -> 종료
                        {
                            CSafeSetImage(PicGo, Properties.Resources.전진off);
                        }
                        //이송후진
                        if (dBACK_TRANS != dBACK_TRANS_1) //값 변경됨 -> 실행
                        {
                            CSafeSetImage(PicBack, Properties.Resources.후진on);
                        }
                        else //값 고정 -> 종료
                        {
                            CSafeSetImage(PicBack, Properties.Resources.후진off);
                        }

                        //횡압전진
                        if (dFOR_PRESS != dFOR_PRESS_1) //값 변경됨 -> 실행
                        {
                            CSafeSetImage(PicPush, Properties.Resources.밀기on);
                        }
                        else //값 고정 -> 종료
                        {
                            CSafeSetImage(PicPush, Properties.Resources.밀기off);
                        }
                        //횡압후진
                        if (dBACK_PRESS != dBACK_PRESS_1) //값 변경됨 -> 실행
                        {
                            CSafeSetImage(PicPull, Properties.Resources.원위치on);
                        }
                        else //값 고정 -> 종료
                        {
                            CSafeSetImage(PicPull, Properties.Resources.원위치off);
                        }

                        dCOVER_OP = dCOVER_OP_1;
                        dCOVER_CL = dCOVER_CL_1;
                        dP_CLIMB = dP_CLIMB_1;    
                        dP_DESCENT = dP_DESCENT_1;   
                        dCUT_CLIMB = dCUT_CLIMB_1;   
                        dCUT_DESCENT = dCUT_DESCENT_1; 
                        dFOR_TRANS = dFOR_TRANS_1;   
                        dBACK_TRANS = dBACK_TRANS_1;
                        dFOR_PRESS = dFOR_PRESS_1;  
                        dBACK_PRESS = dBACK_PRESS_1;
                    }
                    //Thread.Sleep(500);
                //}
            }
            catch (Exception ex)
            {
            }
        }
        #endregion

        #region[ 업데이트 데이터(압력,전류,오일탱크레벨,오일온도,작동유압력) 쓰레드 ]
        private void StratOil(object sender, System.EventArgs e)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "LIST1");

            //while (true)
            //{
                try
                {
                    DataTable dt = DBConn.GetDataTable(this.PROCEDURE_ID, dicParams);

                    if (dt != null && dt.Rows.Count > 0)
                    {
                        string sPRESS_1 = dt.Rows[0]["PRESS_1"]?.ToString();//압력
                        string sPRESS_2 = dt.Rows[0]["PRESS_2"]?.ToString();
                        string sPRESS_3 = dt.Rows[0]["PRESS_3"]?.ToString();
                        string sPRESS_4 = dt.Rows[0]["PRESS_4"]?.ToString();
                        string sPRESS_5 = dt.Rows[0]["PRESS_5"]?.ToString();
                        string sPRESS_6 = dt.Rows[0]["PRESS_6"]?.ToString();
                        string sPRESS_7 = dt.Rows[0]["PRESS_7"]?.ToString();
                        string sPRESS_8 = dt.Rows[0]["PRESS_8"]?.ToString();
                        string sPRESS_9 = dt.Rows[0]["PRESS_9"]?.ToString();
                        string sPRESS_10 = dt.Rows[0]["PRESS_10"]?.ToString();
                        string sCURRENT_1 = dt.Rows[0]["CURRENT_1"]?.ToString();//전류
                        string sCURRENT_2 = dt.Rows[0]["CURRENT_2"]?.ToString();
                        string sCURRENT_3 = dt.Rows[0]["CURRENT_3"]?.ToString();
                        string sCURRENT_4 = dt.Rows[0]["CURRENT_4"]?.ToString();
                        string sCURRENT_5 = dt.Rows[0]["CURRENT_5"]?.ToString();
                        string sCURRENT_6 = dt.Rows[0]["CURRENT_6"]?.ToString();
                        string sCURRENT_7 = dt.Rows[0]["CURRENT_7"]?.ToString();
                        string sCURRENT_8 = dt.Rows[0]["CURRENT_8"]?.ToString();
                        string sCURRENT_9 = dt.Rows[0]["CURRENT_9"]?.ToString();
                        string sCURRENT_10 = dt.Rows[0]["CURRENT_10"]?.ToString();
                        string sTANK_LV = dt.Rows[0]["TANK_LV"]?.ToString();//오일탱크레벨
                        string sOIL_TEMP = dt.Rows[0]["OIL_TEMP"]?.ToString();//오일온도
                        string sLINE_PRESS = dt.Rows[0]["LINE_PRESS"]?.ToString();//작동유압력

                        //오일레벨
                        int iTANK_LV = 0;
                        int.TryParse(sTANK_LV, out iTANK_LV);
                        CSafeSetImageBoxOil(PrgBar2, iTANK_LV);

                        //오일온도
                        if(int.TryParse(sOIL_TEMP, out int iOIL_TEMP))
                        {
                            CSafeSetLabel(la_OilTemp, (iOIL_TEMP / 10).ToString("n1") + "°C");
                            if (iOIL_TEMP > _OilOndo)
                            {
                                if (!_WARNINGMSG.Contains("유압오일 온도 높음"))
                                {
                                    _WARNINGMSG += "\r\n유압오일 온도 높음";
                                    SetWorningMSG(_WARNINGMSG);
                                }
                            }
                            else
                            {
                                _WARNINGMSG = _WARNINGMSG.Replace("\r\n유압오일 온도 높음", "");
                                _WARNINGMSG = _WARNINGMSG.Replace("유압오일 온도 높음", "");
                                SetWorningMSG(_WARNINGMSG);
                            }
                        }


                        //작동유 압력
                        CSafeSetLabel(la_LinePress, sLINE_PRESS+"bar");


                        //압력비교
                        //압력 경고값
                        double iPressChk = _ApRyuk;

                        Dictionary<Label, string> dicPicList = new Dictionary<Label, string>();

                        dicPicList.Clear();
                        dicPicList.Add(La_P1, sPRESS_1);
                        dicPicList.Add(La_P2, sPRESS_2);
                        dicPicList.Add(La_P3, sPRESS_3);
                        dicPicList.Add(La_P4, sPRESS_4);
                        dicPicList.Add(La_P5, sPRESS_5);
                        dicPicList.Add(La_P6, sPRESS_6);
                        dicPicList.Add(La_P7, sPRESS_7);
                        dicPicList.Add(La_P8, sPRESS_8);
                        dicPicList.Add(La_P9, sPRESS_9);
                        dicPicList.Add(La_P10, sPRESS_10);

                        foreach (KeyValuePair<Label, string> param in dicPicList)
                        {
                            if (double.TryParse(param.Value, out double iResult))
                            {
                                if (iResult > iPressChk)
                                {
                                    foreach (KeyValuePair<Label, DateTime> param2 in ApList.ToList())
                                    {
                                        if (param.Key == param2.Key)
                                        {
                                            DateTime dtVal = param2.Value;

                                            TimeSpan dateDiff = DateTime.Now - dtVal;

                                            if (dateDiff.Seconds < _Time1)
                                            {
                                                continue;
                                            }

                                            //경고값보다 높을때 빨간색으로 이미지 변경
                                            CSafeSetLabelImage(param.Key, Properties.Resources.압력_배경_on);
                                        }
                                    }
                                }
                                else
                                {
                                    foreach (KeyValuePair<Label, DateTime> param2 in ApList.ToList())
                                    {
                                        if (param.Key == param2.Key)
                                        {
                                            ApList[param2.Key] = DateTime.Now;
                                        }
                                    }

                                    //경고값 보다 낮을때 기본색으로 변경
                                    CSafeSetLabelImage(param.Key, Properties.Resources.압력배경off);
                                }

                                CSafeSetLabel(param.Key, param.Value);
                            }
                        }

                        //전류비교
                        //전류 경고값
                        double iCurrentChk = _JunRu;

                        dicPicList.Clear();
                        dicPicList.Add(La_J1, sCURRENT_1);
                        dicPicList.Add(La_J2, sCURRENT_2);
                        dicPicList.Add(La_J3, sCURRENT_3);
                        dicPicList.Add(La_J4, sCURRENT_4);
                        dicPicList.Add(La_J5, sCURRENT_5);
                        dicPicList.Add(La_J6, sCURRENT_6);
                        dicPicList.Add(La_J7, sCURRENT_7);
                        dicPicList.Add(La_J8, sCURRENT_8);
                        dicPicList.Add(La_J9, sCURRENT_9);
                        dicPicList.Add(La_J10, sCURRENT_10);

                        foreach (KeyValuePair<Label, string> param in dicPicList)
                        {
                            if (double.TryParse(param.Value, out double iResult))
                            {
                                if (iResult > iCurrentChk)
                                {
                                    foreach (KeyValuePair<Label, DateTime> param2 in JunList.ToList())
                                    {
                                        if(param.Key == param2.Key)
                                        {
                                            DateTime dtVal = param2.Value;

                                            TimeSpan dateDiff = DateTime.Now - dtVal;

                                            if (dateDiff.Seconds < _Time1)
                                            {
                                                continue;
                                            }

                                            if (!_WARNINGMSG.Contains(JunMSGList[param2.Key]))
                                            {
                                                _WARNINGMSG += "\r\n" + JunMSGList[param2.Key];
                                                SetWorningMSG(_WARNINGMSG);
                                            }

                                            //경고값보다 높을때 빨간색으로 이미지 변경
                                            CSafeSetLabelImage(param.Key, Properties.Resources.과전류on);
                                        }
                                    }
                                }
                                else
                                {
                                    foreach (KeyValuePair<Label, DateTime> param2 in JunList.ToList())
                                    {
                                        if(param.Key == param2.Key)
                                        {
                                            string sDelMsg = JunMSGList[param2.Key];

                                            _WARNINGMSG = _WARNINGMSG.Replace("\r\n" + sDelMsg,"").Trim();
                                            _WARNINGMSG = _WARNINGMSG.Replace(sDelMsg, "").Trim();
                                            SetWorningMSG(_WARNINGMSG);

                                            JunList[param2.Key] = DateTime.Now;
                                        }
                                    }

                                    //경고값 보다 낮을때 기본색으로 변경
                                    CSafeSetLabelImage(param.Key, Properties.Resources.과전류off);
                                }

                                CSafeSetLabel(param.Key, param.Value);
                            }
                            else
                            {
                                CSafeSetLabel(param.Key, "0");
                            }
                        }
                    }

                    //Thread.Sleep(100);
                }
                catch(Exception ex)
                {
                    //CSafeSetmemoEdit(memoEdit1, "상시조회오류 : " + ex.Message);
                }
            //}
        }
        #endregion

        #region [ 알람 쓰레드 ]
        private void StartAlram(object sender, System.EventArgs e)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            while (true)
            {
                try
                {
                    dicParams.Clear();
                    dicParams.Add("CMD", "LIST2");

                    DataTable dt = DBConn.GetDataTable(this.PROCEDURE_ID, dicParams);

                    if (dt != null && dt.Rows.Count > 0)
                    {
                        double dALARM0 = 0;
                        double dALARM1 = 0;
                        double dALARM2 = 0;
                        double dALARM3 = 0;
                        double dALARM4 = 0;
                        double dALARM5 = 0;
                        //double dALARM6 = 0;
                        //double dALARM7 = 0;
                        //double dALARM8 = 0;
                        //double dALARM9 = 0;
                        //double dALARM10 = 0;
                        //double dALARM11 = 0;
                        //double dALARM12 = 0;
                        //double dALARM13 = 0;

                        double.TryParse(dt.Rows[0]["ALARM0"].ToString(), out dALARM0);
                        double.TryParse(dt.Rows[0]["ALARM1"].ToString(), out dALARM1);
                        double.TryParse(dt.Rows[0]["ALARM2"].ToString(), out dALARM2);
                        double.TryParse(dt.Rows[0]["ALARM3"].ToString(), out dALARM3);
                        double.TryParse(dt.Rows[0]["ALARM4"].ToString(), out dALARM4);
                        double.TryParse(dt.Rows[0]["ALARM5"].ToString(), out dALARM5);
                        //double.TryParse(dt.Rows[0]["ALARM6"].ToString(), out dALARM6);
                        //double.TryParse(dt.Rows[0]["ALARM7"].ToString(), out dALARM7);
                        //double.TryParse(dt.Rows[0]["ALARM8"].ToString(), out dALARM8);
                        //double.TryParse(dt.Rows[0]["ALARM9"].ToString(), out dALARM9);
                        //double.TryParse(dt.Rows[0]["ALARM10"].ToString(), out dALARM10);
                        //double.TryParse(dt.Rows[0]["ALARM11"].ToString(), out dALARM11);
                        //double.TryParse(dt.Rows[0]["ALARM12"].ToString(), out dALARM12);
                        //double.TryParse(dt.Rows[0]["ALARM13"].ToString(), out dALARM13);


                        //string sMSG = memoEdit1.EditValue?.ToString();
                        string sMSG = string.Empty;

                        if (dALARM0 > 0)
                        {
                            DataTable dt_temp = ReturnAlramMsgDt("ALARM0", dALARM0);
                            sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        }
                        if (dALARM1 > 0)
                        {
                            DataTable dt_temp = ReturnAlramMsgDt("ALARM1", dALARM1);
                            sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        }
                        if (dALARM2 > 0)
                        {
                            DataTable dt_temp = ReturnAlramMsgDt("ALARM2", dALARM2);
                            sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        }
                        if (dALARM3 > 0)
                        {
                            DataTable dt_temp = ReturnAlramMsgDt("ALARM3", dALARM3);
                            sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        }
                        if (dALARM4 > 0)
                        {
                            DataTable dt_temp = ReturnAlramMsgDt("ALARM4", dALARM4);
                            sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        }
                        if (dALARM5 > 0)
                        {
                            DataTable dt_temp = ReturnAlramMsgDt("ALARM5", dALARM5);
                            sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        }

                        #region 6번이후 알람 주석처리
                        //if (dALARM6 > 0)
                        //{
                        //    DataTable dt_temp = ReturnAlramMsgDt("ALARM6", dALARM6);
                        //    sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        //}
                        //if (dALARM7 > 0)
                        //{
                        //    DataTable dt_temp = ReturnAlramMsgDt("ALARM7", dALARM7);
                        //    sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        //}
                        //if (dALARM8 > 0)
                        //{
                        //    DataTable dt_temp = ReturnAlramMsgDt("ALARM8", dALARM8);
                        //    sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        //}
                        //if (dALARM9 > 0)
                        //{
                        //    DataTable dt_temp = ReturnAlramMsgDt("ALARM9", dALARM9);
                        //    sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        //}
                        //if (dALARM10 > 0)
                        //{
                        //    DataTable dt_temp = ReturnAlramMsgDt("ALARM10", dALARM10);
                        //    sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        //}
                        //// 알람11, 12는 버튼 상태 변경 추가
                        //if (dALARM11 > 0)
                        //{
                        //    DataTable dt_temp = ReturnAlramMsgDt("ALARM11", dALARM11);
                        //
                        //    if (dt_temp == null && dt_temp.Rows.Count <= 0)
                        //        continue;
                        //
                        //    //버튼 상태변화
                        //    dt_temp.AcceptChanges();//이전변경 데이터 커밋
                        //    foreach (DataRow row in dt_temp.Rows)
                        //    {
                        //        string sAnum2 = row["ANUM2"]?.ToString();
                        //
                        //        if (string.IsNullOrEmpty(sAnum2))
                        //            continue;
                        //
                        //        if (sAnum2.Equals("2"))//덮개 열림
                        //        {
                        //            CSafeSetImage(PicCoverOp, Properties.Resources.열기_on);
                        //            CSafeSetImage(PicCoverCl, Properties.Resources.닫기_off);
                        //
                        //            row.Delete();//버튼활성/비활성 관련 알람은 알람메세지에 표시 안됨
                        //        }
                        //
                        //        if (sAnum2.Equals("4"))//덮개 닫힘
                        //        {
                        //            CSafeSetImage(PicCoverOp, Properties.Resources.열기_off);
                        //            CSafeSetImage(PicCoverCl, Properties.Resources.닫기_on);
                        //
                        //            row.Delete();
                        //        }
                        //
                        //        if (sAnum2.Equals("16"))//압축상승
                        //        {
                        //            CSafeSetImage(PicAUp, Properties.Resources.상승_on);
                        //            CSafeSetImage(PicADown, Properties.Resources.하강_off);
                        //
                        //            row.Delete();
                        //        }
                        //
                        //        if (sAnum2.Equals("32"))//압축하강
                        //        {
                        //            CSafeSetImage(PicAUp, Properties.Resources.상승_off);
                        //            CSafeSetImage(PicADown, Properties.Resources.하강_on);
                        //
                        //            row.Delete();
                        //        }
                        //    }
                        //    //삭제한 데이터 커밋
                        //    dt_temp.AcceptChanges();
                        //
                        //    sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        //}
                        //if (dALARM12 > 0)
                        //{
                        //    DataTable dt_temp = ReturnAlramMsgDt("ALARM12", dALARM12);
                        //
                        //    //버튼상태변화
                        //    dt_temp.AcceptChanges();//이전변경 데이터 커밋
                        //    foreach (DataRow row in dt_temp.Rows)
                        //    {
                        //        string sAnum2 = row["ANUM2"]?.ToString();
                        //
                        //        if (string.IsNullOrEmpty(sAnum2))
                        //            continue;
                        //
                        //        if (sAnum2.Equals("2"))//절단 상승
                        //        {
                        //            CSafeSetImage(PicCUp, Properties.Resources.상승_on);
                        //            CSafeSetImage(PicCDown, Properties.Resources.하강_off);
                        //
                        //            row.Delete();
                        //        }
                        //
                        //        if (sAnum2.Equals("4"))//절단 하강
                        //        {
                        //            CSafeSetImage(PicCUp, Properties.Resources.상승_off);
                        //            CSafeSetImage(PicCDown, Properties.Resources.하강_on);
                        //
                        //            row.Delete();
                        //        }
                        //
                        //        if (sAnum2.Equals("32"))//이송 전진
                        //        {
                        //            CSafeSetImage(PicGo, Properties.Resources.전진_on);
                        //            CSafeSetImage(PicBack, Properties.Resources.후진_off);
                        //
                        //            row.Delete();
                        //        }
                        //
                        //        if (sAnum2.Equals("64"))//이송 후진
                        //        {
                        //            CSafeSetImage(PicGo, Properties.Resources.전진_off);
                        //            CSafeSetImage(PicBack, Properties.Resources.후진_on);
                        //
                        //            row.Delete();
                        //        }
                        //
                        //        if (sAnum2.Equals("2048"))//횡압 전진
                        //        {
                        //            CSafeSetImage(PicPush, Properties.Resources.밀기_on);
                        //            CSafeSetImage(PicPull, Properties.Resources.원위치_off);
                        //
                        //            row.Delete();
                        //        }
                        //
                        //        if (sAnum2.Equals("4096"))//횡압 후진
                        //        {
                        //            CSafeSetImage(PicPush, Properties.Resources.밀기_off);
                        //            CSafeSetImage(PicPull, Properties.Resources.원위치_on);
                        //
                        //            row.Delete();
                        //        }
                        //
                        //    }
                        //    //삭제한 데이터 커밋
                        //    dt_temp.AcceptChanges();
                        //
                        //    sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        //}
                        ////
                        //if (dALARM13 > 0)
                        //{
                        //    DataTable dt_temp = ReturnAlramMsgDt("ALARM13", dALARM13);
                        //    sMSG += ReturnAlramMsg(dt_temp, sMSG);
                        //}
                        #endregion

                        //CSafeSetmemoEdit(memoEdit1, sMSG.Trim());
                    }

                    Thread.Sleep(500);
                }
                catch (Exception ex)
                {
                    //CSafeSetmemoEdit(memoEdit1, "알람쓰레드오류 : " + ex.Message);
                }
            }
        }

        /// <summary>
        /// 알림메세지 테이블 가져오기
        /// </summary>
        /// <param name="sBungi">알람컬럼명</param>
        /// <param name="dAnum2">알람합수</param>
        /// <returns></returns>
        private DataTable ReturnAlramMsgDt(string sBungi, double dAnum2)
        {
            Dictionary<string, object> dicParams = new Dictionary<string, object>();

            dicParams.Clear();
            dicParams.Add("BUNGI", sBungi);
            dicParams.Add("NUM", dAnum2);

            DataTable dt = DBConn.GetDataTable(this.PROCEDURE_AL, dicParams);

            return dt;
        }

        /// <summary>
        /// 알림메세지 세팅
        /// </summary>
        /// <param name="dt">알람메세지 테이블</param>
        /// <param name="sMsg"></param>
        /// <returns></returns>
        private string ReturnAlramMsg(DataTable dt, string sMsg)
        {
            if (dt != null && dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DateTime now = DateTime.Now;

                    string sAmsg = "["+ now.ToString("HH:mm:ss") + "]"+dt.Rows[i]["AMSG"]?.ToString();

                    if (string.IsNullOrEmpty(sAmsg))
                        continue;

                    sMsg += "\r\n" + sAmsg;
                }
            }

            return sMsg;
        }
        #endregion

        #region [ 크로스 스레드 ]
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

        //1차공정 진행도
        delegate void CrossThreadSafetySetProgress(ProgressBarControl ctl, int iVal);
        private void CSafeSetProgressControl(ProgressBarControl ctl, int iVal)
        {
            if (ctl.InvokeRequired)
            {
                ctl.Invoke(new CrossThreadSafetySetProgress(CSafeSetProgressControl), ctl, iVal);
            }
            else
            {
                ctl.LookAndFeel.UseDefaultLookAndFeel = false;  // 반드시 오프시켜야 함
                ctl.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat;
                //ctl.Properties.Appearance.BackColor = Color.White;

                ctl.Properties.StartColor = Color.FromArgb(50, 172, 192);
                ctl.Properties.EndColor = Color.FromArgb(138, 195, 152);

                ctl.EditValue = iVal;//진행도
            }
        }

        //오일량
        delegate void CrossThreadSafetySetImageBoxOil(ProgressBarControl ctl, int prgval);
        private void CSafeSetImageBoxOil(ProgressBarControl ctl, int prgval)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetImageBoxOil(CSafeSetImageBoxOil), ctl, prgval);
            else
            {
                ctl.LookAndFeel.UseDefaultLookAndFeel = false;  // 반드시 오프시켜야 함
                ctl.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat;
                //ctl.Properties.Appearance.BackColor = Color.White;

                ctl.Properties.StartColor = Color.FromArgb(50, 172, 192);
                ctl.Properties.EndColor = Color.FromArgb(138, 195, 152);

                ctl.EditValue = prgval;//오일량 지정
            }
        }

        //label value 
        delegate void CrossThreadSafetySetTextLabel(Label ctl, string str);
        private void CSafeSetLabel(Label ctl, string str)
        {
            if (ctl.InvokeRequired)
            {
                ctl.Invoke(new CrossThreadSafetySetTextLabel(CSafeSetLabel), ctl, str);
            }
            else
            {
                ctl.Text = str;
            }
        }

        //전류,압력 라벨 이미지변경
        delegate void CrossThreadSafetySetTextLabelImage(Label ctl, Image img);
        private void CSafeSetLabelImage(Label ctl, Image img)
        {
            if (ctl.InvokeRequired)
            {
                ctl.Invoke(new CrossThreadSafetySetTextLabelImage(CSafeSetLabelImage), ctl, img);
            }
            else
            {
                ctl.Image = img;
            }
        }
        #endregion

        #region [ 버튼이벤트 ]
        //작업자 DB저장
        private void SetWorkerBtn(string sEmpid)
        {
            string sworkname = string.Empty;

            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Clear();
            dicParams.Add("CMD", "WORKER_ID");
            dicParams.Add("EMPID", sEmpid);
            DataTable dt = DBConn.GetDataTable("DP_MONITERING_SV", dicParams);
            if (dt.Rows.Count > 0)
            {
                sworkname = dt.Rows[0]["EMP_NM"]?.ToString();

                if (sworkname.Equals(Btnworker1.Text))
                {
                    //버튼변경
                    Btnworker1.Appearance.BackColor = Color.Gray;
                    Btnworker1.Appearance.BackColor2 = Color.Gray;
                    Btnworker1.Appearance.BorderColor = Color.Gray;
                    Btnworker1.Appearance.Options.UseBackColor = true;
                    Btnworker1.Appearance.Options.UseBorderColor = true;

                    foreach (SimpleButton btn in new SimpleButton[] { Btnworker2, Btnworker3, Btnworker4, Btnworker5, Btnworker6 })
                    {
                        btn.ForeColor = Color.Black;
                        btn.Appearance.BackColor = btn.Appearance.BackColor2 = btn.Appearance.BorderColor = Color.White;
                        btn.Appearance.Options.UseBackColor = btn.Appearance.Options.UseBorderColor = true;
                    }

                }

                else if (sworkname.Equals(Btnworker2.Text))
                {
                    //버튼변경
                    Btnworker2.Appearance.BackColor = Color.Gray;
                    Btnworker2.Appearance.BackColor2 = Color.Gray;
                    Btnworker2.Appearance.BorderColor = Color.Gray;
                    Btnworker2.Appearance.Options.UseBackColor = true;
                    Btnworker2.Appearance.Options.UseBorderColor = true;

                    foreach (SimpleButton btn in new SimpleButton[] { Btnworker1, Btnworker3, Btnworker4, Btnworker5, Btnworker6 })
                    {
                        btn.ForeColor = Color.Black;
                        btn.Appearance.BackColor = btn.Appearance.BackColor2 = btn.Appearance.BorderColor = Color.White;
                        btn.Appearance.Options.UseBackColor = btn.Appearance.Options.UseBorderColor = true;
                    }
                }
                else if (sworkname.Equals(Btnworker3.Text))
                {
                    //버튼변경
                    Btnworker3.Appearance.BackColor = Color.Gray;
                    Btnworker3.Appearance.BackColor2 = Color.Gray;
                    Btnworker3.Appearance.BorderColor = Color.Gray;
                    Btnworker3.Appearance.Options.UseBackColor = true;
                    Btnworker3.Appearance.Options.UseBorderColor = true;

                    foreach (SimpleButton btn in new SimpleButton[] { Btnworker1, Btnworker2, Btnworker4, Btnworker5, Btnworker6 })
                    {
                        btn.ForeColor = Color.Black;
                        btn.Appearance.BackColor = btn.Appearance.BackColor2 = btn.Appearance.BorderColor = Color.White;
                        btn.Appearance.Options.UseBackColor = btn.Appearance.Options.UseBorderColor = true;
                    }
                }
                else if (sworkname.Equals(Btnworker4.Text))
                {
                    //버튼변경
                    Btnworker4.Appearance.BackColor = Color.Gray;
                    Btnworker4.Appearance.BackColor2 = Color.Gray;
                    Btnworker4.Appearance.BorderColor = Color.Gray;
                    Btnworker4.Appearance.Options.UseBackColor = true;
                    Btnworker4.Appearance.Options.UseBorderColor = true;

                    foreach (SimpleButton btn in new SimpleButton[] { Btnworker1, Btnworker2, Btnworker3, Btnworker5, Btnworker6 })
                    {
                        btn.ForeColor = Color.Black;
                        btn.Appearance.BackColor = btn.Appearance.BackColor2 = btn.Appearance.BorderColor = Color.White;
                        btn.Appearance.Options.UseBackColor = btn.Appearance.Options.UseBorderColor = true;
                    }
                }
                else if (sworkname.Equals(Btnworker5.Text))
                {
                    //버튼변경
                    Btnworker5.Appearance.BackColor = Color.Gray;
                    Btnworker5.Appearance.BackColor2 = Color.Gray;
                    Btnworker5.Appearance.BorderColor = Color.Gray;
                    Btnworker5.Appearance.Options.UseBackColor = true;
                    Btnworker5.Appearance.Options.UseBorderColor = true;

                    foreach (SimpleButton btn in new SimpleButton[] { Btnworker1, Btnworker2, Btnworker3, Btnworker4, Btnworker6 })
                    {
                        btn.ForeColor = Color.Black;
                        btn.Appearance.BackColor = btn.Appearance.BackColor2 = btn.Appearance.BorderColor = Color.White;
                        btn.Appearance.Options.UseBackColor = btn.Appearance.Options.UseBorderColor = true;
                    }
                }
                else if (sworkname.Equals(Btnworker6.Text))
                {
                    //버튼변경
                    Btnworker6.Appearance.BackColor = Color.Gray;
                    Btnworker6.Appearance.BackColor2 = Color.Gray;
                    Btnworker6.Appearance.BorderColor = Color.Gray;
                    Btnworker6.Appearance.Options.UseBackColor = true;
                    Btnworker6.Appearance.Options.UseBorderColor = true;

                    foreach (SimpleButton btn in new SimpleButton[] { Btnworker1, Btnworker2, Btnworker3, Btnworker4, Btnworker5 })
                    {
                        btn.ForeColor = Color.Black;
                        btn.Appearance.BackColor = btn.Appearance.BackColor2 = btn.Appearance.BorderColor = Color.White;
                        btn.Appearance.Options.UseBackColor = btn.Appearance.Options.UseBorderColor = true;
                    }
                }
            }
            
        }
        private void SetWorkerDB(string sEmpid)
        {
            try
            {
                if (_EMPID.Equals(sEmpid))
                    return;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "SAVE1");
                dicParams.Add("EMPID", sEmpid);
                dicParams.Add("MGUBUN", _GUBUN);

                DataTable dt = DBConn.GetDataTable("DP_MONITERING_SV", dicParams);

                _EMPID = sEmpid;
            }
            catch (Exception ex)
            {

            }
        }
        
        //품목 버튼 선택
        private void PicGub1_DoubleClick(object sender, EventArgs e)
        {
            PictureEdit pictureEdit = (PictureEdit)sender;

            if (pictureEdit == null)
                return;

            string sGub = pictureEdit.Tag.ToString();
            //버튼변경
            SetGubBtn(sGub);
            //DB변경
            SetGubDB(sGub);
        }

        private void SetGubDB(string sGub)
        {
            try
            {
                if (_GUBUN.Equals(sGub))
                {
                    sGub = "G S";//#0001
                }

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "SAVE1");
                dicParams.Add("EMPID", _EMPID);
                dicParams.Add("MGUBUN", sGub);

                DataTable dt = DBConn.GetDataTable("DP_MONITERING_SV", dicParams);

                _GUBUN = sGub;
            }
            catch (Exception ex)
            {

            }
        }

        //품목별 버튼변경
        private void SetGubBtn(string sGub)
        {
            if (sGub.Equals("G S"))
            {
                PicGub1.Image = Properties.Resources.GSON;
                PicGub2.Image = Properties.Resources.압축OFF;
                PicGub3.Image = Properties.Resources.중량OFF;
            }
            else if (sGub.Equals("압축"))
            {
                PicGub1.Image = Properties.Resources.GSOFF;
                PicGub2.Image = Properties.Resources.압축ON;
                PicGub3.Image = Properties.Resources.중량OFF;
            }
            else if (sGub.Equals("중량"))
            {
                PicGub1.Image = Properties.Resources.GSOFF;
                PicGub2.Image = Properties.Resources.압축OFF;
                PicGub3.Image = Properties.Resources.중량ON;
            }
            else
            {
                PicGub1.Image = Properties.Resources.GSON;
                PicGub2.Image = Properties.Resources.압축OFF;
                PicGub3.Image = Properties.Resources.중량OFF;
            }
        }

        //cctv 화면 클릭
        private void PicCCTV_Click(object sender, EventArgs e)
        {
            //도면화면일경우 리턴
            if (_CCTV_GB == CCTV_GB.PRINT)
                return;


            PictureBox pictureBox = (PictureBox)sender;

            if (pictureBox == null)
                return;

            string sRSTPaddr = string.Empty;
            if (pictureBox.Name.Equals("PicCCTV1"))//1카메라 선택시
            {
                sRSTPaddr = _CAM1;
            }
            else if (pictureBox.Name.Equals("PicCCTV2"))//2카메라선택시
            {
                sRSTPaddr = _CAM2;
            }


            CV001F00 frm = new CV001F00();

            frm.Owner = this;
            frm._RSTPaddr = sRSTPaddr;

            if (frm.ShowDialog() == DialogResult.OK)
            {

            }
        }

        //화면전환버튼선택
        private void PicNextDisplay_Click(object sender, EventArgs e)
        {
            if (LayContPicCCTV2.Visibility != DevExpress.XtraLayout.Utils.LayoutVisibility.Always)
            {
                //설계도면화면일때 CCTV화면으로 변경
                _CCTV_GB = CCTV_GB.CCTV;
                LayContPicCCTV2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                PicCCTV1.Image = Properties.Resources.MONITOR1BACK;
            }
            else
            {
                //CCTV 화면일때 설계도면 화면으로 변경
                _CCTV_GB = CCTV_GB.PRINT;
                LayContPicCCTV2.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                //PicCCTV1.Image = Properties.Resources.설계도면;
            }
        }

        //환경설정
        private void BtnSet_Click(object sender, EventArgs e)
        {
            Form fchk = Application.OpenForms["SETTING"];

            if (fchk == null)
            {
                SETTING frm = new SETTING();

                frm.Owner = this;
                MethodInvoker mi = new MethodInvoker(frm.Show);
                this.Invoke(mi);
            }
        }

        //닫기
        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //경고창 보임,안보임 설정
        private void pictureEdit1_Click(object sender, EventArgs e)
        {
            _ALMDISPALY = ALMDISPALY.ON;
        }

        /// <summary>
        /// 경고창 열기
        /// </summary>
        /// <param name="sMSG">경고메세지</param>
        private void SetWorningMSG(string sMSG)
        {
            if(_ALMDISPALY == ALMDISPALY.ON || string.IsNullOrEmpty(sMSG))
            {
                //알람창 보일때 상단 경고 이미지 숨기기
                CSafeSetImage(pictureEdit1, null);
            }
            else
            {
                //알람창 안보일때 상단 경고 이미지 보이기
                //CSafeSetImage(pictureEdit1, Properties.Resources.warning);
                return;
            }

            Form fchk = Application.OpenForms["MS001F00"];

            if (fchk == null)
            {
                if (string.IsNullOrEmpty(sMSG))
                    return;

                MS001F00 frm = new MS001F00();

                frm.Owner = this;
                frm._MSG = sMSG.Trim();
                MethodInvoker mi = new MethodInvoker(frm.Show);
                this.Invoke(mi);

                //배경 form 열기
                BACKGRUOND backFrm = new BACKGRUOND();

                backFrm.Owner = this;
                MethodInvoker miBack = new MethodInvoker(backFrm.Show);
                this.Invoke(miBack);
            }
            else
            {
                if (string.IsNullOrEmpty(sMSG))
                {
                    MethodInvoker mi = new MethodInvoker(fchk.Close);
                    this.Invoke(mi);

                    Form fchk2 = Application.OpenForms["BACKGRUOND"];
                    MethodInvoker mi2 = new MethodInvoker(fchk2.Close);
                    this.Invoke(mi2);

                    return;
                }

                MS001F00 frm = (MS001F00)fchk;

                frm.Owner = this;
                frm._MSG = sMSG.Trim();
                frm.LoadMsg();
            }
        }

        #endregion

        //작업자 선택(더블클릭)
        private void Btnworker_DoubleClick(object sender, EventArgs e)
        {
            SimpleButton sb = (SimpleButton)sender;
            string stag = sb.Tag.ToString();

            if (stag.Equals("worker1"))
            {
                string sEmpid1 = string.Empty;
                string sworker1 = Btnworker1.Text;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();
                dicParams.Clear();
                dicParams.Add("CMD", "WORKER");
                dicParams.Add("EMPNM", sworker1);
                DataTable dt = DBConn.GetDataTable("DP_MONITERING_SV", dicParams);

                sEmpid1 = dt.Rows[0]["EMP_ID"]?.ToString();

                //버튼변경
                Btnworker1.Appearance.BackColor = Color.Gray;
                Btnworker1.Appearance.BackColor2 = Color.Gray;
                Btnworker1.Appearance.BorderColor = Color.Gray;
                Btnworker1.Appearance.Options.UseBackColor = true;
                Btnworker1.Appearance.Options.UseBorderColor = true;

                foreach (SimpleButton btn in new SimpleButton[] { Btnworker2, Btnworker3, Btnworker4, Btnworker5, Btnworker6 })
                {
                    btn.ForeColor = Color.Black;
                    btn.Appearance.BackColor = btn.Appearance.BackColor2 = btn.Appearance.BorderColor = Color.White;
                    btn.Appearance.Options.UseBackColor = btn.Appearance.Options.UseBorderColor = true;
                }
                //DB변경
                SetWorkerDB(sEmpid1);

            }

            else if (stag.Equals("worker2"))
            {
                string sEmpid2 = string.Empty;
                string sworker2 = Btnworker2.Text;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();
                dicParams.Clear();
                dicParams.Add("CMD", "WORKER");
                dicParams.Add("EMPNM", sworker2);
                DataTable dt = DBConn.GetDataTable("DP_MONITERING_SV", dicParams);

                sEmpid2 = dt.Rows[0]["EMP_ID"]?.ToString();

                //버튼변경
                Btnworker2.Appearance.BackColor = Color.Gray;
                Btnworker2.Appearance.BackColor2 = Color.Gray;
                Btnworker2.Appearance.BorderColor = Color.Gray;
                Btnworker2.Appearance.Options.UseBackColor = true;
                Btnworker2.Appearance.Options.UseBorderColor = true;

                foreach (SimpleButton btn in new SimpleButton[] { Btnworker1, Btnworker3, Btnworker4, Btnworker5, Btnworker6 })
                {
                    btn.ForeColor = Color.Black;
                    btn.Appearance.BackColor = btn.Appearance.BackColor2 = btn.Appearance.BorderColor = Color.White;
                    btn.Appearance.Options.UseBackColor = btn.Appearance.Options.UseBorderColor = true;
                }
                //DB변경
                SetWorkerDB(sEmpid2);
            }
            else if (stag.Equals("worker3"))
            {
                string sEmpid3 = string.Empty;
                string sworker3 = Btnworker3.Text;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();
                dicParams.Clear();
                dicParams.Add("CMD", "WORKER");
                dicParams.Add("EMPNM", sworker3);
                DataTable dt = DBConn.GetDataTable("DP_MONITERING_SV", dicParams);

                sEmpid3 = dt.Rows[0]["EMP_ID"]?.ToString();
                //버튼변경
                Btnworker3.Appearance.BackColor = Color.Gray;
                Btnworker3.Appearance.BackColor2 = Color.Gray;
                Btnworker3.Appearance.BorderColor = Color.Gray;
                Btnworker3.Appearance.Options.UseBackColor = true;
                Btnworker3.Appearance.Options.UseBorderColor = true;

                foreach (SimpleButton btn in new SimpleButton[] { Btnworker1, Btnworker2, Btnworker4, Btnworker5, Btnworker6 })
                {
                    btn.ForeColor = Color.Black;
                    btn.Appearance.BackColor = btn.Appearance.BackColor2 = btn.Appearance.BorderColor = Color.White;
                    btn.Appearance.Options.UseBackColor = btn.Appearance.Options.UseBorderColor = true;
                }
                //DB변경
                SetWorkerDB(sEmpid3);
            }
            else if (stag.Equals("worker4"))
            {
                string sEmpid4 = string.Empty;
                string sworker4 = Btnworker4.Text;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();
                dicParams.Clear();
                dicParams.Add("CMD", "WORKER");
                dicParams.Add("EMPNM", sworker4);
                DataTable dt = DBConn.GetDataTable("DP_MONITERING_SV", dicParams);

                sEmpid4 = dt.Rows[0]["EMP_ID"]?.ToString();
                //버튼변경
                Btnworker4.Appearance.BackColor = Color.Gray;
                Btnworker4.Appearance.BackColor2 = Color.Gray;
                Btnworker4.Appearance.BorderColor = Color.Gray;
                Btnworker4.Appearance.Options.UseBackColor = true;
                Btnworker4.Appearance.Options.UseBorderColor = true;

                foreach (SimpleButton btn in new SimpleButton[] { Btnworker1, Btnworker2, Btnworker3, Btnworker5, Btnworker6 })
                {
                    btn.ForeColor = Color.Black;
                    btn.Appearance.BackColor = btn.Appearance.BackColor2 = btn.Appearance.BorderColor = Color.White;
                    btn.Appearance.Options.UseBackColor = btn.Appearance.Options.UseBorderColor = true;
                }
                //DB변경
                SetWorkerDB(sEmpid4);
            }
            else if (stag.Equals("worker5"))
            {
                string sEmpid5 = string.Empty;
                string sworker5 = Btnworker5.Text;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();
                dicParams.Clear();
                dicParams.Add("CMD", "WORKER");
                dicParams.Add("EMPNM", sworker5);
                DataTable dt = DBConn.GetDataTable("DP_MONITERING_SV", dicParams);

                sEmpid5 = dt.Rows[0]["EMP_ID"]?.ToString();
                //버튼변경
                Btnworker5.Appearance.BackColor = Color.Gray;
                Btnworker5.Appearance.BackColor2 = Color.Gray;
                Btnworker5.Appearance.BorderColor = Color.Gray;
                Btnworker5.Appearance.Options.UseBackColor = true;
                Btnworker5.Appearance.Options.UseBorderColor = true;

                foreach (SimpleButton btn in new SimpleButton[] { Btnworker1, Btnworker2, Btnworker3, Btnworker4, Btnworker6 })
                {
                    btn.ForeColor = Color.Black;
                    btn.Appearance.BackColor = btn.Appearance.BackColor2 = btn.Appearance.BorderColor = Color.White;
                    btn.Appearance.Options.UseBackColor = btn.Appearance.Options.UseBorderColor = true;
                }
                //DB변경
                SetWorkerDB(sEmpid5);
            }
            else if (stag.Equals("worker6"))
            {
                string sEmpid6 = string.Empty;
                string sworker6 = Btnworker6.Text;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();
                dicParams.Clear();
                dicParams.Add("CMD", "WORKER");
                dicParams.Add("EMPNM", sworker6);
                DataTable dt = DBConn.GetDataTable("DP_MONITERING_SV", dicParams);

                sEmpid6 = dt.Rows[0]["EMP_ID"]?.ToString();
                //버튼변경
                Btnworker6.Appearance.BackColor = Color.Gray;
                Btnworker6.Appearance.BackColor2 = Color.Gray;
                Btnworker6.Appearance.BorderColor = Color.Gray;
                Btnworker6.Appearance.Options.UseBackColor = true;
                Btnworker6.Appearance.Options.UseBorderColor = true;

                foreach (SimpleButton btn in new SimpleButton[] { Btnworker1, Btnworker2, Btnworker3, Btnworker4, Btnworker5 })
                {
                    btn.ForeColor = Color.Black;
                    btn.Appearance.BackColor = btn.Appearance.BackColor2 = btn.Appearance.BorderColor = Color.White;
                    btn.Appearance.Options.UseBackColor = btn.Appearance.Options.UseBorderColor = true;
                }
                //DB변경
                SetWorkerDB(sEmpid6);
            }
        }
        
    }
}