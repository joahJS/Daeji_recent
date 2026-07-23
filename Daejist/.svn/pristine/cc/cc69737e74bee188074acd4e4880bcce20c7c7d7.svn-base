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
using MySql.Data.MySqlClient;
using System.Net;
using System.Net.Sockets;
using System.IO;
using System.IO.Ports;
using ComLib;
using System.Threading;
using Timer = System.Windows.Forms.Timer;
using System.Diagnostics;
using System.Timers;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid;
using System.Data.SqlClient;

/*
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 수정일자 : 2021-01-20
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)미계근 데이터는 검색일자를 From에서 한달 이전으로 세팅
 * 메소드 GetNotComplteMes() 참고
 * 
 * 수정일자 : 2021-01-24
 * 수정자   : 고혜성
 * 수정내용 : (현업요청) 
 *            1. ORDER BY J_DATE, SUN으로 처리
 *            2. 총중량, 공차중량의 버튼을 레이블로 변경처리
 *            
 * 수정일자 : 2021-02-22
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. MESURING 업체담당자 컬럼추가됨에 따라 입출고 1/2차 INSERT/UPDATE 쿼리에 CHRG_ID컬럼 반영
 *             저장관련 메소드 참고
 *             
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 *            
 * 수정일자 : 2021-03-02
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 1차계근 수기입력 시 무인계근시스템과 Data Sync를 맞추기 위하여
 *            1차 계근 시간을 1/2차 계근시간에 모두 적용하도록 쿼리수정 (FirstTime, SecondTime)
 *            
 * 수정일자 : 2021-03-11
 * 수정자   : 고혜성
 * Reference Key = #0001
 * 수정내용 : (현업요청)
 *            1. 계근변경 시 단가가 날라가는 현상 수정
 *
 * 수정일자 : 2021-03-16
 * 수정자   : 고혜성
 * Reference Key = #0002
 * 수정내용 : (현업요청)
 *            1. 로그추가
 *            
 * 수정일자 : 2021-03-17
 * 수정자   : 고혜성
 * Reference Key = #0004
 * 수정내용 : (현업요청)
 *            1. 로그수정
 *               1-1) 로그 맨 뒷단 (Table ~)은 안나오도록 수정
 *               1-2) 계근일자가 이전/이후가 제대로 매칭이 안되는 현상 수정완료
 *               1-3) 등급 변경 시 등급코드가 아닌 명으로 나오도록 수정
 *               1-4) 기본사항/변경사항을 나누어 입력
 *               
 * 수정일자 : 2021-03-18
 * 수정자   : 고혜성
 * 수정내용 : (현업요청) #0004 관련
 *            1. 기본항목에 입력되는 값 단순화
 *            
 * 수정일자 : 2021-03-22
 * 수정자   : 고혜성
 * Reference Key : #0005
 * 수정내용 : (현업요청)
 *            1. 검수내역조회와 관련하여 최초 감량을 저장하여 불러오기 위하여
 *              감량이 변경되었을 경우 별도 테이블에 감량값 저장
 *              
 * 수정일자 : 2021-03-26
 * 수정자   : 고혜성
 * Reference Key : #0006
 * 수정내용 : (현업요청)
 *            1. 삭제 시 비고사항에 기존정보 추가하여 적용
 */
namespace AccAdm
{
    public partial class MesMgtProgramDev : DevExpress.XtraEditors.XtraForm
    {
        TcpListener Server; // 소켓 서버
        TcpClient Client; // 클라이언트
        StreamReader Reader;
        StreamWriter Writer;
        NetworkStream stream; // 네트워크 스트림 연결

        string reciveData = string.Empty;
        private delegate void AddTextDelegate(string strText); // 크로스 쓰레드 호출
        delegate void SetTextCallback(string Text);
        Thread ListenThread;
        System.Timers.Timer timer;

        //public enum DataChanged { Changed, UnChanged }
        //public Dictionary<string, DataChanged> _dicDataChanged = new Dictionary<string, DataChanged>();

        public MesMgtProgramDev()
        {
            InitializeComponent();
            
            // Initialize Serial Port
            serialPort1.BaudRate = 9600;
            serialPort1.DataBits = 7;
            serialPort1.Parity = Parity.None;
            serialPort1.StopBits = StopBits.One;

            ////로그이력 남기기 위하여 초기세팅
            //_dicDataChanged.Add("J_DATE", DataChanged.UnChanged); //계량일
            //_dicDataChanged.Add("KERATYPE", DataChanged.UnChanged); //입출고선택
            //_dicDataChanged.Add("SUN", DataChanged.UnChanged); //순번
            //_dicDataChanged.Add("J_BNUM", DataChanged.UnChanged); //차량번호
            //_dicDataChanged.Add("COMPANY", DataChanged.UnChanged); //거래처
            //_dicDataChanged.Add("J_SERIAL", DataChanged.UnChanged); //품명
            //_dicDataChanged.Add("GUMSU_SERIAL", DataChanged.UnChanged); //검수자
            //_dicDataChanged.Add("GUMSUBIGO", DataChanged.UnChanged); //검수비고
            //_dicDataChanged.Add("TOT_WEIGHT", DataChanged.UnChanged); //총중량
            //_dicDataChanged.Add("EMPTY_WEIGHT", DataChanged.UnChanged); //공차중량
            //_dicDataChanged.Add("CHAGAM", DataChanged.UnChanged); //감량
            //_dicDataChanged.Add("J_STATE", DataChanged.UnChanged); //감량사유
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void MesMgtProgramDev_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            DateEditFrom.EditValue = DateTime.Now;
            DateEditTo.EditValue = DateTime.Now;
            DateEditMes.EditValue = DateTime.Now;

            DataTable dtDealer = GetLookUpData("1", "Y", "");
            DataTable dtGrade = GetLookUpData("2", "Y", "");
            DataTable dtSawon = GetLookUpData("3", "Y", "");

            ComLib.ComGrid.SetLookUpEdit(LkupEditDealerNM, dtDealer, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupEditGrade, dtGrade, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupEditInspector, dtSawon, "CD", "NM", "Y");

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewRetr };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            BtnRetr_Click(null, null);

            Cursor = Cursors.Default;
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (ChkCompleteMes.Checked)
            {
                GetCompleteMes();
            }
            else
            {
                GetNotComplteMes();
            }

            if(GridViewRetr.RowCount == 0)
            {
                pictureEdit1.Image = null;
                pictureEdit2.Image = null;
                pictureEdit3.Image = null;
                pictureEdit4.Image = null;
                pictureEdit5.Image = null;
                pictureEdit6.Image = null;
            }
        }

        private void ChkCompleteMes_CheckStateChanged(object sender, EventArgs e)
        {
            if (ChkCompleteMes.Checked)
            {
                GetCompleteMes();
            }
            else
            {
                GetNotComplteMes();
            }
        }

        private void GetCompleteMes()
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            Cursor = Cursors.WaitCursor;

            string sYmdFrom = DateEditFrom.EditValue.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue.ToString().Substring(0, 10);

            ClearFps();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.JUNPYOID ");
            strSql.AppendLine(" 	 , A.SUN ");
            strSql.AppendLine("      , A.KERATYPE");
            strSql.AppendLine(" 	 , A.J_DATE ");
            strSql.AppendLine(" 	 , A.J_BNUM ");
            strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHERID ELSE A.J_ASSIGNID END AS DEALER_CD ");
            strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER ");
            strSql.AppendLine("      , A.J_SERIAL ");
            strSql.AppendLine(" 	 , A.GUBUN1 ");
            strSql.AppendLine(" 	 , A.FIRSTWEIGHT ");
            strSql.AppendLine(" 	 , A.FIRSTTIME ");
            strSql.AppendLine(" 	 , A.SECONDWEIGHT ");
            strSql.AppendLine(" 	 , A.SECONDTIME ");
            strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.FIRSTTIME ELSE A.SECONDTIME END AS SECONDTIME ");
            strSql.AppendLine("      , A.WEIGHT AS ACTUALWEIGHT");
            strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.ICHAGAM ELSE A.OCHAGAM END AS LOSS ");
            //strSql.AppendLine(" 	 , (CASE WHEN A.KERATYPE = '입고' THEN A.IWEIGHT ELSE A.OWEIGHT END) - (CASE WHEN A.KERATYPE = '입고' THEN A.ICHAGAM ELSE A.OCHAGAM END) AS ACCEPTWEIGHT ");
            strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.IWEIGHT ELSE A.OWEIGHT END AS ACCEPTWEIGHT ");
            strSql.AppendLine("      , A.J_STATE AS LOSSCAUSE ");
            strSql.AppendLine("      , A.GUMSUBIGO AS INSPECTNOTE");
            strSql.AppendLine(" 	 , B.EMP_NM AS INSPECTOR ");
            strSql.AppendLine("      , A.GUMSU_SERIAL ");
            strSql.AppendLine("   FROM MESURING A ");
            strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
            strSql.AppendLine("     ON B.EMP_ID = A.GUMSU_SERIAL ");
            strSql.AppendLine("  WHERE A.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            strSql.AppendLine("    AND A.KERATYPE != '직송' ");
            strSql.AppendLine("    AND A.FIRSTWEIGHT != '0' ");
            strSql.AppendLine("    AND  A.SECONDWEIGHT != '0' ");
            strSql.AppendLine("  ORDER BY J_DATE, SUN");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void GetNotComplteMes()
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            Cursor = Cursors.WaitCursor;

            string sYmdFrom = DateEditFrom.EditValue.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue.ToString().Substring(0, 10);

            ClearFps();

            StringBuilder strSql = new StringBuilder();
            /*
             * 2021-01-20 현업요청
             * 미계근완료데이터는 조회일자를 한달이전으로 세팅
             */
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.JUNPYOID ");
            strSql.AppendLine(" 	 , A.SUN ");
            strSql.AppendLine("      , A.KERATYPE");
            strSql.AppendLine(" 	 , CONVERT(DATE,A.J_DATE) AS J_DATE ");
            strSql.AppendLine(" 	 , A.J_BNUM ");
            strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHERID ELSE A.J_ASSIGNID END AS DEALER_CD ");
            strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER ");
            strSql.AppendLine("      , A.J_SERIAL ");
            strSql.AppendLine(" 	 , A.GUBUN1 ");
            strSql.AppendLine(" 	 , A.FIRSTWEIGHT ");
            strSql.AppendLine(" 	 , A.FIRSTTIME ");
            strSql.AppendLine(" 	 , A.SECONDWEIGHT ");
            strSql.AppendLine(" 	 , A.SECONDTIME ");
            strSql.AppendLine("      , A.WEIGHT AS ACTUALWEIGHT");
            strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.ICHAGAM ELSE A.OCHAGAM END AS LOSS ");
            strSql.AppendLine(" 	 , A.WEIGHT AS ACCEPTWEIGHT ");
            //strSql.AppendLine(" 	 , A.WEIGHT - (CASE WHEN A.KERATYPE = '입고' THEN A.ICHAGAM ELSE A.OCHAGAM END) AS ACCEPTWEIGHT ");
            strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.IWEIGHT ELSE A.OWEIGHT END AS ACCEPTWEIGHT ");
            strSql.AppendLine("      , A.J_STATE AS LOSSCAUSE ");
            strSql.AppendLine("      , A.GUMSUBIGO AS INSPECTNOTE");
            strSql.AppendLine(" 	 , B.EMP_NM AS INSPECTOR ");
            strSql.AppendLine("      , A.GUMSU_SERIAL ");
            strSql.AppendLine("   FROM MESURING A ");
            strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
            strSql.AppendLine("     ON B.EMP_ID = A.GUMSU_SERIAL ");
            strSql.AppendLine("  WHERE A.KERATYPE != '직송' ");
            strSql.AppendLine("    AND A.J_DATE BETWEEN DATEADD(day, -30,'" + sYmdFrom + "') AND DATEADD(day, 30 ,'" + sYmdTo + "') ");
            strSql.AppendLine("    AND ( A.FIRSTWEIGHT = 0 ");
            strSql.AppendLine("     OR  A.SECONDWEIGHT = 0) ");
            strSql.AppendLine("  ORDER BY J_DATE, SUN");

            #region[2021-01-20 수정 전 쿼리]

            //strSql.AppendLine(" ");
            //strSql.AppendLine(" SELECT A.JUNPYOID ");
            //strSql.AppendLine(" 	 , A.SUN ");
            //strSql.AppendLine("      , A.KERATYPE");
            //strSql.AppendLine(" 	 , A.J_DATE ");
            //strSql.AppendLine(" 	 , A.J_BNUM ");
            //strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHERID ELSE A.J_ASSIGNID END AS DEALER_CD ");
            //strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER ");
            //strSql.AppendLine("      , A.J_SERIAL ");
            //strSql.AppendLine(" 	 , A.GUBUN1 ");
            //strSql.AppendLine(" 	 , A.FIRSTWEIGHT ");
            //strSql.AppendLine(" 	 , A.FIRSTTIME ");
            //strSql.AppendLine(" 	 , A.SECONDWEIGHT ");
            //strSql.AppendLine(" 	 , A.SECONDTIME ");
            //strSql.AppendLine("      , A.WEIGHT AS ACTUALWEIGHT");
            //strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.ICHAGAM ELSE A.OCHAGAM END AS LOSS ");
            //strSql.AppendLine(" 	 , A.WEIGHT AS ACCEPTWEIGHT ");
            ////strSql.AppendLine(" 	 , A.WEIGHT - (CASE WHEN A.KERATYPE = '입고' THEN A.ICHAGAM ELSE A.OCHAGAM END) AS ACCEPTWEIGHT ");
            //strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.IWEIGHT ELSE A.OWEIGHT END AS ACCEPTWEIGHT ");
            //strSql.AppendLine("      , A.J_STATE AS LOSSCAUSE ");
            //strSql.AppendLine("      , A.GUMSUBIGO AS INSPECTNOTE");
            //strSql.AppendLine(" 	 , B.EMP_NM AS INSPECTOR ");
            //strSql.AppendLine("      , A.GUMSU_SERIAL ");
            //strSql.AppendLine("   FROM MESURING A ");
            //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
            //strSql.AppendLine("     ON B.EMP_ID = A.GUMSU_SERIAL ");
            //strSql.AppendLine("  WHERE A.KERATYPE != '직송' ");
            //strSql.AppendLine("    AND A.J_DATE >= '" + sYmdFrom + "' ");
            //strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            //strSql.AppendLine("    AND ( A.FIRSTWEIGHT = 0 ");
            //strSql.AppendLine("     OR  A.SECONDWEIGHT = 0) ");
            //strSql.AppendLine("  ORDER BY J_DATE");

            #endregion[2021-01-20 수정 전 쿼리]

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void BtnNew_Click(object sender, EventArgs e)
        {
            if (BtnNew.Text.Equals("신규(F1)"))
            {
                if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
                {
                    XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                    return;
                }

                DateEditMes.EditValue = DateTime.Today;
                RdGbSelect.SelectedIndex = 0;

                string sYmd = DateEditMes.EditValue?.ToString().Substring(0, 10);

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT CASE WHEN MAX(SUN) IS NULL THEN 1 ELSE MAX(SUN) + 1 END AS MAX_VALUE ");
                strSql.AppendLine("   FROM MESURING ");
                strSql.AppendLine("  WHERE J_DATE = '" + sYmd + "' ");
                strSql.AppendLine("    AND KERATYPE <> '직송' ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                string sMaxValue = dt.Rows[0]["MAX_VALUE"].ToString();

                TxtSeq.EditValue = sMaxValue;
                TxtCarNo.EditValue = string.Empty;
                LkupEditDealerNM.EditValue = string.Empty;
                LkupEditGrade.EditValue = string.Empty;
                LkupEditInspector.EditValue = string.Empty;
                TxtInsptNote.EditValue = string.Empty;
                TxtTotWeight.EditValue = string.Empty;
                TxtEmptyWeight.EditValue = string.Empty;
                TimeEditFirst.EditValue = DateTime.Now;
                TimeEditSecond.EditValue = DateTime.Now;
                TxtActualWeight.EditValue = string.Empty;
                TxtLossWeight.EditValue = string.Empty;
                TxtLossCause.EditValue = string.Empty;
                TxtAcceptWeight.EditValue = string.Empty;

                TxtJunpyoId.EditValue = null;

                MakeUnReadOnly();

                pictureEdit1.EditValue = null;
                pictureEdit2.EditValue = null;
                pictureEdit3.EditValue = null;
                pictureEdit4.EditValue = null;
                pictureEdit5.EditValue = null;
                pictureEdit6.EditValue = null;

                BtnNew.Text = "저장(F3)";
                BtnModify.Text = "취소";
            }
            else if (BtnNew.Text.Equals("저장(F3)"))
            {
                if(BtnSave_Click())
                {
                    BtnNew.Text = "신규(F1)";
                    BtnModify.Text = "수정(F2)";
                }
            }
        }

        public string REMARK_RESULT = string.Empty;
        private bool BtnSave_Click()
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return false;
            }
            
            string sJunpyoId = TxtJunpyoId.EditValue?.ToString();
            string sJDate = DateEditMes.EditValue?.ToString().Substring(0, 10);
            string sSeq = TxtSeq.EditValue?.ToString();
            string sKeratype = RdGbSelect.EditValue?.ToString();
            string sCarNo = TxtCarNo.EditValue?.ToString();

            string sDealerCd = LkupEditDealerNM.EditValue?.ToString();
            string sDealerNM = LkupEditDealerNM.Text;

            string sGradeCd = LkupEditGrade.EditValue?.ToString();
            if (string.IsNullOrEmpty(sGradeCd))
            {
                sGradeCd = "0";
            }
            string sGradeNm = LkupEditGrade.Text;

            string sInspector = LkupEditInspector.EditValue?.ToString();
            if (string.IsNullOrEmpty(sInspector))
            {
                sInspector = FmMainToolBar2.UserID;
            }
            string sInspNote = TxtInsptNote.EditValue?.ToString();
            string sTotWeight = TxtTotWeight.EditValue?.ToString();
            if (string.IsNullOrEmpty(sTotWeight))
            {
                sTotWeight = "0";
            }
            
            string sEmptyWeight = TxtEmptyWeight.EditValue?.ToString();
            if (string.IsNullOrEmpty(sEmptyWeight))
            {
                sEmptyWeight = "0";
            }
            
            string sActualWeight = TxtActualWeight.EditValue?.ToString();
            if (string.IsNullOrEmpty(sActualWeight))
            {
                sActualWeight = "0";
            }

            string sLossWeight = TxtLossWeight.EditValue?.ToString();
            if (string.IsNullOrEmpty(sLossWeight))
            {
                sLossWeight = "0";
            }

            string sInpectNote = TxtInsptNote.EditValue?.ToString();
            string sAcptWeight = TxtAcceptWeight.EditValue?.ToString();
            if (string.IsNullOrEmpty(sAcptWeight))
            {
                sAcptWeight = "0";
            }
            
            if (string.IsNullOrEmpty(sJDate))
            {
                XtraMessageBox.Show("계량일자를 선택하세요.");
                return false;
            }
            else if (string.IsNullOrEmpty(sSeq))
            {
                XtraMessageBox.Show("순번이 존재하지 않습니다.");
                return false;
            }
            else if (string.IsNullOrEmpty(sKeratype))
            {
                XtraMessageBox.Show("거래구분을 선택하세요.");
                return false;
            }
            else if (string.IsNullOrEmpty(sCarNo))
            {
                XtraMessageBox.Show("차량번호를 입력하세요.");
                return false;
            }
            else if (string.IsNullOrEmpty(sDealerCd))
            {
                XtraMessageBox.Show("거래처를 입력하세요.");
                return false;
            }

            if (!string.IsNullOrEmpty(sJunpyoId))
            {
                /*
                 * 요청일자 : 2020-10-29 현업요청
                 * 적용일자 : 2020-10-30
                 * 품명을 지워서 저장하는 경우도 존재하기 때문에 필수입력체크에서 제외
                 * 
                 */

                #region[이전 코드]

                //if (string.IsNullOrEmpty(sGradeCd))
                //{
                //    XtraMessageBox.Show("품명을 입력하세요.");
                //    return false;
                //}
                //else if (string.IsNullOrEmpty(sTotWeight) && string.IsNullOrEmpty(sEmptyWeight))
                //{
                //    XtraMessageBox.Show("총중량 혹은 공차중량을 입력하세요");
                //    return false;
                //}

                #endregion[이전 코드]

                if (string.IsNullOrEmpty(sTotWeight) && string.IsNullOrEmpty(sEmptyWeight))
                {
                    XtraMessageBox.Show("총중량 혹은 공차중량을 입력하세요");
                    return false;
                }
            }

            string sEntId = FmMainToolBar2.drUser["USRCD"].ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT INSANO");
            strSql.AppendLine("   FROM ZUSRLST ");
            strSql.AppendLine("  WHERE USRCD = " + sEntId + " ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count == 0 || dt.Rows[0]["INSANO"].Equals("****"))
            {
                sEntId = "0";
            }
            else
            {
                sEntId = dt.Rows[0]["INSANO"].ToString();
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;
                
                if (!string.IsNullOrEmpty(sJunpyoId))
                {
                    //데이터 존재 시 수정 -> 수정사유 입력처리
                    strSql.Clear();
                    strSql.AppendLine(" SELECT COUNT(1) AS CNT ");
                    strSql.AppendLine("   FROM MESURING A ");
                    strSql.AppendLine("  WHERE A.JUNPYOID = " + sJunpyoId + " ");

                    DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    string sLogGb = "";
                    if (dtChk.Rows.Count > 0)
                    {
                        double dCnt = Convert.ToDouble(dtChk.Rows[0]["CNT"]);
                        if (dCnt > 0)
                        {
                            IN08001F01 frm = new IN08001F01();
                            frm.PARENT_FORM = this;
                            frm.JUNPYO_ID = sJunpyoId;

                            if (frm.ShowDialog() == DialogResult.OK)
                            {
                                sLogGb = "1"; // LOG INSERT 조건 -> 1 : 입력, "" : 미입력
                            }
                        }
                    }
                    
                }

                /*
                 * 2021-03-16
                 * Reference Key : #0002
                 * 로그 추가
                 */
                #region[로그로직]

                string sJunpyoID_Temp = string.IsNullOrEmpty(sJunpyoId) ? "0000" : sJunpyoId;

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT A.J_DATE ");
                strSql.AppendLine("      , A.KERATYPE ");
                strSql.AppendLine("      , A.SUN ");
                strSql.AppendLine("      , A.J_BNUM ");
                strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN MAIPCHERID ELSE J_ASSIGNID END AS COMPANY_ID ");
                strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN MAIPCHER ELSE J_COMPANY END AS COMPANY ");
                strSql.AppendLine("      , A.J_SERIAL ");
                strSql.AppendLine("      , A.GUBUN1 ");
                strSql.AppendLine("      , A.GUMSU_SERIAL ");
                strSql.AppendLine("      , B.EMP_NM ");
                strSql.AppendLine("      , A.GUMSUBIGO ");
                strSql.AppendLine("      , A.FIRSTWEIGHT");
                strSql.AppendLine("      , A.SECONDWEIGHT ");
                strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN IWEIGHT ELSE OWEIGHT END AS WEIGHT ");
                strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN ICHAGAM ELSE OCHAGAM END AS CHAGAM ");
                strSql.AppendLine("      , A.J_STATE ");
                strSql.AppendLine("   FROM MESURING A ");
                strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B ");
                strSql.AppendLine("     ON A.GUMSU_SERIAL = B.EMP_ID ");
                strSql.AppendLine("  WHERE A.JUNPYOID = " + sJunpyoID_Temp + " ");

                DataTable dtLog = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                string sSTD_COLS = string.Empty;
                string sLogMsg = string.Empty;
                string sREF_RMK = string.Empty;
                int iLogCnt = 0;
                int iCCNT = 0;
                if (dtLog.Rows.Count > 0)
                {
                    string sPRV_J_DATE = dtLog.Rows[0]["J_DATE"]?.ToString().Substring(0, 10);
                    string sCur_J_DATE = DateEditMes.EditValue?.ToString().Substring(0, 10);
                    string sPRV_KERATYPE = dtLog.Rows[0]["KERATYPE"]?.ToString();
                    string sPRV_SUN = dtLog.Rows[0]["SUN"]?.ToString();
                    string sPRV_J_BNUM = dtLog.Rows[0]["J_BNUM"]?.ToString();
                    string sPRV_COMPANY_ID = dtLog.Rows[0]["COMPANY_ID"]?.ToString();
                    string sPRV_COMPANY = dtLog.Rows[0]["COMPANY"]?.ToString();
                    string sPRV_J_SERIAL = dtLog.Rows[0]["J_SERIAL"]?.ToString();
                    string sPRV_GUBUN1 = dtLog.Rows[0]["GUBUN1"]?.ToString();
                    string sPRV_GUMSU_SERIAL = dtLog.Rows[0]["GUMSU_SERIAL"]?.ToString();
                    string sPRV_EMP_NM = dtLog.Rows[0]["EMP_NM"]?.ToString();
                    string sPRV_GUMSUBIGO = dtLog.Rows[0]["GUMSUBIGO"]?.ToString();
                    double dPRV_FIRSTWEIGHT = 0;
                    double.TryParse(dtLog.Rows[0]["FIRSTWEIGHT"]?.ToString(), out dPRV_FIRSTWEIGHT);
                    
                    double dPRV_SECONDWEIGHT = 0;
                    double.TryParse(dtLog.Rows[0]["SECONDWEIGHT"]?.ToString(), out dPRV_SECONDWEIGHT);

                    double dPRV_CHAGAM = 0;
                    double.TryParse(dtLog.Rows[0]["CHAGAM"]?.ToString(), out dPRV_CHAGAM);

                    string sPRV_J_STATE = dtLog.Rows[0]["J_STATE"]?.ToString();

                    sSTD_COLS += string.Format("{0}/{1}/{3}/순번:{2}/차번:{4}"
                        , sPRV_J_DATE, sPRV_KERATYPE, sPRV_SUN, sPRV_COMPANY, sPRV_J_BNUM);

                    if (!sPRV_J_DATE.Equals(sCur_J_DATE))
                    {
                        iLogCnt++;
                        if(iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("계근일자 : {0} ▶ {1}", sPRV_J_DATE, sCur_J_DATE);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 계근일자 : {0} ▶ {1}", sPRV_J_DATE, sCur_J_DATE);
                        }
                    }

                    if (!sPRV_KERATYPE.Equals(sKeratype))
                    {
                        iLogCnt++;
                        if (iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("거래구분 : {0} ▶ {1}", sPRV_KERATYPE, sKeratype);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 거래구분 : {0} ▶ {1}", sPRV_KERATYPE, sKeratype);
                        }
                    }

                    if (!sPRV_SUN.Equals(sSeq))
                    {
                        iLogCnt++;
                        if (iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("순번 : {0} ▶ {1}", sPRV_SUN, sSeq);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 순번 : {0} ▶ {1}", sPRV_SUN, sSeq);
                        }
                    }

                    if (!sPRV_J_BNUM.Equals(sCarNo))
                    {
                        iLogCnt++;
                        if (iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("차번 : {0} ▶ {1}", sPRV_J_BNUM, sCarNo);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 차번 : {0} ▶ {1}", sPRV_J_BNUM, sCarNo);
                        }
                    }

                    if (!sPRV_COMPANY_ID.Equals(sDealerCd))
                    {
                        iLogCnt++;
                        if (iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("업체명 : {0} ▶ {1}", sPRV_COMPANY, sDealerNM);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 업체명 : {0} ▶ {1}", sPRV_COMPANY, sDealerNM);
                        }
                    }

                    if (!sPRV_J_SERIAL.Equals(sGradeCd))
                    {
                        iLogCnt++;
                        if (iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("등급 : {0} ▶ {1}", sPRV_GUBUN1, sGradeNm);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 등급 : {0} ▶ {1}", sPRV_GUBUN1, sGradeNm);
                        }
                    }

                    if (!sPRV_GUMSU_SERIAL.Equals(sInspector))
                    {
                        iLogCnt++;
                        if (iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("검수자 : {0} ▶ {1}", sPRV_EMP_NM, LkupEditInspector.Text);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 검수자 : {0} ▶ {1}", sPRV_EMP_NM, LkupEditInspector.Text);
                        }
                    }

                    if (!sPRV_GUMSUBIGO.Equals(sInspNote))
                    {
                        iLogCnt++;
                        if (iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("검수비고 : {0} ▶ {1}", sPRV_GUMSUBIGO, sInspNote);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 검수비고 : {0} ▶ {1}", sPRV_GUMSUBIGO, sInspNote);
                        }
                    }

                    double dTotWeight = 0;
                    double.TryParse(sTotWeight, out dTotWeight);
                    if (dPRV_SECONDWEIGHT != dTotWeight)
                    {
                        iLogCnt++;
                        if (iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("총중량 : {0} ▶ {1}", dPRV_SECONDWEIGHT, dTotWeight);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 총중량 : {0} ▶ {1}", dPRV_SECONDWEIGHT, dTotWeight);
                        }
                    }

                    double dEmptyWeight = 0;
                    double.TryParse(sEmptyWeight, out dEmptyWeight);
                    if (dPRV_FIRSTWEIGHT != dEmptyWeight)
                    {
                        iLogCnt++;
                        if (iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("공차중량 : {0} ▶ {1}", dPRV_FIRSTWEIGHT, dEmptyWeight);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 공차중량 : {0} ▶ {1}", dPRV_FIRSTWEIGHT, dEmptyWeight);
                        }
                    }

                    double dLossWeight = 0;
                    double.TryParse(sLossWeight, out dLossWeight);
                    if (dPRV_CHAGAM != dLossWeight)
                    {
                        iLogCnt++;
                        if (iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("감량 : {0} ▶ {1}", dPRV_CHAGAM, dLossWeight);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 감량 : {0} ▶ {1}", dPRV_CHAGAM, dLossWeight);
                        }
                    }

                    string sLossCause = TxtLossCause.EditValue?.ToString();
                    if (!sPRV_J_STATE.Equals(sLossCause))
                    {
                        iLogCnt++;
                        if (iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("감량사유 : {0} ▶ {1}", sPRV_J_STATE, sLossCause);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 감량사유 : {0} ▶ {1}", sPRV_J_STATE, sLossCause);
                        }
                    }

                    if (!string.IsNullOrEmpty(REMARK_RESULT))
                    {
                        iLogCnt++;
                        if (iCCNT++ == 0)
                        {
                            sLogMsg += string.Format("수정사유 : {0}", REMARK_RESULT);
                        }
                        else
                        {
                            sLogMsg += string.Format(" | 수정사유 : {0}", REMARK_RESULT);
                        }
                    }
                    
                    /*
                     * #0004
                     */
                    sREF_RMK += string.Format("TABLE : MESURING / JUNPYOID : {0}", sJunpyoId);
                    sLogMsg = sLogMsg.Length > 500 ? sLogMsg.Substring(0, 500) : sLogMsg;

                    /*
                     * #0005
                     */
                    if (dPRV_CHAGAM != dLossWeight)
                    {
                        StringBuilder sbPrv = new StringBuilder();

                        sbPrv.Clear();
                        sbPrv.AppendFormat(" ");
                        sbPrv.AppendFormat(" INSERT INTO MESURING_SEQ ");
                        sbPrv.AppendFormat("           ( JUNPYOID ");
                        sbPrv.AppendFormat("           , CHG_SEQ ");
                        sbPrv.AppendFormat("           , CHAGAM ");
                        sbPrv.AppendFormat("           , OCCUR_DATE ");
                        sbPrv.AppendFormat("           , ENT_ID ) ");
                        sbPrv.AppendFormat("     VALUES( {0} ", sJunpyoId);
                        sbPrv.AppendFormat("           , ( SELECT ISNULL(MAX(X1.CHG_SEQ), 0) + 1 AS CHG_SEQ ");
                        sbPrv.AppendFormat("                 FROM MESURING_SEQ AS X1 ");
                        sbPrv.AppendFormat("                WHERE X1.JUNPYOID = {0} )  ", sJunpyoId);
                        sbPrv.AppendFormat("           , {0} ", dPRV_CHAGAM);
                        sbPrv.AppendFormat("           , CONVERT(VARCHAR(19),GETDATE(),20) ");
                        sbPrv.AppendFormat("           , {0} ); ", FmMainToolBar2.UserID);

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = sbPrv.ToString();
                        cmd.Parameters.Clear();
                        cmd.ExecuteNonQuery();
                    }
                }

                #endregion[로그로직]

                if (!string.IsNullOrEmpty(sJunpyoId)) //2차 계근
                {
                    /*
                     * 2021-01-11 (현업요청) 
                     * 입고 1차 계근 -> 공차중량 : SecondWeight / 대지중량 : FirstWeight
                     * 출고 1차 계근 -> 공차중량 : FirstWeight / 대지중량 : SecondWeight
                     * 공차중량과 대지중량은 서로 값을 비교하여 입출고에 따라 다르게 설정
                     * 예 ) 입고 시 공차중량 9000 / 대지중량 10000 정상
                     *      입고 시 공차중량 9000 / 대지중량 5000 -> 중량값 뒤바꾸어 자동 수정
                     */
                    int iTotWeight = 0;
                    int iEmptyWeight = 0;

                    string sIDANGA = "0";
                    string sIKONGKEP = "0";
                    string sODANGA = "0";
                    string sOKONGKEP = "0";

                    if (sKeratype.Equals("입고"))
                    {
                        int.TryParse(sEmptyWeight, out iEmptyWeight);
                        int.TryParse(sTotWeight, out iTotWeight);
                        if(iEmptyWeight > iTotWeight)
                        {
                            int iTemp = Convert.ToInt32(sEmptyWeight);
                            sEmptyWeight = iTotWeight.ToString();
                            sTotWeight = iTemp.ToString();

                            sActualWeight = Convert.ToInt32(iEmptyWeight - iTotWeight).ToString();
                            sAcptWeight = (Convert.ToInt32(iEmptyWeight - iTotWeight) - Convert.ToInt32(sLossWeight)).ToString();
                        }
                        else
                        {
                            //int iTemp = Convert.ToInt32(sTotWeight);
                            //sTotWeight = sEmptyWeight.ToString();
                            //sEmptyWeight = iTemp.ToString();

                            sActualWeight = Convert.ToInt32(iTotWeight - iEmptyWeight).ToString();
                            sAcptWeight = (Convert.ToInt32(iTotWeight - iEmptyWeight) - Convert.ToInt32(sLossWeight)).ToString();
                        }

                        //2021-01-16 주석처리
                        string sSecMesTime = "";//총중량 시각
                        if (DateTime.TryParse(TimeEditSecond.EditValue?.ToString(), out DateTime secTime))
                        {
                            sSecMesTime = secTime.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                          
                        //if (string.IsNullOrEmpty(sTotWeight) || sTotWeight.Equals("0"))
                        //{
                        //    sSecMesTime = "";
                        //}

                        string sFirstMesTime = ""; //공차중량 시각
                        if(DateTime.TryParse(TimeEditFirst.EditValue?.ToString(), out DateTime firstTime))
                        {
                            sFirstMesTime = firstTime.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        //if (string.IsNullOrEmpty(sEmptyWeight) || sEmptyWeight.Equals("0"))
                        //{
                        //    sFirstMesTime = "";
                        //}

                        int iKYERYANG12 = 1;

                        //입출고 구분이 바뀌었을 때 FirstWeight와 SecondWeight 바꿈
                        Dictionary<string, string> dicParams = new Dictionary<string, string>();
                        dicParams.Clear();
                        dicParams.Add("KERATYPE", sKeratype);
                        dicParams.Add("JUNPYOID", sJunpyoId);

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" SELECT CASE WHEN @KERATYPE = KERATYPE THEN 'Y' ELSE 'N' END AS YN  ");
                        strSql.AppendLine("   FROM MESURING ");
                        strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                        DataTable dtChk = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
                        if(dtChk.Rows.Count > 0)
                        {
                            //출고에서 입고로 변경된 경우
                            string sYN = dtChk.Rows[0]["YN"]?.ToString();
                            sYN = !string.IsNullOrEmpty(sYN) ? sYN : string.Empty;
                            if (sYN.Equals("N"))
                            {
                                //string sTemp = sEmptyWeight;
                                //sEmptyWeight = sTotWeight;
                                //sTotWeight = sTemp;

                                string sTemp = sFirstMesTime;
                                sFirstMesTime = sSecMesTime;
                                sSecMesTime = sTemp;

                                /*
                                 * 2021-01-12 
                                 * 계근프로그램에서 수정 시 영업에서 단가입력되는 부분이 날라가는 현상발생
                                 * 그러나 입출고가 바뀔 시 컬럼 또한 변경해야하기때문에 해당 데이터 가져와 세팅
                                 */
                                dicParams.Clear();
                                dicParams.Add("JUNPYOID", sJunpyoId);

                                strSql.Clear();
                                strSql.AppendLine(" ");
                                strSql.AppendLine(" SELECT * ");
                                strSql.AppendLine("   FROM MESURING ");
                                strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                                DataTable dtPrv = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);

                                if (dtPrv.Rows.Count > 0)
                                {
                                    //해당 변수를 입고쪽 컬럼에 세팅
                                    sIDANGA = string.IsNullOrEmpty(dtPrv.Rows[0]["ODANGA"]?.ToString()) ? "0" : dtPrv.Rows[0]["ODANGA"]?.ToString();
                                    sIKONGKEP = string.IsNullOrEmpty(dtPrv.Rows[0]["OKONGKEP"]?.ToString()) ? "0" : dtPrv.Rows[0]["OKONGKEP"]?.ToString();
                                    string sKyeryang12 = string.IsNullOrEmpty(dtPrv.Rows[0]["OKONGKEP"]?.ToString()) ? "1" : dtPrv.Rows[0]["OKONGKEP"]?.ToString();
                                    iKYERYANG12 = Convert.ToInt32(sKyeryang12);
                                }
                            }
                        }
                        
                        //검수프로그램에서 2차계근을 위하여 Weight 값이 0이어야 하므로
                        //공차 / 대지중량이 입력되지 않을 경우 weight 및 iweight 컬럼에 0 처리
                        if (!(iEmptyWeight > 0 && iTotWeight > 0))
                        {
                            sActualWeight = "0";
                            sAcptWeight = "0";
                            iKYERYANG12 = 1;
                        }
                        else
                        {
                            iKYERYANG12 = 2;
                        }

                        /*
                         * 2021-01-12(화)
                         * WEIGHT는 무인계근시스템에서 차량정보 입력 시 1/2차 계근을 구분짓는 컬럼이므로
                         * 등급코드(J_SERIAL)이 입력되어있을 시 WEIGHT로 세팅하여 무인계근에서 2차계근 처리으로 판정
                         * 추후 kyeryang12컬럼으로 판단하여야함 
                         */
                        if(string.IsNullOrEmpty(sGradeCd) || sGradeCd.Equals("0"))
                        {
                            sActualWeight = "0";
                            iKYERYANG12 = 1;
                        }

                        if (iKYERYANG12 == 1)
                        {
                            string sTemp = string.Empty;
                            if (string.IsNullOrEmpty(sFirstMesTime))
                            {
                                sFirstMesTime = sSecMesTime;
                            }
                            else if(string.IsNullOrEmpty(sSecMesTime))
                            {
                                sSecMesTime = sFirstMesTime;
                            }
                        }

                        /*
                         * 2021-02-22
                         * MESURING 테이블에 업체담당자 컬럼 추가에 따라 해당 컬럼에 UPDATE 로직 추가
                         */
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE MESURING ");
                        strSql.AppendLine("    SET KERATYPE = '" + sKeratype + "' ");
                        strSql.AppendLine("      , MAIPCHERID = " + sDealerCd + " ");
                        strSql.AppendLine("      , MAIPCHER = ( SELECT DEALER_NM FROM ACC_DEALER_CD WHERE DEALER_CD = " + sDealerCd + " ) ");
                        strSql.AppendLine("      , SUN = " + sSeq + " ");
                        strSql.AppendLine("      , J_DATE = '" + sJDate + "' ");
                        strSql.AppendLine("      , J_ASSIGNID = 0 ");
                        strSql.AppendLine("      , J_COMPANY = '' ");
                        if (string.IsNullOrEmpty(sFirstMesTime))
                        {
                            strSql.AppendLine("      , FIRSTTIME = NULL ");
                        }
                        else
                        {
                            strSql.AppendLine("      , FIRSTTIME = '" + sFirstMesTime + "' ");
                        }

                        if (string.IsNullOrEmpty(sSecMesTime))
                        {
                            strSql.AppendLine("      , SECONDTIME = NULL ");
                        }
                        else
                        {
                            strSql.AppendLine("      , SECONDTIME = '" + sSecMesTime + "' ");
                        }
                        strSql.AppendLine("      , FIRSTWEIGHT = " + sEmptyWeight + " ");
                        strSql.AppendLine("      , SECONDWEIGHT = " + sTotWeight + " ");
                        strSql.AppendLine("      , WEIGHT = " + sActualWeight + " ");
                        strSql.AppendLine("      , J_SERIAL = " + sGradeCd + " ");
                        strSql.AppendLine("      , GUBUN1 = ( SELECT GUBUN1 FROM JAJAE WHERE J_SERIAL = " + sGradeCd + " ) ");
                        strSql.AppendLine("      , J_BNUM = '" + sCarNo + "' ");
                        strSql.AppendLine("      , K_NAME = '상품' ");
                        strSql.AppendLine("      , ICHAGAM = " + sLossWeight + " ");
                        strSql.AppendLine("      , OCHAGAM = 0 ");
                        strSql.AppendLine("      , IGAMGA = 0 ");
                        strSql.AppendLine("      , OGAMGA = 0 ");
                        strSql.AppendLine("      , IWEIGHT = " + sAcptWeight + " ");
                        strSql.AppendLine("      , OWEIGHT = 0 ");
                        /*
                         * 2021-03-11
                         * #0001
                         * 계근변경 시 단가가 날라가는 현상 수정
                         */
                        strSql.AppendLine("      , IDANGA = CASE WHEN " + sIDANGA + " = 0 THEN IDANGA ELSE " + sIDANGA + " END ");
                        strSql.AppendLine("      , ODANGA = 0 ");
                        strSql.AppendLine("      , IKONGKEP = CASE WHEN " + sIKONGKEP + " = 0 THEN IKONGKEP ELSE " + sIKONGKEP + " END ");
                        strSql.AppendLine("      , OKONGKEP = 0 ");
                        strSql.AppendLine("      , USERCODE = " + sEntId + " ");
                        strSql.AppendLine("      , BUSEOCODE = 0 ");
                        strSql.AppendLine("      , P_ID = 100 ");
                        strSql.AppendLine("      , J_GARAGE = 1 ");
                        strSql.AppendLine("      , J_ID = 0 ");
                        strSql.AppendLine("      , KYERYANG12 = " + iKYERYANG12 + " ");
                        strSql.AppendLine("      , DRIVER_INOUT = 0 ");
                        strSql.AppendLine("      , AGREE_DATE = AGREE_DATE ");
                        strSql.AppendLine("      , J_STATE2 = J_STATE2 ");
                        strSql.AppendLine("      , REGIONSTART = '' ");
                        strSql.AppendLine("      , REGIONDEST = '' ");
                        strSql.AppendLine("      , TRANSPORTDANGA = TRANSPORTDANGA ");
                        strSql.AppendLine("      , TRANSPORTKUMAK = TRANSPORTKUMAK ");
                        strSql.AppendLine("      , TRANSPORTPERSON = '' ");
                        strSql.AppendLine("      , TRANSPORTCUSTOM = '' ");
                        strSql.AppendLine("      , TRANSPORTC_SERIAL = 0 ");
                        strSql.AppendLine("      , TRANSPORTBIGO = '' ");
                        strSql.AppendLine("      , CUSTOMWEIGHT = " + sActualWeight + "");
                        strSql.AppendLine("      , LOSSWEIGHT = LOSSWEIGHT ");
                        strSql.AppendLine("      , TRANSPORTJUNGSAN = 0 ");
                        strSql.AppendLine("      , MAGAM_FLAG = '0' ");
                        strSql.AppendLine("      , IPCHULGO_MAIPID = IPCHULGO_MAIPID ");
                        strSql.AppendLine("      , IPCHULGO_MACHULID = IPCHULGO_MACHULID ");
                        strSql.AppendLine("      , WEIGHT_GUBUN = 0 ");
                        strSql.AppendLine("      , LENGTHSID = 0 ");
                        strSql.AppendLine("      , BLNO = '' ");
                        strSql.AppendLine("      , CONTAINERNO = '' ");
                        strSql.AppendLine("      , DAMAGE = 0 ");
                        strSql.AppendLine("      , GUMSUBIGO = '" + sInspNote + "'   ");
                        if (string.IsNullOrEmpty(sInspector))
                        {
                            strSql.AppendLine(" , GUMSU_SERIAL = NULL");
                        }
                        else
                        {
                            strSql.AppendLine(" , GUMSU_SERIAL = " + sInspector + "");
                        }
                        strSql.AppendLine("      , EDIT_GB = 'B' "); // 수정구분(트리거 위함) -> A: 마감 및 업로드, B: 계근프로그램, C: 단가입력
                        strSql.AppendLine("      , EDIT_RMK = '" + REMARK_RESULT + "' ");
                        //strSql.AppendLine("      , CHRG_ID = (SELECT CHRG_ID FROM ACC_DEALER_CD WHERE DEALER_CD = '" + sDealerCd + "')");
                        strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoId + " ");

                        //LogInsert
                        //Reference : #0002
                        if (iLogCnt > 0)
                        {
                            StringBuilder strLog = new StringBuilder();

                            strLog.Clear();
                            strLog.AppendLine(" ");
                            strLog.AppendLine(" INSERT INTO ZSYS_LOG ");
                            strLog.AppendLine("           ( OCCUR_DATE ");
                            strLog.AppendLine("           , USRCD ");
                            strLog.AppendLine("           , LOG_SEQ ");
                            strLog.AppendLine("           , EDIT_KIND ");
                            strLog.AppendLine("           , PGM_ID ");
                            strLog.AppendLine("           , ACS_IP ");
                            strLog.AppendLine("           , STD_COLS ");
                            strLog.AppendLine("           , REF_RMK ");
                            strLog.AppendLine("           , EDIT_RMK ) ");
                            strLog.AppendLine("     VALUES( @OCCUR_DATE ");
                            strLog.AppendLine("           , @USRCD ");
                            strLog.AppendLine("           , ( SELECT ISNULL(MAX(X1.LOG_SEQ), 0) + 1 FROM ZSYS_LOG X1 WHERE X1.OCCUR_DATE = @OCCUR_DATE AND X1.USRCD = @USRCD ) ");
                            strLog.AppendLine("           , 'U' ");
                            strLog.AppendLine("           , @PGM_ID ");
                            strLog.AppendLine("           , @ACS_IP ");
                            strLog.AppendLine("           , @STD_COLS ");
                            strLog.AppendLine("           , @REF_RMK ");
                            strLog.AppendLine("           , @EDIT_RMK ) ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strLog.ToString();
                            cmd.Parameters.Clear();
                            cmd.Parameters.AddWithValue("@OCCUR_DATE", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                            cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                            cmd.Parameters.AddWithValue("@PGM_ID", this.Name);
                            cmd.Parameters.AddWithValue("@ACS_IP", ComnEtcFunc.GetLocalIP());
                            cmd.Parameters.AddWithValue("@STD_COLS", sSTD_COLS);
                            cmd.Parameters.AddWithValue("@REF_RMK", sREF_RMK);
                            cmd.Parameters.AddWithValue("@EDIT_RMK", sLogMsg);
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                        }
                    }
                    else if (sKeratype.Equals("출고"))
                    {
                        int.TryParse(sEmptyWeight, out iEmptyWeight);
                        int.TryParse(sTotWeight, out iTotWeight);

                        int iKYERYANG12 = 1;
                        if(iTotWeight == 0 || iEmptyWeight == 0)
                        {
                            if (iTotWeight > iEmptyWeight)
                            { //5000, 2000
                                int iTemp = Convert.ToInt32(sEmptyWeight);
                                sEmptyWeight = iTotWeight.ToString();
                                sTotWeight = iTemp.ToString();
                                //int iTemp = Convert.ToInt32(iTotWeight);
                                //sTotWeight = sEmptyWeight.ToString();
                                //sEmptyWeight = iTemp.ToString();

                                sActualWeight = Convert.ToInt32(iEmptyWeight - iTotWeight).ToString();
                                sAcptWeight = (Convert.ToInt32(iEmptyWeight - iTotWeight) - Convert.ToInt32(sLossWeight)).ToString();
                            }
                            else
                            {
                                sActualWeight = Convert.ToInt32(iTotWeight - iEmptyWeight).ToString();
                                sAcptWeight = (Convert.ToInt32(iTotWeight - iEmptyWeight) - Convert.ToInt32(sLossWeight)).ToString();
                            }
                        }
                        
                        string sSecMesTime = "";  //총중량 시각
                        if(DateTime.TryParse(TimeEditSecond.EditValue?.ToString(), out DateTime secTime))
                        {
                            sSecMesTime = secTime.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        //if (string.IsNullOrEmpty(sEmptyWeight) || sEmptyWeight.Equals("0"))
                        //{
                        //    sSecMesTime = "";
                        //}

                        string sFirstMesTime = ""; //공차중량 시각
                        if(DateTime.TryParse(TimeEditFirst.EditValue?.ToString(), out DateTime firstTime))
                        {
                            sFirstMesTime = firstTime.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        //if (string.IsNullOrEmpty(sTotWeight) || sTotWeight.Equals("0"))
                        //{
                        //    sFirstMesTime = "";
                        //}

                        //입출고 구분이 바뀌었을 때 FirstWeight와 SecondWeight 바꿈
                        Dictionary<string, string> dicParams = new Dictionary<string, string>();
                        dicParams.Clear();
                        dicParams.Add("KERATYPE", sKeratype);
                        dicParams.Add("JUNPYOID", sJunpyoId);

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" SELECT CASE WHEN @KERATYPE = KERATYPE THEN 'Y' ELSE 'N' END AS YN  ");
                        strSql.AppendLine("   FROM MESURING ");
                        strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID ");

                        DataTable dtChk = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
                        if (dtChk.Rows.Count > 0)
                        {
                            string sYN = dtChk.Rows[0]["YN"]?.ToString();
                            sYN = !string.IsNullOrEmpty(sYN) ? sYN : string.Empty;
                            if (sYN.Equals("N"))
                            {
                                //string sTemp = sEmptyWeight;
                                //sEmptyWeight = sTotWeight;
                                //sTotWeight = sTemp;

                                string sTemp = sSecMesTime;
                                sSecMesTime = sFirstMesTime;
                                sFirstMesTime = sTemp;

                                /*
                                 * 2021-01-12 
                                 * 계근프로그램에서 수정 시 영업에서 단가입력되는 부분이 날라가는 현상발생
                                 * 그러나 입출고가 바뀔 시 컬럼 또한 변경해야하기때문에 해당 데이터 가져와 세팅
                                 */
                                dicParams.Clear();
                                dicParams.Add("JUNPYOID", sJunpyoId);

                                strSql.Clear();
                                strSql.AppendLine(" ");
                                strSql.AppendLine(" SELECT * ");
                                strSql.AppendLine("   FROM MESURING ");
                                strSql.AppendLine("  WHERE JUNPYOID = "+ sJunpyoId);

                                //SqlConnection dbConn = ComnEtcFunc.DbConn();
                                //DataTable dtPrv = ComnEtcFunc.GetDataTable(strSql.ToString(), dicParams);
                                //ComnEtcFunc.DbDisConn(dbConn);

                                DataTable dtPrv = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);

                                if (dtPrv.Rows.Count > 0)
                                {
                                    //해당 변수를 입고쪽 컬럼에 세팅
                                    sODANGA = string.IsNullOrEmpty(dtPrv.Rows[0]["IDANGA"]?.ToString()) ? "0" : dtPrv.Rows[0]["IDANGA"]?.ToString();
                                    sOKONGKEP = string.IsNullOrEmpty(dtPrv.Rows[0]["IKONGKEP"]?.ToString()) ? "0" : dtPrv.Rows[0]["IKONGKEP"]?.ToString();
                                }

                                //if (iEmptyWeight == 0)
                                //{   //2000, 5000
                                //    int iTemp = Convert.ToInt32(iTotWeight);
                                //    sTotWeight = sEmptyWeight.ToString();
                                //    sEmptyWeight = iTemp.ToString();

                                //    sActualWeight = Convert.ToInt32(iTotWeight - iEmptyWeight).ToString();
                                //    sAcptWeight = (Convert.ToInt32(iTotWeight - iEmptyWeight) - Convert.ToInt32(sLossWeight)).ToString();
                                //}
                            }
                        }

                        //검수프로그램에서 2차계근을 위하여 Weight 값이 0이어야 하므로
                        //공차 / 대지중량이 입력되지 않을 경우 weight 및 iweight 컬럼에 0 처리
                        if (!(iEmptyWeight > 0 && iTotWeight > 0))
                        {
                            sActualWeight = "0";
                            sAcptWeight = "0";
                            iKYERYANG12 = 1;
                        }
                        else
                        {
                            iKYERYANG12 = 2;
                        }

                        /*
                         * 2021-01-12(화)
                         * WEIGHT는 무인계근시스템에서 차량정보 입력 시 1/2차 계근을 구분짓는 컬럼이므로
                         * 등급코드(J_SERIAL)이 입력되어있을 시 WEIGHT로 세팅하여 무인계근에서 2차계근 처리으로 판정
                         * 추후 kyeryang12컬럼으로 판단하여야함 
                         */
                        if (string.IsNullOrEmpty(sGradeCd) || sGradeCd.Equals("0"))
                        {
                            sActualWeight = "0";
                            iKYERYANG12 = 1;
                        }

                        if(iKYERYANG12 == 1)
                        {
                            //int iTemp = Convert.ToInt32(sEmptyWeight);
                            //sEmptyWeight = iTotWeight.ToString();
                            //sTotWeight = iTemp.ToString();

                            string sTemp = string.Empty;
                            if (string.IsNullOrEmpty(sFirstMesTime))
                            {
                                sFirstMesTime = sSecMesTime;
                            }
                            else if(string.IsNullOrEmpty(sSecMesTime))
                            {
                                sSecMesTime = sFirstMesTime;
                            }
                        }

                        /*
                         * 2021-02-22
                         * MESURING 테이블에 업체담당자 컬럼 추가에 따라 해당 컬럼에 UPDATE 로직 추가
                         */
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE MESURING ");
                        strSql.AppendLine("    SET ");
                        strSql.AppendLine("        KERATYPE = '" + sKeratype + "' ");
                        strSql.AppendLine("      , J_ASSIGNID = " + sDealerCd + " ");
                        strSql.AppendLine("      , J_COMPANY = ( SELECT DEALER_NM FROM ACC_DEALER_CD WHERE DEALER_CD = " + sDealerCd + " ) ");
                        strSql.AppendLine("      , SUN = " + sSeq + " ");
                        strSql.AppendLine("      , J_DATE = '" + sJDate + "' ");
                        strSql.AppendLine("      , MAIPCHERID = 0 ");
                        strSql.AppendLine("      , MAIPCHER = '' ");
                        if (string.IsNullOrEmpty(sFirstMesTime))
                        {
                            strSql.AppendLine("      , FIRSTTIME = NULL ");
                        }
                        else
                        {
                            strSql.AppendLine("      , FIRSTTIME = '" + sFirstMesTime + "' ");
                        }

                        if (string.IsNullOrEmpty(sSecMesTime))
                        {
                            strSql.AppendLine("      , SECONDTIME = NULL ");
                        }
                        else
                        {
                            strSql.AppendLine("      , SECONDTIME = '" + sSecMesTime + "' ");
                        }
                        strSql.AppendLine("      , FIRSTWEIGHT = " + sEmptyWeight + " ");
                        strSql.AppendLine("      , SECONDWEIGHT = " + sTotWeight + " ");
                        strSql.AppendLine("      , WEIGHT = " + sActualWeight + " ");
                        strSql.AppendLine("      , J_SERIAL = " + sGradeCd + " ");
                        strSql.AppendLine("      , GUBUN1 = ( SELECT GUBUN1 FROM JAJAE WHERE J_SERIAL = " + sGradeCd + " ) ");
                        strSql.AppendLine("      , J_BNUM = '" + sCarNo + "' ");
                        strSql.AppendLine("      , K_NAME = '상품' ");
                        strSql.AppendLine("      , ICHAGAM = 0 ");
                        strSql.AppendLine("      , OCHAGAM = " + sLossWeight + " ");
                        strSql.AppendLine("      , IGAMGA = 0 ");
                        strSql.AppendLine("      , OGAMGA = 0 ");
                        strSql.AppendLine("      , IWEIGHT = 0 ");
                        strSql.AppendLine("      , OWEIGHT = " + sAcptWeight + " ");
                        strSql.AppendLine("      , IDANGA = 0 ");
                        /*
                         * 2021-03-11
                         * #0001
                         * 계근변경 시 단가가 날라가는 현상 수정
                         */
                        strSql.AppendLine("      , ODANGA = CASE WHEN " + sODANGA + " = 0 THEN ODANGA ELSE " + sODANGA + " END ");
                        strSql.AppendLine("      , IKONGKEP = 0 ");
                        strSql.AppendLine("      , OKONGKEP =  CASE WHEN " + sOKONGKEP + " = 0 THEN OKONGKEP ELSE " + sOKONGKEP + " END ");
                        strSql.AppendLine("      , USERCODE = " + sEntId + " ");
                        strSql.AppendLine("      , BUSEOCODE = 0 ");
                        strSql.AppendLine("      , P_ID = 100 ");
                        strSql.AppendLine("      , J_GARAGE = 1 ");
                        strSql.AppendLine("      , J_ID = 0 ");
                        strSql.AppendLine("      , KYERYANG12 = " + iKYERYANG12 + " ");
                        strSql.AppendLine("      , DRIVER_INOUT = 0 ");
                        strSql.AppendLine("      , AGREE_DATE = '0001-01-01' ");
                        strSql.AppendLine("      , J_STATE2 = J_STATE2 ");
                        strSql.AppendLine("      , REGIONSTART = '' ");
                        strSql.AppendLine("      , REGIONDEST = '' ");
                        strSql.AppendLine("      , TRANSPORTDANGA = 0 ");
                        strSql.AppendLine("      , TRANSPORTKUMAK = TRANSPORTKUMAK ");
                        strSql.AppendLine("      , TRANSPORTPERSON = '' ");
                        strSql.AppendLine("      , TRANSPORTCUSTOM = '' ");
                        strSql.AppendLine("      , TRANSPORTC_SERIAL = 0 ");
                        strSql.AppendLine("      , TRANSPORTBIGO = '' ");
                        strSql.AppendLine("      , CUSTOMWEIGHT = " + sActualWeight + " ");
                        strSql.AppendLine("      , LOSSWEIGHT = LOSSWEIGHT ");
                        strSql.AppendLine("      , TRANSPORTJUNGSAN = 0 ");
                        strSql.AppendLine("      , MAGAM_FLAG = MAGAM_FLAG ");
                        strSql.AppendLine("      , IPCHULGO_MAIPID = IPCHULGO_MAIPID ");
                        strSql.AppendLine("      , IPCHULGO_MACHULID = IPCHULGO_MACHULID ");
                        strSql.AppendLine("      , WEIGHT_GUBUN = 0 ");
                        strSql.AppendLine("      , LENGTHSID = 0 ");
                        strSql.AppendLine("      , BLNO = '' ");
                        strSql.AppendLine("      , CONTAINERNO = '' ");
                        strSql.AppendLine("      , DAMAGE = 0 ");
                        strSql.AppendLine("      , GUMSUBIGO = '" + sInspNote + "' ");
                        strSql.AppendLine("      , GUMSU_SERIAL=" + sInspector + " ");
                        strSql.AppendLine("      , EDIT_GB = 'B'  "); // 수정구분(트리거 위함) -> A: 마감 및 업로드, B: 계근프로그램, C: 단가입력
                        strSql.AppendLine("      , EDIT_RMK = '" + REMARK_RESULT + "' ");
                        //strSql.AppendLine("      , CHRG_ID = (SELECT CHRG_ID FROM ACC_DEALER_CD WHERE DEALER_CD = '" + sDealerCd + "')");
                        strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoId + " ");

                        //LogInsert
                        //Reference : #0002
                        if (iLogCnt > 0)
                        {
                            StringBuilder strLog = new StringBuilder();

                            strLog.Clear();
                            strLog.AppendLine(" ");
                            strLog.AppendLine(" INSERT INTO ZSYS_LOG ");
                            strLog.AppendLine("           ( OCCUR_DATE ");
                            strLog.AppendLine("           , USRCD ");
                            strLog.AppendLine("           , LOG_SEQ ");
                            strLog.AppendLine("           , EDIT_KIND ");
                            strLog.AppendLine("           , PGM_ID ");
                            strLog.AppendLine("           , ACS_IP ");
                            strLog.AppendLine("           , STD_COLS ");
                            strLog.AppendLine("           , REF_RMK ");
                            strLog.AppendLine("           , EDIT_RMK ) ");
                            strLog.AppendLine("     VALUES( @OCCUR_DATE ");
                            strLog.AppendLine("           , @USRCD ");
                            strLog.AppendLine("           , ( SELECT ISNULL(MAX(X1.LOG_SEQ), 0) + 1 FROM ZSYS_LOG X1 WHERE X1.OCCUR_DATE = @OCCUR_DATE AND X1.USRCD = @USRCD ) ");
                            strLog.AppendLine("           , 'U' ");
                            strLog.AppendLine("           , @PGM_ID ");
                            strLog.AppendLine("           , @ACS_IP ");
                            strLog.AppendLine("           , @STD_COLS ");
                            strLog.AppendLine("           , @REF_RMK ");
                            strLog.AppendLine("           , @EDIT_RMK ) ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strLog.ToString();
                            cmd.Parameters.Clear();
                            cmd.Parameters.AddWithValue("@OCCUR_DATE", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                            cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                            cmd.Parameters.AddWithValue("@PGM_ID", this.Name);
                            cmd.Parameters.AddWithValue("@ACS_IP", ComnEtcFunc.GetLocalIP());
                            cmd.Parameters.AddWithValue("@STD_COLS", sSTD_COLS);
                            cmd.Parameters.AddWithValue("@REF_RMK", sREF_RMK);
                            cmd.Parameters.AddWithValue("@EDIT_RMK", sLogMsg);
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                        }
                    }
                }
                else //1차 계근
                {
                    if (sKeratype.Equals("입고"))
                    {
                        string sSecMesTime = "";  //총중량 시각
                        if(DateTime.TryParse(TimeEditSecond.EditValue?.ToString(), out DateTime secTime))
                        {
                            sSecMesTime = secTime.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        if (string.IsNullOrEmpty(sEmptyWeight) || sEmptyWeight.Equals("0"))
                        {
                            sSecMesTime = "";
                        }

                        string sFirstMesTime = ""; //공차중량 시각
                        if(DateTime.TryParse(TimeEditFirst.EditValue?.ToString(), out DateTime firstTime))
                        {
                            sFirstMesTime = firstTime.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        if (string.IsNullOrEmpty(sTotWeight) || sTotWeight.Equals("0"))
                        {
                            sFirstMesTime = "";
                        }

                        /*
                         * 2021-02-22
                         * MESURING 테이블에 업체담당자 컬럼 추가에 따라 해당 컬럼에 UPDATE 로직 추가
                         * -> (현업요청) MESURING은 수정하지 않는 것으로 다시 쿼리수정
                         */
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" INSERT INTO MESURING ");
                        strSql.AppendLine("           ( KERATYPE ");
                        strSql.AppendLine("           , MAIPCHERID ");
                        strSql.AppendLine("           , MAIPCHER ");
                        strSql.AppendLine("           , SUN ");
                        strSql.AppendLine("           , J_DATE ");
                        strSql.AppendLine("           , J_ASSIGNID ");
                        //strSql.AppendLine(" ");
                        strSql.AppendLine("           , FIRSTTIME      , SECONDTIME       , FIRSTWEIGHT      , SECONDWEIGHT     , WEIGHT ");
                        strSql.AppendLine("           , J_SERIAL       , GUBUN1           , J_BNUM           , K_NAME           , ICHAGAM ");
                        strSql.AppendLine("           , OCHAGAM        , IGAMGA           , OGAMGA           , IWEIGHT          , OWEIGHT ");
                        strSql.AppendLine("           , IDANGA         , ODANGA           , IKONGKEP         , OKONGKEP         , USERCODE ");
                        strSql.AppendLine("           , BUSEOCODE      , P_ID             , J_GARAGE         , J_ID             , KYERYANG12 ");
                        strSql.AppendLine("           , DRIVER_INOUT   , AGREE_DATE       , J_STATE2         , REGIONSTART      , REGIONDEST ");
                        strSql.AppendLine("           , TRANSPORTDANGA , TRANSPORTKUMAK   , TRANSPORTPERSON  , TRANSPORTCUSTOM  , TRANSPORTC_SERIAL ");
                        strSql.AppendLine("           , TRANSPORTBIGO  , CUSTOMWEIGHT     , LOSSWEIGHT       , TRANSPORTJUNGSAN , MAGAM_FLAG ");
                        strSql.AppendLine("           , IPCHULGO_MAIPID, IPCHULGO_MACHULID, WEIGHT_GUBUN     , LENGTHSID        , BLNO ");
                        strSql.AppendLine("           , CONTAINERNO    , DAMAGE           , GUMSU_SERIAL     , GUMSUBIGO        , J_CHECK");
                        //strSql.AppendLine("           , CHRG_ID ");
                        strSql.AppendLine("           ) ");
                        strSql.AppendLine("      VALUES ");
                        strSql.AppendLine("           ( ");
                        strSql.AppendLine("             '" + sKeratype + "',  " + sDealerCd + " , '" + sDealerNM + "', " + sSeq + ", '" + sJDate + "', 0 ");
                        if (string.IsNullOrEmpty(sFirstMesTime))
                        {
                            strSql.AppendLine("      , '" + sSecMesTime + "' ");
                        }
                        else
                        {
                            strSql.AppendLine("      , '" + sFirstMesTime + "'");
                        }

                        if (string.IsNullOrEmpty(sSecMesTime))
                        {
                            strSql.AppendLine("      , '" + sFirstMesTime + "' ");
                        }
                        else
                        {
                            strSql.AppendLine("      , '" + sSecMesTime + "' ");
                        }

                        strSql.AppendLine("          , " + sEmptyWeight + ", " + sTotWeight + ", " + sActualWeight + " ");
                        if (string.IsNullOrEmpty(sGradeCd))
                        {
                            strSql.AppendLine("           ,  NULL");
                        }
                        else
                        {
                            strSql.AppendLine("           ,  " + sGradeCd + "");
                        }
                        strSql.AppendLine("           , '" + sGradeNm + "', '" + sCarNo + "', '상품'                 , " + sLossWeight + " ");
                        strSql.AppendLine("           ,  0                     , 0                    , 0                   , " + sAcptWeight + ", 0 ");
                        strSql.AppendLine("           ,  0                     , 0                    , 0                   , 0                  , " + sEntId + " ");
                        strSql.AppendLine("           ,  0                     , 100                  , 1                   , 0                  , 1 ");
                        strSql.AppendLine("           ,  0                     , '0001-01-01'         , ''                  , ''                 , '' ");
                        strSql.AppendLine("           ,  0                     , 0                    , ''                  , ''                 , 0 ");
                        strSql.AppendLine("           , ''                     , " + sActualWeight + ", 0                   , 0                  , 0 ");
                        strSql.AppendLine("           ,  0                     , 0                    , 0                   , 0                  , '' ");
                        strSql.AppendLine("           , ''                     , 0 ");
                        if (string.IsNullOrEmpty(sInspector))
                        {
                            strSql.AppendLine("                    , NULL   ");
                        }
                        else
                        {
                            strSql.AppendLine("                    , " + sInspector + "   ");
                        }
                        strSql.AppendLine("                    , '" + sInspNote + "'");
                        strSql.AppendLine("          , ''  ");
                        //strSql.AppendLine("          , (SELECT CHRG_ID FROM ACC_DEALER_CD WHERE DEALER_CD = '" + sDealerCd + "') ");
                        strSql.AppendLine("          ) ");

                        #region[2021-01-14 쿼리정리 이전]

                        //strSql.Clear();
                        //strSql.AppendLine(" ");
                        //strSql.AppendLine(" INSERT INTO MESURING ");
                        //strSql.AppendLine("           ( KERATYPE       , MAIPCHERID       , MAIPCHER         , SUN              , J_DATE , J_ASSIGNID ");
                        //strSql.AppendLine("           , FIRSTTIME      , SECONDTIME       , FIRSTWEIGHT      , SECONDWEIGHT     , WEIGHT ");
                        //strSql.AppendLine("           , J_SERIAL       , GUBUN1           , J_BNUM           , K_NAME           , ICHAGAM ");
                        //strSql.AppendLine("           , OCHAGAM        , IGAMGA           , OGAMGA           , IWEIGHT          , OWEIGHT ");
                        //strSql.AppendLine("           , IDANGA         , ODANGA           , IKONGKEP         , OKONGKEP         , USERCODE ");
                        //strSql.AppendLine("           , BUSEOCODE      , P_ID             , J_GARAGE         , J_ID             , KYERYANG12 ");
                        //strSql.AppendLine("           , DRIVER_INOUT   , AGREE_DATE       , J_STATE2         , REGIONSTART      , REGIONDEST ");
                        //strSql.AppendLine("           , TRANSPORTDANGA , TRANSPORTKUMAK   , TRANSPORTPERSON  , TRANSPORTCUSTOM  , TRANSPORTC_SERIAL ");
                        //strSql.AppendLine("           , TRANSPORTBIGO  , CUSTOMWEIGHT     , LOSSWEIGHT       , TRANSPORTJUNGSAN , MAGAM_FLAG ");
                        //strSql.AppendLine("           , IPCHULGO_MAIPID, IPCHULGO_MACHULID, WEIGHT_GUBUN     , LENGTHSID        , BLNO ");
                        //strSql.AppendLine("           , CONTAINERNO    , DAMAGE           , GUMSU_SERIAL     , GUMSUBIGO        , J_CHECK");
                        //strSql.AppendLine("           ) ");
                        //strSql.AppendLine("      VALUES ");
                        //strSql.AppendLine("           ( ");
                        //strSql.AppendLine("             '" + sKeratype + "',  " + sDealerCd + " , '" + sDealerNM + "', " + sSeq + ", " + sJDate + ", 0 ");
                        //if (string.IsNullOrEmpty(sFirstMesTime))
                        //{
                        //    strSql.AppendLine("      , NULL "); 
                        //}
                        //else
                        //{
                        //    strSql.AppendLine("      , '" + sFirstMesTime + "'");
                        //}

                        //if (string.IsNullOrEmpty(sSecMesTime))
                        //{
                        //    strSql.AppendLine("      , NULL ");
                        //}
                        //else
                        //{
                        //    strSql.AppendLine("      , '" + sSecMesTime + "' ");
                        //}

                        //strSql.AppendLine("          , "+ sEmptyWeight + ", " + sTotWeight + ", " + sActualWeight + " ");
                        //if (string.IsNullOrEmpty(sGradeCd))
                        //{
                        //    strSql.AppendLine("           ,  NULL");
                        //}
                        //else
                        //{
                        //    strSql.AppendLine("           ,  " + sGradeCd + "");
                        //}
                        //strSql.AppendLine("           , '" + sGradeNm + "', '" + sCarNo + "', '상품'                 , " + sLossWeight + " ");
                        //strSql.AppendLine("           ,  0                     , 0                    , 0                   , " + sAcptWeight + ", 0 ");
                        //strSql.AppendLine("           ,  0                     , 0                    , 0                   , 0                  , " + sEntId + " ");
                        //strSql.AppendLine("           ,  0                     , 100                  , 1                   , 0                  , 1 ");
                        //strSql.AppendLine("           ,  0                     , '0001-01-01'         , ''                  , ''                 , '' ");
                        //strSql.AppendLine("           ,  0                     , 0                    , ''                  , ''                 , 0 ");
                        //strSql.AppendLine("           , ''                     , " + sActualWeight + ", 0                   , 0                  , 0 ");
                        //strSql.AppendLine("           ,  0                     , 0                    , 0                   , 0                  , '' ");
                        //strSql.AppendLine("           , ''                     , 0 ");
                        //if (string.IsNullOrEmpty(sInspector))
                        //{
                        //    strSql.AppendLine("                    , NULL   ");
                        //}
                        //else
                        //{
                        //    strSql.AppendLine("                    , " + sInspector + "   ");
                        //}
                        //strSql.AppendLine("                    , '" + sInspNote + "'");
                        //strSql.AppendLine("          , '' ) ");

                        #endregion[2021-01-14 쿼리정리 이전]
                    }
                    else if (sKeratype.Equals("출고"))
                    {
                        string sSecMesTime = "";  //총중량 시각
                        if(DateTime.TryParse(TimeEditSecond.EditValue?.ToString(), out DateTime secTime))
                        {
                            sSecMesTime = secTime.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        if (string.IsNullOrEmpty(sEmptyWeight) || sEmptyWeight.Equals("0"))
                        {
                            sSecMesTime = "";
                        }

                        string sFirstMesTime = ""; //공차중량 시각
                        if(DateTime.TryParse(TimeEditFirst.EditValue?.ToString(), out DateTime firstTime))
                        {
                            sFirstMesTime = firstTime.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        if (string.IsNullOrEmpty(sTotWeight) || sTotWeight.Equals("0"))
                        {
                            sFirstMesTime = "";
                        }

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" INSERT INTO MESURING ");
                        strSql.AppendLine("           ( KERATYPE       , J_ASSIGNID       , J_COMPANY        , SUN              , J_DATE , MAIPCHERID ");
                        strSql.AppendLine("           , FIRSTTIME      , SECONDTIME       , FIRSTWEIGHT      , SECONDWEIGHT     , WEIGHT ");
                        strSql.AppendLine("           , J_SERIAL       , GUBUN1           , J_BNUM           , K_NAME           , ICHAGAM ");
                        strSql.AppendLine("           , OCHAGAM        , IGAMGA           , OGAMGA           , IWEIGHT          , OWEIGHT ");
                        strSql.AppendLine("           , IDANGA         , ODANGA           , IKONGKEP         , OKONGKEP         , USERCODE ");
                        strSql.AppendLine("           , BUSEOCODE      , P_ID             , J_GARAGE         , J_ID             , KYERYANG12 ");
                        strSql.AppendLine("           , DRIVER_INOUT   , AGREE_DATE       , J_STATE2         , REGIONSTART      , REGIONDEST ");
                        strSql.AppendLine("           , TRANSPORTDANGA , TRANSPORTKUMAK   , TRANSPORTPERSON  , TRANSPORTCUSTOM  , TRANSPORTC_SERIAL ");
                        strSql.AppendLine("           , TRANSPORTBIGO  , CUSTOMWEIGHT     , LOSSWEIGHT       , TRANSPORTJUNGSAN , MAGAM_FLAG ");
                        strSql.AppendLine("           , IPCHULGO_MAIPID, IPCHULGO_MACHULID, WEIGHT_GUBUN     , LENGTHSID        , BLNO ");
                        strSql.AppendLine("           , CONTAINERNO    , DAMAGE           , GUMSU_SERIAL     , GUMSUBIGO        , J_CHECK");
                        //strSql.AppendLine("           , CHRG_ID ");
                        strSql.AppendLine("           ) ");
                        strSql.AppendLine("      VALUES ");
                        strSql.AppendLine("           ( ");
                        strSql.AppendLine("             '" + sKeratype + "',  " + sDealerCd + ",  '" + sDealerNM + "', " + sSeq + ", '" + sJDate + "', 0 ");
                        //if (string.IsNullOrEmpty(sSecMesTime))
                        //{
                        //    strSql.AppendLine("      , NULL ");
                        //}
                        //else
                        //{
                        //    strSql.AppendLine("      , '" + sSecMesTime + "' ");
                        //}

                        //if (string.IsNullOrEmpty(sFirstMesTime))
                        //{
                        //    strSql.AppendLine("      , NULL ");
                        //}
                        //else
                        //{
                        //    strSql.AppendLine("      , '" + sFirstMesTime + "' ");
                        //}
                        //strSql.AppendLine("          ,"+ sTotWeight + " , " + sEmptyWeight + " , " + sActualWeight + " ");

                        if (string.IsNullOrEmpty(sFirstMesTime))
                        {
                            strSql.AppendLine("      , '" + sSecMesTime + "' ");
                        }
                        else
                        {
                            strSql.AppendLine("      , '" + sFirstMesTime + "'");
                        }

                        if (string.IsNullOrEmpty(sSecMesTime))
                        {
                            strSql.AppendLine("      , '" + sFirstMesTime + "' ");
                        }
                        else
                        {
                            strSql.AppendLine("      , '" + sSecMesTime + "' ");
                        }

                        strSql.AppendLine("          , " + sEmptyWeight + ", " + sTotWeight + ", " + sActualWeight + " ");

                        if (string.IsNullOrEmpty(sGradeCd))
                        {
                            strSql.AppendLine("           ,  NULL");
                        }
                        else
                        {
                            strSql.AppendLine("           ,  " + sGradeCd + "");
                        }
                        strSql.AppendLine("           , '" + sGradeNm + "', '" + sCarNo + "', '상품'                , 0 ");
                        strSql.AppendLine("           ,  " + sLossWeight + ",   0                    , 0                   , 0   , " + sAcptWeight + " ");
                        strSql.AppendLine("           ,  0                     ,   0                    , 0                   , 0                  , " + sEntId + " ");
                        strSql.AppendLine("           ,  0                     ,   100                  , 1                   , 0                  , 1 ");
                        strSql.AppendLine("           ,  0                     ,   '0001-01-01'         , ''                  , ''                 , '' ");
                        strSql.AppendLine("           ,  0                     ,   0                    , ''                  , ''                 , 0 ");
                        strSql.AppendLine("           , ''                     ,  " + sActualWeight + ", 0                   , 0                  , 0 ");
                        strSql.AppendLine("           ,  0                     ,   0                    , 0                   , 0                  , '' ");
                        strSql.AppendLine("           , ''                     ,   0 ");
                        if (string.IsNullOrEmpty(sInspector))
                        {
                            strSql.AppendLine("                    , NULL   ");
                        }
                        else
                        {
                            strSql.AppendLine("                    , " + sInspector + "   ");
                        }
                        strSql.AppendLine("                    , '" + sInspNote + "'");
                        strSql.AppendLine("           , ''  ");
                        //strSql.AppendLine("           , (SELECT CHRG_ID FROM ACC_DEALER_CD WHERE DEALER_CD = '" + sDealerCd + "') )");
                        strSql.AppendLine("           ) ");
                    }
                }
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                //string sLOG = string.Format("[계근프로그램] 거래구분 : {0}" +
                //                        ", 순번 : {1}" +
                //                        ", 거래처명 : {2}" +
                //                        ", 공차중량 : {3}" +
                //                        ", 대지중량 : {4}" +
                //                        ", 인수량 : {5}"
                //                        , sKeratype
                //                        , sSeq
                //                        , sDealerNM
                //                        , sEmptyWeight
                //                        , sTotWeight
                //                        , sAcptWeight);

                //strSql.Clear();
                //strSql.AppendLine(" ");
                //strSql.AppendLine(" INSERT INTO ZSYS_LOG ");
                //strSql.AppendLine(" 	           ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK )   ");
                //strSql.AppendLine(" 	      VALUES ");
                //strSql.AppendLine(" 	           ( @OCCUR_DATE ");
                //strSql.AppendLine(" 	           , @USRCD ");
                //strSql.AppendLine(" 	           , ( SELECT IFNULL(MAX(X1.LOG_SEQ), 0) + 1  ");
                //strSql.AppendLine("                      FROM DAEJIERP.ZSYS_LOG X1 ");
                //strSql.AppendLine("                     WHERE X1.OCCUR_DATE = @OCCUR_DATE ");
                //strSql.AppendLine("                       AND X1.USRCD = @USRCD ) #LOG_SEQ(구분자) ");
                //strSql.AppendLine(" 	           , @PGM_ID ");
                //strSql.AppendLine(" 	           , @EDIT_KIND ");
                //strSql.AppendLine(" 	           , @_IP ");
                //strSql.AppendLine(" 	           , @EDIT_RMK ); ");

                //cmd.CommandType = CommandType.Text;
                //cmd.CommandText = strSql.ToString();
                //cmd.Parameters.Clear();
                //cmd.Parameters.AddWithValue("@OCCUR_DATE", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                //cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                //cmd.Parameters.AddWithValue("@PGM_ID", this.Name);
                //cmd.Parameters.AddWithValue("@EDIT_KIND", "S");
                //cmd.Parameters.AddWithValue("@_IP", GetLocalIPAddress());
                //cmd.Parameters.AddWithValue("@EDIT_RMK", sLOG);
                //cmd.ExecuteNonQuery();
                //cmd.Parameters.Clear();

                //if (sLogGb.Equals("1"))
                //{
                //    strSql.Clear();
                //    strSql.AppendLine(" UPDATE MESURING ");
                //    strSql.AppendLine("    SET EDIT_GB = @EDIT_GB ");
                //    strSql.AppendLine("      , EDIT_RMK = @EDIT_REMARK");
                //    strSql.AppendLine("  WHERE JUNPYOID = @JUNPYOID   ");

                //    cmd.CommandType = CommandType.Text;
                //    cmd.CommandText = strSql.ToString();
                //    cmd.Parameters.AddWithValue("@EDIT_GB", "B"); // 수정구분(트리거 위함) -> A: 마감 및 업로드, B: 계근프로그램, C: 단가입력
                //    cmd.Parameters.AddWithValue("@EDIT_REMARK", REMARK_RESULT);
                //    cmd.Parameters.AddWithValue("@JUNPYOID", sJunpyoId);
                //    cmd.ExecuteNonQuery();
                //}

                Cursor = Cursors.Default;
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");
                REMARK_RESULT = string.Empty;//초기화
                BtnRetr_Click(null, null);
                GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColSeq, sSeq);
                GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(GridViewRetr.FocusedRowHandle, GridColDate, sJDate);
                GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(GridViewRetr.FocusedRowHandle, GridColCarNo, sCarNo);
                return true;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
                REMARK_RESULT = string.Empty;//초기화
                return false;
            }
        }

        public string GetLocalIPAddress()
        {
            var host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (var ip in host.AddressList)
            {
                if (ip.AddressFamily == AddressFamily.InterNetwork)
                {
                    return ip.ToString();
                }
            }
            throw new Exception("No network adapters with an IPv4 address in the system!");
        }

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT null AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else
            {

            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT CAST(DEALER_CD AS VARCHAR) AS CD ");
                strSql.AppendLine("      , DEALER_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEALER_NM) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
                strSql.AppendLine("  WHERE A.EOB_YN = 'N' ");
                strSql.AppendLine("    AND (A.DEALER_GB = '매입' OR A.DEALER_GB = '매출' OR A.DEALER_GB = '입출') ");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT CAST(A.J_SERIAL AS VARCHAR) AS CD");
                strSql.AppendLine("      , A.GUBUN1 AS NM");
                strSql.AppendLine("      , A.J_SERIAL AS SEQ");
                strSql.AppendLine("   FROM JAJAE A");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.EMP_ID AS CD ");
                strSql.AppendLine("      , A.EMP_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY EMP_NM) AS SEQ ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A ");
                strSql.AppendLine("  WHERE A.EMPL_GB = 'Y' ");
            }

            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
            strSql.AppendLine("     ORDER BY SEQ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private void ClearFps()
        {
            GridViewRetr.FocusedRowChanged -= GridViewRetr_FocusedRowChanged;
            GridRetr.DataSource = null;
            GridViewRetr.FocusedRowChanged += GridViewRetr_FocusedRowChanged;
        }

        public void ClearAllForm(Control Ctrl)
        {
            if (Ctrl.HasChildren)
            {
                foreach (Control ctrl in Ctrl.Controls)
                {
                    if (ctrl is DevExpress.XtraEditors.TextEdit)
                        (ctrl as DevExpress.XtraEditors.TextEdit).ResetText();

                    if (ctrl is DevExpress.XtraEditors.LookUpEdit)
                        (ctrl as DevExpress.XtraEditors.LookUpEdit).EditValue = "";

                    if (ctrl is DevExpress.XtraEditors.DateEdit)
                        (ctrl as DevExpress.XtraEditors.DateEdit).EditValue = DateTime.Now.ToString("yyyy-MM-dd");

                    if (ctrl is DevExpress.XtraEditors.TimeEdit)
                        (ctrl as DevExpress.XtraEditors.TimeEdit).EditValue = DateTime.Now.ToString();

                    if (ctrl is DevExpress.XtraEditors.ComboBoxEdit)
                        (ctrl as DevExpress.XtraEditors.ComboBoxEdit).ResetText();

                    if (ctrl is DevExpress.XtraEditors.PictureEdit)
                        (ctrl as DevExpress.XtraEditors.PictureEdit).Image = null;

                    if (ctrl.HasChildren)
                        ClearAllForm(ctrl);//Recursive
                }
            }
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            SetData();
        }

        private void SetData()
        {
            //if (GridViewRetr.RowCount > 0) return;

            MakeReadOnly();

            Cursor = Cursors.WaitCursor;

            DateTime j_date;
            DateTime.TryParse(GridViewRetr.GetFocusedRowCellValue("J_DATE")?.ToString(), out j_date);
            DateEditMes.EditValue = j_date;
            TxtSeq.Text = GridViewRetr.GetFocusedRowCellValue("SUN")?.ToString();

            string sDealGb = GridViewRetr.GetFocusedRowCellValue("KERATYPE")?.ToString();
            if (sDealGb == "입고")
            {
                RdGbSelect.SelectedIndex = 0;
            }
            else if (sDealGb == "출고")
            {
                RdGbSelect.SelectedIndex = 1;
            }

            BtnNew.Text = "신규(F1)";
            BtnModify.Text = "수정(F2)";

            TxtCarNo.Text = GridViewRetr.GetFocusedRowCellValue("J_BNUM")?.ToString();
            string sDEALER_CD = GridViewRetr.GetFocusedRowCellValue("DEALER_CD")?.ToString();
            LkupEditDealerNM.EditValue = sDEALER_CD;
            string sJ_SERIAL = GridViewRetr.GetFocusedRowCellValue("J_SERIAL")?.ToString();
            LkupEditGrade.EditValue = sJ_SERIAL;
            LkupEditInspector.EditValue = GridViewRetr.GetFocusedRowCellValue("GUMSU_SERIAL")?.ToString();

            TxtInsptNote.Text = GridViewRetr.GetFocusedRowCellValue("INSPECTNOTE")?.ToString();

            TimeEditFirst.EditValue = GridViewRetr.GetFocusedRowCellValue("FIRSTTIME")?.ToString();
            TimeEditSecond.EditValue = GridViewRetr.GetFocusedRowCellValue("SECONDTIME")?.ToString();

            TxtEmptyWeight.Text = GridViewRetr.GetFocusedRowCellValue("FIRSTWEIGHT")?.ToString();
            TxtTotWeight.Text = GridViewRetr.GetFocusedRowCellValue("SECONDWEIGHT")?.ToString();

            TxtActualWeight.Text = GridViewRetr.GetFocusedRowCellValue("ACTUALWEIGHT")?.ToString();
            TxtLossWeight.Text = GridViewRetr.GetFocusedRowCellValue("LOSS")?.ToString();
            TxtLossCause.Text = GridViewRetr.GetFocusedRowCellValue("LOSSCAUSE")?.ToString();
            TxtAcceptWeight.Text = GridViewRetr.GetFocusedRowCellValue("ACCEPTWEIGHT")?.ToString();
            TxtImageNo.Text = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();

            string sJunpyoID = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();
            TxtJunpyoId.EditValue = sJunpyoID;


            //foreach(KeyValuePair<string, DataChanged> param in _dicDataChanged)
            //{
            //    _dicDataChanged[param.Key] = DataChanged.UnChanged;
            //}

            //PictureEdit Reset
            ResetPictureEdit(pictureEdit1, pictureEdit2, pictureEdit3, pictureEdit4, pictureEdit5, pictureEdit6);

            string[] sJDateArr = GridViewRetr.GetFocusedRowCellValue("J_DATE")?.ToString().Split(' ');
            if (GridViewRetr.FocusedRowHandle < 0)
            {
                pictureEdit1.Image = null;
                pictureEdit2.Image = null;
                pictureEdit3.Image = null;
                pictureEdit4.Image = null;
                pictureEdit5.Image = null;
                pictureEdit6.Image = null;
            }
            else
            {
                string sJDate = sJDateArr[0];
                string[] strArr = sJDate.Split('-');
                //string ftpPath = @"ftp://192.168.0.202/Images/" + strArr[0] + "/" + strArr[1] + "/" + sJDate;
                string ftpPath = @"ftp://"+ComnEtcFunc.FTP_IP+"/Images/" + strArr[0] + "/" + strArr[1] + "/" + sJDate;
                string user = ComnEtcFunc.FTP_USER;
                string pw = ComnEtcFunc.FTP_PW;

                FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(ftpPath);

                req1.Credentials = new NetworkCredential(user, pw);
                req1.Method = WebRequestMethods.Ftp.ListDirectory;

                //FTP 이미지를 Byte[]로 파싱하여 저장할 Dictionary 객체 초기세팅
                Dictionary<string, Image> dicPicture = new Dictionary<string, Image>();
                dicPicture.Add("1_1", null);
                dicPicture.Add("1_2", null);
                dicPicture.Add("1_3", null);
                dicPicture.Add("2_1", null);
                dicPicture.Add("2_2", null);
                dicPicture.Add("2_3", null);

                Dictionary<string, Image> dicCopy = new Dictionary<string, Image>();
                foreach (KeyValuePair<string, Image> item in dicPicture)
                    dicCopy.Add(item.Key, null);

                try
                {
                    string[] filesInDirectory = null;
                    using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                    {
                        StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                        string strData = reader1.ReadToEnd();
                        //폴더 내 파일이름
                        filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                        reader1.Close();

                        foreach (KeyValuePair<string, Image> item in dicPicture)
                        {
                            //해당 파일 Index
                            int findIndex = Array.FindIndex(filesInDirectory, i => i == string.Format("{0}_{1}.jpg", sJunpyoID, item.Key));
                            if (findIndex >= 0)
                            {
                                string fileName = filesInDirectory[findIndex];
                                dicCopy[item.Key] = DownloadFTPFile(string.Format(@"{0}\{1}", ftpPath, fileName), user, pw);
                            }
                        }
                    }

                    pictureEdit1.Image = dicCopy["1_1"];
                    pictureEdit2.Image = dicCopy["1_2"];
                    pictureEdit3.Image = dicCopy["1_3"];
                    pictureEdit4.Image = dicCopy["2_1"];
                    pictureEdit5.Image = dicCopy["2_2"];
                    pictureEdit6.Image = dicCopy["2_3"];

                    //dicPicture = null;
                    Cursor = Cursors.Default;
                }
                catch (Exception ex)
                {
                    Cursor = Cursors.Default;
                    pictureEdit1.Image = null;
                    pictureEdit2.Image = null;
                    pictureEdit3.Image = null;
                    pictureEdit4.Image = null;
                    pictureEdit5.Image = null;
                    pictureEdit6.Image = null;

                    return;
                }

                #region[이전 FTP 관련 코드]

                //try
                //{
                //    #region[TEMP]
                //    //using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                //    //{
                //    //    StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                //    //    string strData = reader1.ReadToEnd();

                //    //    string[] filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                //    //    reader1.Close();

                //    //    for (int i = 0; i < filesInDirectory.Length; i++)
                //    //    {
                //    //        fileName = filesInDirectory[i];
                //    //        fileNameArr = fileName.Split('_');

                //    //        if (fileNameArr[2].Equals(sJunpyoID))
                //    //        {
                //    //            if (fileNameArr[3].Equals("1"))
                //    //            {
                //    //                if (fileNameArr[4].Substring(0, 1).Equals("1"))
                //    //                {
                //    //                    Stream stream = GetFtpImg(ftpPath + "/" + fileName, user, pw);
                //    //                    dicPicture["1_1"] = Image.FromStream(stream);
                //    //                    //pictureEdit1.Image = Image.FromStream(stream);
                //    //                }
                //    //                else if (fileNameArr[4].Substring(0, 1).Equals("2"))
                //    //                {
                //    //                    Stream stream = GetFtpImg(ftpPath + "/" + fileName, user, pw);
                //    //                    //pictureEdit2.Image = Image.FromStream(stream);
                //    //                    dicPicture["1_2"] = Image.FromStream(stream);
                //    //                }
                //    //                else if (fileNameArr[4].Substring(0, 1).Equals("3"))
                //    //                {
                //    //                    Stream stream = GetFtpImg(ftpPath + "/" + fileName, user, pw);
                //    //                    dicPicture["1_3"] = Image.FromStream(stream);
                //    //                    //pictureEdit3.Image = Image.FromStream(stream);
                //    //                }
                //    //            }
                //    //            else if (fileNameArr[3].Equals("2"))
                //    //            {
                //    //                if (fileNameArr[4].Substring(0, 1).Equals("1"))
                //    //                {
                //    //                    Stream stream = GetFtpImg(ftpPath + "/" + fileName, user, pw);
                //    //                    //pictureEdit4.Image = Image.FromStream(stream);
                //    //                    dicPicture["2_1"] = Image.FromStream(stream);
                //    //                }
                //    //                else if (fileNameArr[4].Substring(0, 1).Equals("2"))
                //    //                {
                //    //                    Stream stream = GetFtpImg(ftpPath + "/" + fileName, user, pw);
                //    //                    //pictureEdit5.Image = Image.FromStream(stream);
                //    //                    dicPicture["2_2"] = Image.FromStream(stream);
                //    //                }
                //    //                else if (fileNameArr[4].Substring(0, 1).Equals("3"))
                //    //                {
                //    //                    Stream stream = GetFtpImg(ftpPath + "/" + fileName, user, pw);
                //    //                    //pictureEdit6.Image = Image.FromStream(stream);
                //    //                    dicPicture["2_3"] = Image.FromStream(stream);
                //    //                }
                //    //            }
                //    //        }
                //    //    }
                //    //}
                //    #endregion[]

                //    string[] filesInDirectory = null;
                //    using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                //    {
                //        StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                //        string strData = reader1.ReadToEnd();
                //        //폴더 내 파일이름
                //        filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                //        reader1.Close();

                //        string[] filesCopy = new string[filesInDirectory.Length];
                //        //파일이름은 전표_(1, 2)_(1, 2, 3) 순으로 다시 파싱
                //        for (int i = 0; i < filesInDirectory.Length; i++)
                //        {
                //            string[] sVal = filesInDirectory[i].Split('_');
                //            if (sVal.Length < 3)
                //            {
                //                filesCopy[i] = string.Empty;
                //            }
                //            else
                //            {
                //                string[] sVal2 = sVal[2].Split('.');
                //                filesCopy[i] = string.Format("{0}_{1}_{2}", sVal[0], sVal[1], sVal2[0]);
                //            }
                //        }

                //        string[] names = new string[dicPicture.Count];
                //        int[] iArrIdx = new int[names.Length];
                //        foreach (KeyValuePair<string, Image> item in dicPicture)
                //        {
                //            //해당 파일 Index
                //            int findIndex = Array.FindIndex(filesCopy, i => i == string.Format("{0}_{1}", sJunpyoID, item.Key));
                //            if (findIndex >= 0)
                //            {
                //                fileName = filesInDirectory[findIndex];
                //                dicCopy[item.Key] = DownloadFTPFile(string.Format(@"{0}\{1}", ftpPath, fileName), user, pw);
                //            }

                //            if (CheckImageNull(dicCopy))
                //                continue;
                //            else
                //                break;
                //        }
                //    }

                //    pictureEdit1.Image = dicCopy["1_1"];
                //    pictureEdit2.Image = dicCopy["1_2"];
                //    pictureEdit3.Image = dicCopy["1_3"];
                //    pictureEdit4.Image = dicCopy["2_1"];
                //    pictureEdit5.Image = dicCopy["2_2"];
                //    pictureEdit6.Image = dicCopy["2_3"];

                //    dicPicture = null;

                //}
                //catch (Exception ex)
                //{
                //    Cursor = Cursors.Default;
                //    pictureEdit1.Image = null;
                //    pictureEdit2.Image = null;
                //    pictureEdit3.Image = null;
                //    pictureEdit4.Image = null;
                //    pictureEdit5.Image = null;
                //    pictureEdit6.Image = null;

                //    return;
                //}

                //Cursor = Cursors.Default;

                #endregion[이전 FTP 관련 코드]
            }
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

        public Image DownloadFTPFile(string sourceFileURI, string user, string pw)
        {
            Image img = null;
            try
            {
                Uri sourceFileUri = new Uri(sourceFileURI);
                FtpWebRequest ftpWebRequest = WebRequest.Create(sourceFileUri) as FtpWebRequest;
                
                ftpWebRequest.Credentials = new NetworkCredential(user, pw);
                ftpWebRequest.Method = WebRequestMethods.Ftp.DownloadFile;

                FtpWebResponse ftpWebResponse = ftpWebRequest.GetResponse() as FtpWebResponse;

                Stream sourceStream = ftpWebResponse.GetResponseStream();
                img = Image.FromStream(sourceStream);
                sourceStream.Close();
            }
            catch (Exception ex)  
            {
                return null;
            }

            return img;
        }

        #endregion
        
        private bool CheckImageNull(Dictionary<string, Image> dicParams)
        {
            bool bYN = false;
            foreach (KeyValuePair<string, Image> item in dicParams)
            {
                if (item.Value == null)
                {
                    bYN = true;
                    break;
                }
            }
            return bYN;
        }

        private Stream GetFtpImg(string ftpPath, string user, string pw)
        {
            FtpWebRequest req = (FtpWebRequest)WebRequest.Create(ftpPath);

            req.Credentials = new NetworkCredential(user, pw);
            req.Method = WebRequestMethods.Ftp.DownloadFile;

            FtpWebResponse resp = (FtpWebResponse)req.GetResponse();

            Stream stream = resp.GetResponseStream();

            return stream;
        }

        private void BtnUpload_Click(object sender, EventArgs e)
        {
            string sJunpyoId = TxtJunpyoId.EditValue?.ToString();
            if (string.IsNullOrEmpty(sJunpyoId))
            {
                XtraMessageBox.Show("업로드하려는 데이터를 정확히 선택하세요.");
                return;
            }
            IN05001F02 frm = new IN05001F02();
            frm.JUNPYO_ID = sJunpyoId;
            if(frm.ShowDialog() == DialogResult.OK)
            {
                BtnRetr_Click(null, null);

                GridColumn col1 = GridViewRetr.Columns["JUNPYOID"];
                GridViewRetr.OptionsSelection.MultiSelect = true;
                GridViewRetr.ClearSelection();
                int rowHandle = -1;

                while (rowHandle != GridControl.InvalidRowHandle)
                {
                    rowHandle = GridViewRetr.LocateByDisplayText(rowHandle + 1, col1, sJunpyoId);
                    GridViewRetr.FocusedRowHandle = rowHandle;
                    break;
                }
            }
        }

        private void DateEditFrom_Leave(object sender, EventArgs e)
        {

        }

        private void DateEditTo_EditValueChanged(object sender, EventArgs e)
        {
            //if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            //{
            //    MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
            //    DateEditFrom.EditValue = DateEditTo.EditValue;
            //    return;
            //}
        }
        
        private void ResetPictureEdit(PictureEdit ptEdit1, PictureEdit ptEdit2, PictureEdit ptEdit3, PictureEdit ptEdit4, PictureEdit ptEdit5, PictureEdit ptEdit6)
        {
            ptEdit1.EditValue = null;
            ptEdit2.EditValue = null;
            ptEdit3.EditValue = null;
            ptEdit4.EditValue = null;
            ptEdit5.EditValue = null;
            ptEdit6.EditValue = null;
        }

        private void BtnTotWeight_Click(object sender, EventArgs e)
        {
            //TxtTotWeight.ResetText();
        }

        private void BtnEmptyWeight_Click(object sender, EventArgs e)
        {
            //TxtEmptyWeight.ResetText();
        }

        private void DateEditMes_Leave(object sender, EventArgs e)
        {
            TimeEditFirst.EditValue = DateEditMes.EditValue;
            TimeEditSecond.EditValue = DateEditMes.EditValue;
        }

        private void TxtTotWeight_Leave(object sender, EventArgs e)
        {
            string sTotWeight = TxtTotWeight.EditValue.ToString();
            string sEmptyWeight = TxtEmptyWeight.EditValue.ToString();

            if (!int.TryParse(sTotWeight, out int iOut))
            {
                MessageBox.Show("총중량은 숫자형식으로 입력하세요");
                TxtTotWeight.Focus();
            }
            else
            {
                SetActualWeight(sTotWeight, sEmptyWeight);
            }
        }

        private void TxtEmptyWeight_Leave(object sender, EventArgs e)
        {
            string sTotWeight = TxtTotWeight.EditValue.ToString();
            string sEmptyWeight = TxtEmptyWeight.EditValue.ToString();

            if (!int.TryParse(sEmptyWeight, out int iOut))
            {
                MessageBox.Show("공차중량은 숫자형식으로 입력하세요");
                TxtEmptyWeight.Focus();
            }
            else
            {
                SetActualWeight(sTotWeight, sEmptyWeight);
            }
        }

        private void SetActualWeight(string sTotWeight, string sEmptyWeight)
        {
            if (sTotWeight == "" || sEmptyWeight == "")
            {
                MessageBox.Show("총중량 및 공차중량을 확인하세요!");
                TxtEmptyWeight.Focus();
                return;
            }
            else
            {
                double dTotWeight = Convert.ToDouble(sTotWeight);
                double dEmptyWeight = Convert.ToDouble(sEmptyWeight);
                TxtActualWeight.EditValue = dTotWeight - dEmptyWeight;
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            //if (Reader != null) Reader.Close();
            //if (Writer != null) Writer.Close();
            //if (Server != null) Server.Stop();
            //if (Client != null) Client.Close();
            //Close();

            if (serialPort1.IsOpen)
            {
                serialPort1.Close();
                if (timer != null)
                {
                    timer.Close();
                }
                ListenThread = null;
            }

            Close();
        }

        private void MesMgtProgramDev_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnNew.PerformClick();
            }
            else if (e.KeyCode == Keys.F3)
            {
                if (BtnNew.Text.Equals("신규(F1)"))
                    return;

                BtnSave_Click();
            }
            else if(e.KeyCode == Keys.F4)
            {
                BtnDelete.PerformClick();
            }
            else if (e.KeyCode == Keys.F2)
            {
                BtnModify.PerformClick();
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr.PerformClick();
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnExcel.PerformClick();
            }
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        private readonly object _lockObject = new object();
        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
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
                    readData = serialPort1.ReadLine();
                    if (readData.Length == 17)
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
                            CSafeSetText(TxtDataRecv, weight.ToString());
                        }
                    }
                    
                }
                catch (Exception ex)
                {
                    
                }

                //Thread.Sleep(300);
            }
        }

        delegate void CrossThreadSafetySetText(Control ctl, String text);
        private void CSafeSetText(Control ctl, String text)
        {
            /*
             * InvokeRequired 속성 (Control.InvokeRequired, MSDN)
             *   짧게 말해서, 이 컨트롤이 만들어진 스레드와 현재의 스레드가 달라서
             *   컨트롤에서 스레드를 만들어야 하는지를 나타내는 속성입니다.  
             * 
             * InvokeRequired 속성의 값이 참이면, 컨트롤에서 스레드를 만들어 텍스트를 변경하고,
             * 그렇지 않은 경우에는 그냥 변경해도 아무 오류가 없기 때문에 텍스트를 변경합니다.
             */
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetText(CSafeSetText), ctl, text);
            else
                ctl.Text = text;
        }
        
        private void BtnSend_Click(object sender, EventArgs e)
        {
            this.Invoke(new EventHandler(TransComm));
        }

        private void TransComm(object s, EventArgs e)
        {

            string str = TxtDataRecv.Text?.ToString();
            this.BeginInvoke(new EventHandler(sportRCV));

            if (string.IsNullOrEmpty(str))
            {

                str = " ";
            }

            if (Writer != null)
            {

                try
                {
                    Writer.WriteLine(str); // 보내기
                    Writer.Flush();

                }
                catch (Exception)
                {
                    timer.Dispose();
                    timer.Close();
                    XtraMessageBox.Show("연결이 해제되었습니다.");

                }
            }
        }


        private void sportRCV(object sender, EventArgs e)
        {
            //if (serialPort1.IsOpen) { 
            //    if (serialPort1.BytesToRead > 7)
            //    {
            //        serialPort1.DiscardOutBuffer();
            //        //serialPort1.DiscardInBuffer();

            //        string str = serialPort1.ReadLine();
            //        //string str = "dataSand";
            //        string result = "0";
            //        if (str.Length > 12)
            //        {
            //            result = str.Substring(7,5);
            //        }

            //        TxtDataRecv.EditValue = result;
            //    }
            //}
            //else {
            //    Close();
            //
            if (serialPort1.IsOpen)
            {
                if (serialPort1.BytesToRead > 7)
                {
                    serialPort1.DiscardOutBuffer();
                    string str = serialPort1.ReadLine();
                    //string str = "dataSand@@@@@@"+i;

                    string result = "0";
                    if (str.Length > 12)
                    {
                        result = str.Substring(7, 7);
                    }

                    TxtDataRecv.Text = result;
                }
            }
            else
            {
                MessageBox.Show("포트가 닫혔습니다.");
            }
        }


        string icComportNum = string.Empty;
        Parity parity = Parity.Even;
        StopBits stopBits = StopBits.One;
        int iBaudRate = 0;
        int iDataBit = 0;
        string iTcpip = string.Empty;
        string iPortNum = string.Empty;

        private void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (serialPort1.IsOpen) { 
            BtnSend_Click(null, null);
            }
        }

        private void BtnSetting_Click(object sender, EventArgs e)
        {
            PopUpSerialTcpIp frm = new PopUpSerialTcpIp();

            if (frm.ShowDialog() == DialogResult.OK)
            {
                if (serialPort1.IsOpen)
                {
                    serialPort1.Close();
                }

                icComportNum = frm.cComportNum;
                iBaudRate = frm.ciBaudRate;
                iDataBit = frm.ciDataBit;
                parity = frm.cParity;
                stopBits = frm.cStopBits;

                iTcpip = frm.ciTcpip;
                iPortNum = frm.ciPortNum;

                serialPort1.PortName = icComportNum;
                serialPort1.BaudRate = iBaudRate;
                serialPort1.DataBits = iDataBit;
                serialPort1.Parity = parity;
                serialPort1.StopBits = stopBits;

                try
                {
                    serialPort1.Open();
                }

                catch(Exception ex)
                {
                    XtraMessageBox.Show(ex.ToString()); 
                    return;
                }
            }
            else
            {

            }

        }
        
        private void Listen() // 클라이언트와 연결하기
        {

                IPAddress addr = IPAddress.Parse(iTcpip); // 서버 ip

                int port = Convert.ToInt32(iPortNum); // 서버 포트
                Server = new TcpListener(addr, port);

                Server.Start(); // 서버 시작
            

                Client = Server.AcceptTcpClient(); // 클라이언트 연결 수락
                stream = Client.GetStream(); // 클라이언트 스트림 값 받아오기

            
                Reader = new StreamReader(stream);
                Writer = new StreamWriter(stream);


        }
        
        private void MesMgtProgramDev_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string FileName = string.Empty;
            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                string sFileNM = "계근리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    GridRetr.ExportToXls(FileName + ".xls");
                    Process.Start(FileName + ".xls");
                }
                fileDlg.Dispose();
            }
            catch (Exception ex)
            {
                if (ex.Message.Equals("Error Opening file"))
                {
                    //파일이 열려있음 체크
                    MessageBox.Show(((ex.InnerException).InnerException).Message);
                }
            }
        }

        private void CalculateWeight()
        {
            string sTotWeight = TxtTotWeight.EditValue?.ToString();
            string sEptWeight = TxtEmptyWeight.EditValue?.ToString();
            string sLossWeight = TxtLossWeight.EditValue?.ToString();

            double dTotWeight = string.IsNullOrEmpty(sTotWeight) ? 0 : Convert.ToDouble(sTotWeight);
            double dEptWeight = string.IsNullOrEmpty(sEptWeight) ? 0 : Convert.ToDouble(sEptWeight);
            double dLossWeight = string.IsNullOrEmpty(sLossWeight) ? 0 : Convert.ToDouble(sLossWeight);

            TxtActualWeight.EditValue = dTotWeight - dEptWeight;

            string sAtlWeight = TxtActualWeight.EditValue?.ToString();
            double dAtlWeight = string.IsNullOrEmpty(sAtlWeight) ? 0 : Convert.ToDouble(sAtlWeight);

            TxtAcceptWeight.EditValue = dAtlWeight - dLossWeight;

            if (dTotWeight == 0 || dEptWeight == 0)
            {
                TxtActualWeight.EditValue = 0;
                TxtAcceptWeight.EditValue = 0;
            }
            
        }

        private void TxtTotWeight_Leave_1(object sender, EventArgs e)
        {
            CalculateWeight();
        }

        private void TxtEmptyWeight_Leave_1(object sender, EventArgs e)
        {
            CalculateWeight();
        }

        private void TxtActualWeight_Leave(object sender, EventArgs e)
        {
            CalculateWeight();
        }

        private void TxtLossWeight_Leave(object sender, EventArgs e)
        {
            CalculateWeight();
        }

        private void MesMgtProgramDev_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.Close();
                if (timer != null)
                {
                    timer.Close();
                }
                //ListenThread = null;
            }

            if (timer != null)
            {
                timer.Close();
            }
            
        }

        private void BtnModify_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(TimeEditFirst.EditValue.ToString()))
                TimeEditFirst.EditValue = DateTime.Now;
            if (string.IsNullOrEmpty(TimeEditSecond.EditValue.ToString()))
                TimeEditFirst.EditValue = DateTime.Now;

            string sJunpyoId = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();
            string sCompany = GridViewRetr.GetFocusedRowCellValue("DEALER")?.ToString();
            if (string.IsNullOrEmpty(sJunpyoId))
            {
                XtraMessageBox.Show("JUNPYOID가 존재하지 않습니다.\r\n올바른 데이터를 선택하세요.");
                return;
            }
            else if (CheckInlistInfo(sJunpyoId,sCompany))
            {
                string sKeratype = GridViewRetr.GetFocusedRowCellValue("KERATYPE")?.ToString();
                string sDealerNm = GridViewRetr.GetFocusedRowCellValue("DEALER")?.ToString(); 
                string sGrade = GridViewRetr.GetFocusedRowCellValue("GUBUN1")?.ToString();
                string sAcptWeight = GridViewRetr.GetFocusedRowCellValue("ACCEPTWEIGHT")?.ToString();
                XtraMessageBox.Show( "거래구분 : " + sKeratype
                                   + "\r\n거래처 : " + sDealerNm
                                   + "\r\n등급 : " + sGrade
                                   + "\r\n인수량 : " + sAcptWeight
                                   + "\r\n해당 데이터는 마감된 자료이므로 수정할 수 없습니다.");
                return;
            }
            SetButtonText();
        }

        private void SetButtonText()
        {
            if (BtnModify.Text.Equals("수정(F2)"))
            {
                MakeUnReadOnly();
                BtnNew.Text = "저장(F3)";
                BtnModify.Text = "취소";
            }
            else if (BtnModify.Text.Equals("취소"))
            {
                MakeReadOnly();
                BtnNew.Text = "신규(F1)";
                BtnModify.Text = "수정(F2)";

                SetData();
            }
        }

        private bool CheckInlistInfo(string sJunpyoId, string sCompany)
        {
            if (sCompany != "재고이동")
            {
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT CASE WHEN COUNT(1) > 0 THEN 'Y' ELSE 'N' END AS YN  ");
                strSql.AppendLine("   FROM INLIST ");
                strSql.AppendLine("  WHERE J_RID = " + sJunpyoId + " ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                if (dt.Rows.Count > 0)
                {
                    string sYn = dt.Rows[0]["YN"].ToString();
                    if (sYn.Equals("Y"))
                        return true;
                    else
                        return false;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT CASE WHEN COUNT(1) > 0 THEN 'Y' ELSE 'N' END AS YN   ");
                strSql.AppendLine("   FROM Mesuring                                             ");
                strSql.AppendLine("  WHERE JunpyoID = " + sJunpyoId + "                         ");
                strSql.AppendLine("    AND J_Check = '1'                                        ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                if (dt.Rows.Count > 0)
                {
                    string sYn = dt.Rows[0]["YN"].ToString();
                    if (sYn.Equals("Y"))
                        return true;
                    else
                        return false;
                }
                else
                {
                    return false;
                }
            }
            
        }

        private void MakeUnReadOnly()
        {
            DateEditMes.ReadOnly = false;
            RdGbSelect.ReadOnly = false;
            TxtSeq.ReadOnly = false;
            TxtCarNo.ReadOnly = false;
            LkupEditDealerNM.ReadOnly = false;
            LkupEditGrade.ReadOnly = false;
            LkupEditInspector.ReadOnly = false;
            TxtInsptNote.ReadOnly = false;
            TxtTotWeight.ReadOnly = false;
            TimeEditSecond.ReadOnly = false;
            TxtEmptyWeight.ReadOnly = false;
            TimeEditFirst.ReadOnly = false;
            //BtnTotWeight.Enabled = true;
            //BtnEmptyWeight.Enabled = true;
            TxtActualWeight.ReadOnly = false;
            TxtLossWeight.ReadOnly = false;
            TxtLossCause.ReadOnly = false;
            TxtAcceptWeight.ReadOnly = false;
        }

        private void MakeReadOnly()
        {
            DateEditMes.ReadOnly = true;
            RdGbSelect.ReadOnly = true;
            TxtSeq.ReadOnly = true;
            TxtCarNo.ReadOnly = true;
            LkupEditDealerNM.ReadOnly = true;
            LkupEditGrade.ReadOnly = true;
            LkupEditInspector.ReadOnly = true;
            TxtInsptNote.ReadOnly = true;
            TxtTotWeight.ReadOnly = true;
            TimeEditSecond.ReadOnly = true;
            TxtEmptyWeight.ReadOnly = true;
            TimeEditFirst.ReadOnly = true;
            //BtnTotWeight.Enabled = false;
            //BtnEmptyWeight.Enabled = false;
            TxtActualWeight.ReadOnly = true;
            TxtLossWeight.ReadOnly = true;
            TxtLossCause.ReadOnly = true;
            TxtAcceptWeight.ReadOnly = true;
        }

        private void UpAndDownKey_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
                e.SuppressKeyPress = true;
        }

        #region [정렬기능(2020-06-02 정은영)]
        private void GridViewColumnSort_MouseUp(object sender, MouseEventArgs e)
        {
            //GridView view = (GridView)sender;
            //GridHitInfo hitInfo = view.CalcHitInfo(e.Location);

            //if (hitInfo.InColumn)
            //{
            //    if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.None)
            //    {
            //        hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.Descending;
            //        view.FocusedRowHandle = 0;
            //    }
            //    else if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.Descending)
            //    {
            //        hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
            //        view.FocusedRowHandle = 0;
            //    }
            //    else if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.Ascending)
            //    {
            //        hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.None;
            //        view.FocusedRowHandle = 0;
            //    }
            //    // if ((ModifierKeys & Keys.Control) == Keys.Control) return;
            //    //if ((ModifierKeys & Keys.Shift) != Keys.Shift) view.ClearSorting();
            //}
        }
        #endregion

        private void TxtJunpyoId_EditValueChanged(object sender, EventArgs e)
        {
            string sJunpyoId = TxtJunpyoId.EditValue?.ToString();
            if (string.IsNullOrEmpty(sJunpyoId))
                BtnUpload.Enabled = false;
            else
                BtnUpload.Enabled = true;
        }

        private void MesMgtProgramDev_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                string path = ComnEtcFunc.GetLayoutPath();
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }

                layoutControl1.SaveLayoutToXml(path + @"\" + this.Name + "_Layout.xaml");

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        public string _IP;
        TcpClient clientSocket = new TcpClient();
        NetworkStream stream1 = default(NetworkStream);
        string message = string.Empty;
        Thread t_handler;
        private void BtnConn_Click(object sender, EventArgs e)
        {
            PD03001F01 frm = new PD03001F01();
            frm._ParentForm = this;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (clientSocket == null)
                        clientSocket = new TcpClient();

                    if (clientSocket.Connected)
                    {
                        MessageBox.Show("연결 중입니다.");
                        return;
                    }
                    Cursor = Cursors.WaitCursor;
                    clientSocket = new TcpClient();

                    clientSocket.Connect(_IP, 9999); // 접속 IP 및 포트
                    stream1 = clientSocket.GetStream();
                    Cursor = Cursors.Default;
                }
                catch (Exception e2)
                {
                    Cursor = Cursors.Default;
                    MessageBox.Show("서버가 실행중이 아닙니다.", "연결 실패!");
                    MessageBox.Show(e2.ToString());
                    return;
                    //Application.Exit();
                }

                LblStatus.Text = "무인계근 서버에 연결 되었습니다.";
                LayoutBtnDisConn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                byte[] buffer = Encoding.Unicode.GetBytes(FmMainToolBar2.UserID + "$");
                stream1.Write(buffer, 0, buffer.Length);
                stream1.Flush();

                if (t_handler == null || !t_handler.IsAlive)
                {
                    t_handler = new Thread(GetMessage);
                    t_handler.IsBackground = true;
                    t_handler.Start();
                }
                else
                {
                    MessageBox.Show("이미 연결진행 중...");
                }
            }
        }

        private void GetMessage() // 메세지 받기
        {
            while (true)
            {
                try
                {
                    //clientSocket.
                    if (!clientSocket.Connected)
                    {
                        XtraMessageBox.Show("쓰레드 안 함수 : 연결 끊김");
                        StreanDisConn();
                        t_handler.Abort();
                        break;
                    }

                    stream1 = clientSocket.GetStream();
                    int BUFFERSIZE = clientSocket.ReceiveBufferSize;
                    byte[] buffer = new byte[500];
                    //byte[] buffer = new byte[stream1.rea];
                    int bytes = 0;
                    if (stream1.CanRead)
                    {
                        bytes = stream1.Read(buffer, 0, buffer.Length);
                        //stream1.ReadAsync(buffer, 0, buffer.Length);
                    }
                    //stream1.Close();

                    //string message = Encoding.Unicode.GetString(buffer, 0, bytes);
                    //int iVal = BitConverter.ToInt32(buffer, 0);
                    string sVal = Encoding.ASCII.GetString(buffer, 0, bytes);
                    int iWeight = 0;
                    int.TryParse(sVal.Split('b')[0]?.ToString(), out iWeight);

                    //double dWeight = string.IsNullOrEmpty(sVal) ? 0 : Convert.ToDouble(sVal);
                    DisplayText(iWeight.ToString());

                    Thread.Sleep(50);
                }
                catch (Exception ex)
                {
                    DisplayText("0");
                }
            }
        }

        private void DisplayText(string text) // Server에 메세지 출력
        {
            if (TxtDataRecv.InvokeRequired)
            {
                TxtDataRecv.BeginInvoke(new MethodInvoker(delegate
                {
                    TxtDataRecv.EditValue = text;
                    TxtDataRecv.Properties.Appearance.BackColor = Color.RoyalBlue;
                }));
            }
            else
            {
                TxtDataRecv.EditValue = text;
                TxtDataRecv.Properties.Appearance.BackColor = Color.RoyalBlue;
            }
        }

        private void StreanDisConn()
        {
            if (stream1.CanWrite)
            {
                byte[] buffer = Encoding.Unicode.GetBytes("leaveChat" + "$");
                stream1.Write(buffer, 0, buffer.Length);
                stream1.Flush();
            }

            if (t_handler == null)
                return;

            if (t_handler.IsAlive)
                t_handler.Abort();

            if (clientSocket != null)
            {
                if (clientSocket.Connected)
                    clientSocket.Close();
            }
        }

        private void BtnDisConn_Click(object sender, EventArgs e)
        {
            StreanDisConn();
            LblStatus.Text = "";
            LayoutBtnDisConn.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            TxtDataRecv.Properties.Appearance.BackColor = Color.SlateGray;
        }

        TcpClient clientSocket2 = new TcpClient();
        NetworkStream stream2 = default(NetworkStream);
        string message2 = string.Empty;
        Thread t_handler2;
        private void BtnImage_Click(object sender, EventArgs e)
        {
            //if (GridViewRetr.IsGroupRow(GridViewRetr.FocusedRowHandle))
            //{
            //    XtraMessageBox.Show("이미지가 없습니다.\r\n원하시는 행을 선택하신 후 클릭하세요.");
            //    return;
            //}

            //Cursor = Cursors.WaitCursor;

            //string sJunpyoID = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();
            //if (string.IsNullOrEmpty(sJunpyoID))
            //{
            //    XtraMessageBox.Show("올바른 데이터를 선택하세요.");
            //    return;
            //}

            //AccFieldPSRetrImage.junpyoID = sJunpyoID;
            //AccFieldPSRetrImage fm = new AccFieldPSRetrImage();
            //fm.ShowDialog();

            /*
             * 2020-10-28 (수)
             * 무인계근시스템과 연동하여 무인계근 캡쳐 값 가져오는 로직으로 변경
             */

            if (string.IsNullOrEmpty(_IP))
            {
                PD03001F01 frm = new PD03001F01();
                frm._ParentForm = this;
                if (frm.ShowDialog() == DialogResult.OK)
                {
                   
                }
            }
            
            try
            {
                #region[FTP 관련 테스트 TEMP]
                //string user = ComnEtcFunc.FTP_USER;
                //string pw = ComnEtcFunc.FTP_PW;
                //string fileName = null;

                //FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(ftpPath);

                //req1.Credentials = new NetworkCredential(user, pw);
                //req1.Method = WebRequestMethods.Ftp.ListDirectory;



                //string sInitDir = string.Format(@"ftp://{0}/Img_Req/{1}/", ComnEtcFunc.FTP_IP, FmMainToolBar2.UserID);
                //string sInitDir = @"ftp://192.168.0.202/Img_Req/" + FmMainToolBar2.UserID + "\\";

                //FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(sInitDir);
                //string user = ComnEtcFunc.FTP_USER;
                //string pw = ComnEtcFunc.FTP_PW;
                //req1.Credentials = new NetworkCredential(user, pw);
                //req1.Method = WebRequestMethods.Ftp.ListDirectory;

                //string[] filesInDirectory = null;
                //Dictionary<string, Image> dicImages = new Dictionary<string, Image>();
                //using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                //{
                //    StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                //    string strData = reader1.ReadToEnd();
                //    //폴더 내 파일이름
                //    filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                //    reader1.Close();

                //    foreach (string filePath in filesInDirectory)
                //    {
                //        string[] filesCopy = filePath.Split('\\');
                //        dicImages.Add(filesCopy[filesCopy.Length - 1], DownloadFTPFile(string.Format(@"{0}\{1}", sInitDir, filePath), user, pw));
                //    }
                //}

                #endregion[FTP 관련 테스트 TEMP]

                if (clientSocket2 == null)
                    clientSocket2 = new TcpClient();

                //if (clientSocket2.Connected)
                //{
                //    MessageBox.Show("연결 중입니다.");
                //    return;
                //}
                Cursor = Cursors.WaitCursor;
                clientSocket2 = new TcpClient();

                clientSocket2.Connect(_IP, 8888); // 접속 IP 및 포트
                stream2 = clientSocket2.GetStream();
                Cursor = Cursors.Default;
            }
            catch (Exception e2)
            {
                Cursor = Cursors.Default;
                MessageBox.Show("서버가 실행중이 아닙니다.", "연결 실패!");
                MessageBox.Show(e2.ToString());
                return;
                //Application.Exit();
            }

            byte[] buffer = Encoding.Unicode.GetBytes(FmMainToolBar2.UserID + "$");
            stream2.Write(buffer, 0, buffer.Length);
            stream2.Flush();

            while (true)
            {
                try
                {
                    Cursor = Cursors.WaitCursor;
                    if (!clientSocket2.Connected)
                    {
                        clientSocket2.Close();
                        break;
                    }

                    stream2 = clientSocket2.GetStream();
                    int BUFFERSIZE = clientSocket2.ReceiveBufferSize;
                    buffer = new byte[BUFFERSIZE];
                    int bytes = stream2.Read(buffer, 0, buffer.Length);

                    string message = Encoding.Unicode.GetString(buffer, 0, bytes);
                    message = string.IsNullOrEmpty(message) ? string.Empty : message;
                    if (message.Equals("COMPLETE"))
                    {
                        Cursor = Cursors.Default;
                        Dictionary<string, Image> img = GetImagesFromFTP();
                        DeleteAllFiles(); //테스트 필요
                        IN05001F03 frm = new IN05001F03();
                        frm._RESULT = img;
                        frm.ShowDialog();
                        break;
                    }
                    else if (message.Length > 2)
                    {
                        Cursor = Cursors.Default;
                        XtraMessageBox.Show(message + "\r\n무인계근시스템에서 이미지캡처부분 에러발생하였습니다.");
                        break;
                    }
                }
                catch (Exception ex)
                {
                    //clientSocket2.Close();
                    //DisplayText2("0");
                }
            }

            //if (t_handler2 == null || !t_handler2.IsAlive)
            //{
            //    t_handler2 = new Thread(GetMessage2);
            //    t_handler2.IsBackground = true;
            //    t_handler2.Start();
            //}
            //else
            //{
            //    t_handler2.Abort();
            //    t_handler2 = new Thread(GetMessage2);
            //    t_handler2.IsBackground = true;
            //    t_handler2.Start();
            //    //MessageBox.Show("이미 연결진행 중...");
            //}


            Cursor = Cursors.Default;

        }

        private void GetMessage2() // 메세지 받기
        {
            while (true)
            {
                try
                {
                    if (!clientSocket2.Connected)
                    {
                        StreanDisConn2();
                        break;
                    }

                    stream2 = clientSocket2.GetStream();
                    int BUFFERSIZE = clientSocket2.ReceiveBufferSize;
                    byte[] buffer = new byte[BUFFERSIZE];
                    int bytes = stream2.Read(buffer, 0, buffer.Length);

                    string message = Encoding.Unicode.GetString(buffer, 0, bytes);
                    message = string.IsNullOrEmpty(message) ? string.Empty : message;
                    if (message.Equals("COMPLETE"))
                    {
                        Dictionary<string, Image> img = GetImagesFromFTP();
                        IN05001F03 frm = new IN05001F03();
                        frm._RESULT = img;
                        frm.ShowDialog();
                        t_handler2.Abort();
                    }
                    else if(message.Length > 2)
                    {
                        XtraMessageBox.Show(message + "\r\n무인계근시스템에서 이미지캡처부분 에러발생하였습니다.");
                        t_handler2.Abort();
                    }
                }
                catch (Exception ex)
                {
                    //clientSocket2.Close();
                    //DisplayText2("0");
                }
            }
        }

        private void DisplayText2(string text) // Server에 메세지 출력
        {
            if (TxtDataRecv.InvokeRequired)
            {
                TxtDataRecv.BeginInvoke(new MethodInvoker(delegate
                {
                    TxtDataRecv.EditValue = text;
                    TxtDataRecv.Properties.Appearance.BackColor = Color.RoyalBlue;
                }));
            }
            else
            {
                TxtDataRecv.EditValue = text;
                TxtDataRecv.Properties.Appearance.BackColor = Color.RoyalBlue;
            }
        }

        private void StreanDisConn2()
        {
            if (stream2.CanWrite)
            {

            }

            if (t_handler2 == null)
                return;

            if (t_handler2.IsAlive)
                t_handler2.Abort();

            if (clientSocket2 != null)
            {
                if (clientSocket2.Connected)
                    clientSocket2.Close();
            }
        }

        private Dictionary<string, Image> GetImagesFromFTP()
        {
            try
            {
                //string sInitDir = @"ftp://192.168.0.202/Img_Req/" + FmMainToolBar2.UserID + "\\";
                string sInitDir = string.Format(@"ftp://{0}/Img_Req/{1}/", ComnEtcFunc.FTP_IP, FmMainToolBar2.UserID);

                FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(sInitDir);
                string user = ComnEtcFunc.FTP_USER;
                string pw = ComnEtcFunc.FTP_PW;
                req1.Credentials = new NetworkCredential(user, pw);
                req1.Method = WebRequestMethods.Ftp.ListDirectory;

                string[] filesInDirectory = null;
                Dictionary<string, Image> dicImages = new Dictionary<string, Image>();
                using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                {
                    StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                    string strData = reader1.ReadToEnd();
                    //폴더 내 파일이름
                    filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    reader1.Close();

                    foreach (string filePath in filesInDirectory)
                    {
                        string[] filesCopy = filePath.Split('\\');
                        dicImages.Add(filesCopy[filesCopy.Length - 1], DownloadFTPFile(string.Format(@"{0}\{1}", sInitDir, filePath), user, pw));
                    }
                }

                return dicImages;
                
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private void DeleteAllFiles()
        {
            try
            {
                //string sInitDir = @"ftp://192.168.0.202/Img_Req/" + FmMainToolBar2.UserID + "/";
                string sInitDir = string.Format(@"ftp://{0}/Img_Req/{1}/", ComnEtcFunc.FTP_IP, FmMainToolBar2.UserID);

                FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(sInitDir);
                string user = ComnEtcFunc.FTP_USER;
                string pw = ComnEtcFunc.FTP_PW;
                req1.Credentials = new NetworkCredential(user, pw);
                req1.Method = WebRequestMethods.Ftp.ListDirectory;

                string[] filesInDirectory = null;
                Dictionary<string, Image> dicImages = new Dictionary<string, Image>();
                using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                {
                    StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                    string strData = reader1.ReadToEnd();
                    //폴더 내 파일이름
                    filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    reader1.Close();

                    foreach (string filePath in filesInDirectory)
                    {
                        DeleteFTPFile(string.Format(@"{0}\{1}", sInitDir, filePath), ComnEtcFunc.FTP_USER, ComnEtcFunc.FTP_PW);
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        public byte[] ImageToByteArray(System.Drawing.Image imageIn)
        {
            using (var ms = new MemoryStream())
            {
                imageIn.Save(ms, imageIn.RawFormat);
                return ms.ToArray();
            }
        }

        #region FTP 파일 삭제하기 - DeleteFTPFile(targetURI, userID, password)

        /// <summary>
        /// FTP 파일 삭제하기
        /// </summary>
        /// <param name="userID">사용자 ID</param>
        /// <param name="password">패스워드</param>
        /// <param name="targetURI">타겟 URI</param>
        /// <returns>처리 결과</returns>
        public bool DeleteFTPFile(string targetURI, string userID, string password)
        {
            try
            {
                FtpWebRequest ftpWebRequest = WebRequest.Create(targetURI) as FtpWebRequest;

                ftpWebRequest.Credentials = new NetworkCredential(userID, password);
                ftpWebRequest.Method = WebRequestMethods.Ftp.DeleteFile;

                FtpWebResponse ftpWebResponse = ftpWebRequest.GetResponse() as FtpWebResponse;
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }

        #endregion

        /*
         * 2020-10-25 현업요청
         * 계근 데이터 삭제로직 추가
         */
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sMSG = string.Format("계량일 : {0}" +
                    "\r\n차량번호 : {1}" +
                    "\r\n순번 : {2}" +
                    "\r\n거래처 : {3}" +
                    "\r\n품명 : {4}" +
                    "\r\n해당 계근 건에 대하여 삭제를 진행하시겠습니까?"
                    , DateEditMes.EditValue?.ToString().Substring(0, 10)
                    , TxtCarNo.EditValue?.ToString().Trim()
                    , TxtSeq.EditValue?.ToString()
                    , LkupEditDealerNM.Text
                    , LkupEditGrade.Text);

            if (XtraMessageBox.Show(sMSG, "계근데이터 삭제건", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }


            //JUNPYOID ValueCheck
            string sJunpyoId = TxtJunpyoId.EditValue?.ToString().Trim();
            if (string.IsNullOrEmpty(sJunpyoId))
                return;

            StringBuilder strSql = new StringBuilder();

            /*
             * 마감여부 조회 KERATYPE(거래구분)의 값에 따라 마감컬럼(IPCHULGO_~) 다르게 SELECT하여 0인 경우 미마감('N')으로 조회
             */
            strSql.Clear();
            strSql.AppendLine("");
            strSql.AppendLine(" SELECT CASE WHEN X1.CLOSE_GB = 0 OR X1.CLOSE_GB IS NULL THEN 'N' ELSE 'Y' END AS CLOSE_GB ");
            strSql.AppendLine("   FROM ( ");
            strSql.AppendLine("          SELECT CASE WHEN KERATYPE = '입고' THEN IPCHULGO_MAIPID ELSE IPCHULGO_MACHULID END AS CLOSE_GB ");
            strSql.AppendLine("            FROM MESURING ");
            strSql.AppendLine("           WHERE JUNPYOID = @JUNPYOID ");
            strSql.AppendLine("        ) X1 ");

            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("JUNPYOID", sJunpyoId);

            DataTable dt = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
            string sYn = string.Empty;
            if(dt.Rows.Count > 0)
            {
                sYn = dt.Rows[0]["CLOSE_GB"]?.ToString();
            }

            sYn = string.IsNullOrEmpty(sYn) ? string.Empty : sYn;

            //Y일 경우 마감자료이므로 리턴
            if (sYn.Equals("Y"))
            {
                sMSG = string.Format("계량일 : {0}" +
                    "\r\n차량번호 : {1}" +
                    "\r\n순번 : {2}" +
                    "\r\n거래처 : {3}" +
                    "\r\n품명 : {4}" +
                    "\r\n해당 계근 건은 마감처리가 되어있으므로 삭제가 불가합니다."
                    , DateEditMes.EditValue?.ToString().Substring(0, 10)
                    , TxtCarNo.EditValue?.ToString().Trim()
                    , TxtSeq.EditValue?.ToString()
                    , LkupEditDealerNM.Text
                    , LkupEditGrade.Text);

                XtraMessageBox.Show(sMSG);
                return;
            }

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                /*
                 * 2021-03-16
                 * Reference Key : #0006
                 * 로그 추가
                 */
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT CONVERT(DATE,A.J_DATE) AS J_DATE");
                strSql.AppendLine("      , A.KERATYPE ");
                strSql.AppendLine("      , A.SUN ");
                strSql.AppendLine("      , A.J_BNUM ");
                strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN MAIPCHERID ELSE J_ASSIGNID END AS COMPANY_ID ");
                strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN MAIPCHER ELSE J_COMPANY END AS COMPANY ");
                strSql.AppendLine("      , A.J_SERIAL ");
                strSql.AppendLine("      , A.GUBUN1 ");
                strSql.AppendLine("      , A.GUMSU_SERIAL ");
                strSql.AppendLine("      , B.EMP_NM ");
                strSql.AppendLine("      , A.GUMSUBIGO ");
                strSql.AppendLine("      , A.FIRSTWEIGHT");
                strSql.AppendLine("      , A.SECONDWEIGHT ");
                strSql.AppendLine("      , A.FIRSTTIME");
                strSql.AppendLine("      , A.SECONDTIME ");
                strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN IWEIGHT ELSE OWEIGHT END AS WEIGHT ");
                strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN ICHAGAM ELSE OCHAGAM END AS CHAGAM ");
                strSql.AppendLine("      , A.J_STATE ");
                strSql.AppendLine("      , A.KYERYANG12 ");
                strSql.AppendLine("   FROM MESURING A ");
                strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS B ");
                strSql.AppendLine("     ON A.GUMSU_SERIAL = B.EMP_ID ");
                strSql.AppendLine("  WHERE A.JUNPYOID = " + sJunpyoId + " ");

                DataTable dtLog = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                string sLogMsg = string.Empty;
                string sSTD_COLS = string.Empty;
                string sREF_RMK = string.Empty;
                int iLogCnt = 0;
                if (dtLog.Rows.Count > 0)
                {
                    iLogCnt++;
                    string sPRV_J_DATE = dtLog.Rows[0]["J_DATE"]?.ToString().Substring(0, 10);
                    string sPRV_KERATYPE = dtLog.Rows[0]["KERATYPE"]?.ToString();
                    string sPRV_SUN = dtLog.Rows[0]["SUN"]?.ToString();
                    string sPRV_J_BNUM = dtLog.Rows[0]["J_BNUM"]?.ToString();
                    string sPRV_COMPANY_ID = dtLog.Rows[0]["COMPANY_ID"]?.ToString();
                    string sPRV_COMPANY = dtLog.Rows[0]["COMPANY"]?.ToString();
                    string sPRV_J_SERIAL = dtLog.Rows[0]["J_SERIAL"]?.ToString();
                    string sPRV_GUBUN1 = dtLog.Rows[0]["GUBUN1"]?.ToString();
                    string sPRV_GUMSU_SERIAL = dtLog.Rows[0]["GUMSU_SERIAL"]?.ToString();
                    string sPRV_EMPNM = dtLog.Rows[0]["EMP_NM"]?.ToString();
                    string sPRV_GUMSUBIGO = dtLog.Rows[0]["GUMSUBIGO"]?.ToString();

                    /*
                     * 2021-03-26
                     * Reference Key : #0006
                     */
                    string sPRV_J_STATE = dtLog.Rows[0]["J_STATE"]?.ToString();
                    string sPRV_KYERYANG12 = dtLog.Rows[0]["KYERYANG12"]?.ToString();

                    int iPRV_CHAGAM = 0;
                    int.TryParse(dtLog.Rows[0]["CHAGAM"]?.ToString(), out iPRV_CHAGAM);

                    int iPRV_FIRSTWEIGHT = 0;
                    int.TryParse(dtLog.Rows[0]["FIRSTWEIGHT"]?.ToString(), out iPRV_FIRSTWEIGHT);

                    int iPRV_SECONDWEIGHT = 0;
                    int.TryParse(dtLog.Rows[0]["SECONDWEIGHT"]?.ToString(), out iPRV_SECONDWEIGHT);

                    DateTime dtFirstTime;
                    DateTime.TryParse(dtLog.Rows[0]["FIRSTTIME"]?.ToString(), out dtFirstTime);

                    DateTime dtSecondTime;
                    DateTime.TryParse(dtLog.Rows[0]["SECONDTIME"]?.ToString(), out dtSecondTime);

                    /*
                     * #0006
                     */
                    string sEMPTYWEIGHT = string.Empty;
                    string sFULLWEIGHT = string.Empty;
                    string sFIRSTTIME = string.Empty;
                    string sSECONDTIME = string.Empty;
                    if (sPRV_KYERYANG12.Equals("1"))
                    {
                        if (sPRV_KERATYPE.Equals("입고"))
                        {
                            sFIRSTTIME = dtSecondTime.ToString("HH:mm");
                            sSECONDTIME = dtFirstTime.ToString("HH:mm");
                            sEMPTYWEIGHT = iPRV_SECONDWEIGHT.ToString();
                            sFULLWEIGHT = iPRV_FIRSTWEIGHT.ToString();

                            sLogMsg += string.Format("대지중량:{0} | 1차시각:{1} | 공차중량:{2} | 2차시각:{3} | "
                                , sFULLWEIGHT, sFIRSTTIME, sEMPTYWEIGHT, string.Empty);
                        }
                        else
                        {
                            sFIRSTTIME = dtFirstTime.ToString("HH:mm");
                            sSECONDTIME = dtSecondTime.ToString("HH:mm");
                            sEMPTYWEIGHT = iPRV_FIRSTWEIGHT.ToString();
                            sFULLWEIGHT = iPRV_SECONDWEIGHT.ToString();

                            sLogMsg += string.Format("공차중량:{0} | 1차시각:{1} | 대지중량:{2} | 2차시각:{3} | "
                                , sEMPTYWEIGHT, sFIRSTTIME, sFULLWEIGHT, string.Empty);
                        }
                    }
                    else
                    {
                        if (sPRV_KERATYPE.Equals("입고"))
                        {
                            sFIRSTTIME = dtSecondTime.ToString("HH:mm");
                            sSECONDTIME = dtFirstTime.ToString("HH:mm");
                            sEMPTYWEIGHT = iPRV_SECONDWEIGHT.ToString();
                            sFULLWEIGHT = iPRV_FIRSTWEIGHT.ToString();

                            sLogMsg += string.Format("대지중량:{0} | 1차시각:{1} | 공차중량:{2} | 2차시각:{3} | "
                                , sFULLWEIGHT, sFIRSTTIME, sEMPTYWEIGHT, sSECONDTIME);
                        }
                        else
                        {
                            sFIRSTTIME = dtFirstTime.ToString("HH:mm");
                            sSECONDTIME = dtSecondTime.ToString("HH:mm");
                            sEMPTYWEIGHT = iPRV_FIRSTWEIGHT.ToString();
                            sFULLWEIGHT = iPRV_SECONDWEIGHT.ToString();

                            sLogMsg += string.Format("공차중량:{0} | 1차시각:{1} | 대지중량:{2} | 2차시각:{3} | "
                                , sEMPTYWEIGHT, sFIRSTTIME, sFULLWEIGHT, sSECONDTIME);
                        }
                    }
                    
                    sLogMsg += string.Format("검수자:{0} | 검수비고:{1} | 감량중량:{2} | 감량사유:{3} | 수정사유:삭제"
                        , sPRV_EMPNM, sPRV_GUMSUBIGO, iPRV_CHAGAM.ToString(), sPRV_J_STATE);

                    sLogMsg = sLogMsg.Length > 500 ? sLogMsg.Substring(0, 500) : sLogMsg;

                    sSTD_COLS += string.Format("{0}/{1}/순번:{2}/{3}/차번:{4}/{5}"
                        , sPRV_J_DATE, sPRV_KERATYPE, sPRV_SUN, sPRV_COMPANY, sPRV_J_BNUM, sPRV_GUBUN1);

                    sREF_RMK += string.Format("Table(MESURING, JUNPYOID : {0})", sJunpyoId);
                    
                }

                //LogInsert
                //Reference : #0002
                if (iLogCnt > 0)
                {
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" INSERT INTO ZSYS_LOG ");
                    strSql.AppendLine("           ( OCCUR_DATE ");
                    strSql.AppendLine("           , USRCD ");
                    strSql.AppendLine("           , LOG_SEQ ");
                    strSql.AppendLine("           , EDIT_KIND ");
                    strSql.AppendLine("           , PGM_ID ");
                    strSql.AppendLine("           , ACS_IP ");
                    strSql.AppendLine("           , STD_COLS ");
                    strSql.AppendLine("           , REF_RMK ");
                    strSql.AppendLine("           , EDIT_RMK ) ");
                    strSql.AppendLine("     VALUES( @OCCUR_DATE ");
                    strSql.AppendLine("           , @USRCD ");
                    strSql.AppendLine("           , ( SELECT ISNULL(MAX(X1.LOG_SEQ), 0) + 1 FROM ZSYS_LOG X1 WHERE X1.OCCUR_DATE = @OCCUR_DATE AND X1.USRCD = @USRCD ) ");
                    strSql.AppendLine("           , 'D' ");
                    strSql.AppendLine("           , @PGM_ID ");
                    strSql.AppendLine("           , @ACS_IP ");
                    strSql.AppendLine("           , @STD_COLS ");
                    strSql.AppendLine("           , @REF_RMK ");
                    strSql.AppendLine("           , @EDIT_RMK ) ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@OCCUR_DATE", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                    cmd.Parameters.AddWithValue("@PGM_ID", this.Name);
                    cmd.Parameters.AddWithValue("@ACS_IP", ComnEtcFunc.GetLocalIP());
                    cmd.Parameters.AddWithValue("@STD_COLS", sSTD_COLS);
                    cmd.Parameters.AddWithValue("@REF_RMK", sREF_RMK);
                    cmd.Parameters.AddWithValue("@EDIT_RMK", sLogMsg);
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                }

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" DELETE FROM MESURING ");
                strSql.AppendLine("       WHERE JUNPYOID = @JUNPYOID ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.AddWithValue("@JUNPYOID", sJunpyoId);
                cmd.ExecuteNonQuery();
            
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                int idx = GridViewRetr.FocusedRowHandle;
                BtnRetr.PerformClick();

                GridViewRetr.FocusedRowHandle = idx-1;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }

        }

        private void pictureEdit1_DoubleClick(object sender, EventArgs e)
        {
            SendImage();
        }

        private void SendImage()
        {
            Dictionary<string, Image> dicParams = new Dictionary<string, Image>();

            dicParams.Add("1_1", pictureEdit1.Image == null ? Properties.Resources.No_Img : pictureEdit1.Image);
            dicParams.Add("1_2", pictureEdit2.Image == null ? Properties.Resources.No_Img : pictureEdit2.Image);
            dicParams.Add("1_3", pictureEdit3.Image == null ? Properties.Resources.No_Img : pictureEdit3.Image);
            dicParams.Add("2_1", pictureEdit4.Image == null ? Properties.Resources.No_Img : pictureEdit4.Image);
            dicParams.Add("2_2", pictureEdit5.Image == null ? Properties.Resources.No_Img : pictureEdit5.Image);
            dicParams.Add("2_3", pictureEdit6.Image == null ? Properties.Resources.No_Img : pictureEdit6.Image);

            MesImageRetr form = new MesImageRetr();
            form._MESURE_IMAGE = dicParams;
            form.Show();
        }

        private void DateEditTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}