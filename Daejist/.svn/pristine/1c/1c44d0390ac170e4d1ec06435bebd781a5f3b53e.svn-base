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
using DevExpress.Data;
using DevExpress.XtraGrid.Views.Grid;
using ComLib;
using DevExpress.XtraGrid;
using MySql.Data.MySqlClient;
using DevExpress.DataAccess.Excel;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.EditForm.Helpers.Controls;
using DevExpress.XtraEditors.Controls;
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System.Runtime.InteropServices;
using System.Data.SqlClient;

/*
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 수정일자 : 2021-02-22
 * 수정자 : 고혜성
 * 수정내용 : (현업요청)
 *            1. INLIST 업체담당자 컬럼추가됨에 따라 입출고 1/2차 INSERT/UPDATE 쿼리에 CHRG_ID, CHRG_NM 컬럼 반영
 *             저장관련 메소드 참고
 *        
 * 수정일자 : 2021-02-22
 * 수정자 : 고혜성
 * 수정내용 : (현업요청)
 *           1. MESURING은 CHRG_ID를 제외
 *           
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 *            
 * 수정일자 : 2021-03-09
 * 수정자   : 고혜성
 * 참조ID   : #00001
 * 수정내용 : (현업요청)
 *            1. LOG 로직 추가
 *            
 * 수정일자 : 2021-03-11
 * 수정자   : 고혜성
 * 참조 ID  : #00002
 * 수정내용 : 감량중량은 계근프로그램에서 입력하므로 Readonly 해제에서 제외시킴
 * 
 * 수정일자 : 2021-03-12
 * 수정자   : 고혜성
 * 참조 ID  : #00003
 * 수정내용 : !!참조ID(#00002)와 연관
 *            출고일 경우 감량중량을 계근프로그램이 아닌 마감및업로드에서 입력하므로 해당 내용 반영
 *          
 * 수정일자 : 2021-03-17
 * 수정자   : 고혜성
 * Reference Key : #00003
 * 수정내용 : (현업요청)
 *            1. 로그작업
 *              1-1) 기본사항/변경사항을 나누어 입력
 *              
 * 수정일자 : 2021-03-18
 * 수정자   : 고혜성
 * 수정내용 : (현업요청) #00003 관련
 *            1. 기본항목에 입력되는 값 단순화
 *            
 * 수정일자 : 2021-03-22
 * 수정자   : 고혜성
 * Reference Key : #00004
 * 수정내용 : (현업요청)
 *            1. 검수내역조회와 관련하여 최초 감량을 저장하여 불러오기 위하여
 *              감량이 변경되었을 경우 별도 테이블에 감량값 저장
 *              
 * 수정일자 : 2021-03-24
 * 수정자   : 고혜성
 * Reference Key : #00005
 * 수정내용 : (현업요청)
 *            스크랩반품, 슈레더반품(JAJAE)의 경우 리스트에 나타나지 않도록 쿼리수정
 *            
 * 수정일자 : 2021-03-26
 * 수정자   : 고혜성
 * Reference Key : #00006
 * 수정내용 : (현업요청)
 *            1. #00003와 연관 -> 마감취소시 미적용된 부분 존재
 *            
 * 수정일자 : 2022-09-06
 * 수정자   : 정은영
 * Reference Key : #00007
 * 수정내용 : (현업요청)
 *            1. 직송업로드 운반비 업체 추가 및 기타비용에 운반비 및 운반비 업체 저장
 *            
 * 수정일자 : 2022-12-09
 * 수정자   : 정은영
 * Reference Key : #00008
 * 수정내용 :(현업요청)
 *           1. 매출의 경우 기준단가를 현대제철로 변경
 *  
 *            
 */
namespace AccAdm
{
    public partial class AccMeasureCloseDev : DevExpress.XtraEditors.XtraForm
    {
        public AccMeasureCloseDev()
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

        public DataRow rowUserInfo { get; set; }
        private bool clickBtnModifyTF;
        public GridView[] arrGrdView;
        private void AccMeasureCloseDev_Load(object sender, EventArgs e)
        {
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            //ComnEtcFunc.SetDateFromToValue(DtpRetr, DtpRetrEnd);
            DtpRetr.EditValue = DateTime.Now;
            DtpRetrEnd.EditValue = DateTime.Now;
            DtpClose.EditValue = DtpRetr.EditValue;
            
            ChkComboGb.CheckAll();

            repositoryItemDateEdit1.NullDate = DateTime.MinValue;
            repositoryItemDateEdit1.NullText = String.Empty;
            
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

            clickBtnModifyTF = false;
            BtnModify.Text = "재마감(F4)";
        }

        private void AccMeasureCloseDev_TextChanged(object sender, EventArgs e)
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

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DtpRetr, DtpRetrEnd))
            {
                XtraMessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DtpRetr.EditValue = DtpRetrEnd.EditValue;
                return;
            }

            string sYmd = DtpRetr.EditValue.ToString().Substring(0, 10);
            string sEndYmd = DtpRetrEnd.EditValue.ToString().Substring(0, 10);
            //DtpClose.EditValue = DtpRetr.EditValue;
            string sClo = RdGrClose.EditValue.ToString();
            string sGb = "'',";

            foreach(CheckedListBoxItem item in ChkComboGb.Properties.Items)
            {
                if(item.CheckState == CheckState.Checked)
                {
                    sGb = sGb + item.Value + ",";
                }
            }

            sGb = sGb.Substring(0, sGb.Length - 1);

            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            #region mariaDB
            //strSql.AppendLine(" SELECT  IFNULL(A.J_Check, '0') AS CLO ");
            //strSql.AppendLine("       , A.SUN AS SUN ");
            //strSql.AppendLine("       , A.KERATYPE AS GB ");
            //strSql.AppendLine("       , DATE_FORMAT(A.J_DATE, '%Y-%m-%d') AS WEIGHT_YMD ");
            //strSql.AppendLine("       , DATE_FORMAT(A.AGREE_DATE, '%Y-%m-%d') AS CLO_YMD ");
            //strSql.AppendLine("       , A.J_BNUM AS CARNO  ");
            //strSql.AppendLine("       , CASE WHEN A.KERATYPE = '직송' THEN A.MAIPCHER ELSE '' END MAIPCHER ");      //2019-10-27 이차장님 요청으로 추가(직송 부분의 매입처를 보고 싶다고 함)
            //strSql.AppendLine("       , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER ");
            //strSql.AppendLine("       , A.GUBUN1 AS GRADE ");
            //strSql.AppendLine("       , A.SECONDWEIGHT AS TOT ");
            //strSql.AppendLine(" 	  , A.FIRSTTIME AS TIME1 ");
            ////strSql.AppendLine("       , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.SECONDTIME, '%H : %I') ELSE DATE_FORMAT(A.FIRSTTIME, '%H : %I') END AS TIME1 ");
            //strSql.AppendLine("       , A.FIRSTWEIGHT AS EMPAMT  ");
            ////strSql.AppendLine("       , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.FIRSTTIME, '%H : %I') ELSE DATE_FORMAT(A.SECONDTIME, '%H : %I') END AS TIME2 ");
            //strSql.AppendLine(" 	  , A.SECONDTIME AS TIME2 ");
            //strSql.AppendLine("       , A.WEIGHT AS OWN ");
            //strSql.AppendLine("       , A.CUSTOMWEIGHT AS CUSTOM ");
            //strSql.AppendLine("       , CASE WHEN A.KERATYPE = '출고' THEN A.OCHAGAM ELSE A.ICHAGAM END AS REDUCE ");
            //strSql.AppendLine("       , A.LOSSWEIGHT AS LOSS  ");
            //strSql.AppendLine("       , CASE WHEN A.KERATYPE = '출고' THEN A.OWEIGHT ELSE A.IWEIGHT END AS AMOUNT ");
            //strSql.AppendLine("       , A.TRANSPORTKUMAK AS COST ");
            //strSql.AppendLine("       , A.K_NAME AS GEA  ");
            //strSql.AppendLine("       , A.GUMSUBIGO AS RMK  ");
            //strSql.AppendLine("       , B.EMP_NM AS INSPECTOR ");
            //strSql.AppendLine("       , A.HALINYUL AS RATE ");
            //strSql.AppendLine("       , A.IDANGA AS IDANGA  ");
            //strSql.AppendLine("       , A.ODANGA AS ODANGA  ");
            //strSql.AppendLine("       , A.JUNPYOID AS JUNPYOID ");
            //strSql.AppendLine("       , A.WEIGHT_GUBUN ");
            //strSql.AppendLine("       , IF(A.KERATYPE = '입고', IPCHULGO_MAIPID, IPCHULGO_MACHULID) AS J_ID ");
            //strSql.AppendLine("       , A.MAIPCHER AS SUPPLYER ");
            //strSql.AppendLine("       , A.J_COMPANY AS PURCHASER ");
            //strSql.AppendLine("       , A.J_ASSIGNID AS PURCHASER_CD ");
            //strSql.AppendLine("       , A.OCHAGAM AS OUT_LOOSE_WEIGHT ");
            //strSql.AppendLine("       , A.OWEIGHT AS OUT_WEIGHT ");
            //strSql.AppendLine("       , A.OKONGKEP AS OUT_AMT ");
            //strSql.AppendLine("       , A.SEAK_POHAM ");
            //strSql.AppendLine("    FROM MESURING A    ");
            //strSql.AppendLine("    LEFT OUTER JOIN HR_EMP_BASIS B ");
            //strSql.AppendLine("      ON A.GUMSU_SERIAL = B.EMP_ID "); //부가세 조회쿼리 해야됨
            //strSql.AppendLine("   WHERE A.J_DATE >= '" + sYmd + "' ");
            //strSql.AppendLine("     AND A.J_DATE <= '" + sEndYmd + "' ");
            //strSql.AppendLine("     AND ('*' = '" + sClo + "' OR ('*' <> '" + sClo + "' AND IFNULL(A.J_Check, '') = '" + sClo + "')) ");
            //strSql.AppendLine("     AND A.KERATYPE IN (" + sGb + ") ");
            //strSql.AppendLine("     AND A.FIRSTWEIGHT != '0' ");
            //strSql.AppendLine("     AND A.SECONDWEIGHT != '0' ");
            //strSql.AppendLine("     AND A.J_SERIAL NOT IN ('4049042', '5050042') "); //#00005
            ////strSql.AppendLine("     AND (A.J_COMPANY <> '재고이동'  ");//HARDCODING
            ////strSql.AppendLine("     OR A.J_SERIAL NOT IN('2025163', '9000'))");//HARDCODING(ASR, 유류)
            //strSql.AppendLine("   ORDER BY A.J_DATE, A.JUNPYOID ");
            #endregion

            strSql.AppendLine("SELECT ISNULL(A.J_Check, '0') AS CLO                                               ");
            strSql.AppendLine("     , A.SUN AS SUN                                                                ");
            strSql.AppendLine("     , A.KERATYPE AS GB                                                            ");
            strSql.AppendLine("     , A.J_DATE AS WEIGHT_YMD                                                      ");
            strSql.AppendLine("     , A.AGREE_DATE AS CLO_YMD                                                     ");
            strSql.AppendLine("     , A.J_BNUM AS CARNO                                                           ");
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '직송' THEN A.MAIPCHER ELSE '' END MAIPCHER          ");
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN A.MAIPCHER ELSE A.J_COMPANY END AS DEALER");
            strSql.AppendLine("     , A.GUBUN1 AS GRADE                                                           ");
            strSql.AppendLine("     , A.SECONDWEIGHT AS TOT                                                       ");
 	        strSql.AppendLine("     , A.FIRSTTIME AS TIME1                                                        ");
            strSql.AppendLine("     , A.FIRSTWEIGHT AS EMPAMT                                                     ");
 	        strSql.AppendLine("     , A.SECONDTIME AS TIME2                                                       ");
            strSql.AppendLine("     , A.WEIGHT AS OWN                                                             ");
            strSql.AppendLine("     , A.CUSTOMWEIGHT AS CUSTOM                                                    ");
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '출고' THEN A.OCHAGAM ELSE A.ICHAGAM END AS REDUCE   ");
            strSql.AppendLine("     , A.LOSSWEIGHT AS LOSS                                                        ");
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '출고' THEN A.OWEIGHT ELSE A.IWEIGHT END AS AMOUNT   ");
            strSql.AppendLine("     , A.TRANSPORTKUMAK AS COST                                                    ");
            strSql.AppendLine("     , A.K_NAME AS GEA                                                             ");
            strSql.AppendLine("     , A.GUMSUBIGO AS RMK                                                          ");
            strSql.AppendLine("     , B.EMP_NM AS INSPECTOR                                                       ");
            strSql.AppendLine("     , A.HALINYUL AS RATE                                                          ");
            strSql.AppendLine("     , A.IDANGA AS IDANGA                                                          ");
            strSql.AppendLine("     , A.ODANGA AS ODANGA                                                          ");
            strSql.AppendLine("     , A.JUNPYOID AS JUNPYOID                                                      ");
            strSql.AppendLine("     , A.WEIGHT_GUBUN                                                              ");
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN IPCHULGO_MAIPID ELSE IPCHULGO_MACHULID END AS J_ID");
            strSql.AppendLine("     , A.MAIPCHER AS SUPPLYER                                 ");
            strSql.AppendLine("     , A.J_COMPANY AS PURCHASER                               ");
            strSql.AppendLine("     , A.J_ASSIGNID AS PURCHASER_CD                           ");
            strSql.AppendLine("     , A.OCHAGAM AS OUT_LOOSE_WEIGHT                          ");
            strSql.AppendLine("     , A.OWEIGHT AS OUT_WEIGHT                                ");
            strSql.AppendLine("     , A.OKONGKEP AS OUT_AMT                                  ");
            strSql.AppendLine("     , A.SEAK_POHAM                                           ");
            strSql.AppendLine("  FROM MESURING A                                             ");
            strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS B                              ");
            strSql.AppendLine("    ON A.GUMSU_SERIAL = B.EMP_ID                              ");
            strSql.AppendLine(" WHERE A.J_DATE BETWEEN '"+ sYmd + "' AND '"+ sEndYmd + "'             ");
            strSql.AppendLine("     AND ('*' = '" + sClo + "' OR ('*' <> '" + sClo + "' AND ISNULL(A.J_Check, '') = '" + sClo + "')) ");
            strSql.AppendLine("     AND A.KERATYPE IN (" + sGb + ") ");
            strSql.AppendLine("     AND A.FIRSTWEIGHT != '0' ");
            strSql.AppendLine("     AND A.SECONDWEIGHT != '0' ");
            strSql.AppendLine("     AND A.J_SERIAL NOT IN ('4049042', '5050042') ");
            strSql.AppendLine("     AND (A.MaipCherID != 6303870006 AND A.J_AssignID != 6303870006) ");//재고이동 제외(2022-02-16)
            strSql.AppendLine(" ORDER BY A.J_DATE, A.JUNPYOID                                ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = null;
            GridRetr.DataSource = dt;

            Cursor = Cursors.Default;

            clickBtnModifyTF = false;
            bReCloseYn = false;
            BtnModify.Text = "재마감(F4)";
        }

        private int iUpdtRowIdx;

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sCloseDt = DtpClose.EditValue?.ToString();
            if (string.IsNullOrEmpty(sCloseDt))
            {
                XtraMessageBox.Show("마감일자를 입력하세요.");
                DtpClose.Focus();
                return;
            }

            int[] selectedRows = GridViewRetr.GetSelectedRows();

            SetClose(selectedRows);

            /*
             * 2020-06-24
             * 단가가 입력이 안되어도 마감이 되도록 수정
             * 단가는 단가입력에서 입력
             */
            //iUpdtRowIdx = ChkBfrSave(selectedRows);

            //if (iUpdtRowIdx >= 0)
            //{
            //    XtraMessageBox.Show(Convert.ToString(iUpdtRowIdx + 1) + "번째 행의 입고단가 또는 출고단가가 입력되지 않았습니다.\r\n입고단가 및 출고단가를 저장하세요");
            //    GridViewRetr.FocusedRowHandle = iUpdtRowIdx;
            //    GridViewRetr.ShowEditForm();
            //    return;
            //}
            //else
            //{
            //    SetClose(selectedRows); 
            //    BtnRetr_Click(null, null);
            //}
        }

        private int ChkBfrSave(int[] selectedRows)
        {
            int iRtn = -1;
            int iRowIdx = 0;
            double dIDanGa = 0;
            double dODanGa = 0;
            string sGb = string.Empty;

            for (int i = 0; i < selectedRows.Length; i++)
            {
                iRowIdx = selectedRows[i];
                sGb = Convert.ToString(GridViewRetr.GetRowCellValue(iRowIdx, "GB"));
                dIDanGa = Convert.ToDouble(GridViewRetr.GetRowCellValue(iRowIdx, "IDANGA"));
                dODanGa = Convert.ToDouble(GridViewRetr.GetRowCellValue(iRowIdx, "ODANGA"));

                if (sGb.Equals("직송") && (dIDanGa == 0 || dODanGa == 0))
                {
                    iRtn = iRowIdx;
                }
                else if (sGb.Equals("입고") && dIDanGa == 0)
                {
                    iRtn = iRowIdx;
                }
                else if (sGb.Equals("출고") && dODanGa == 0)
                {
                    iRtn = iRowIdx;
                }
            }
            return iRtn;
        }

        public string REMARK_RESULT = string.Empty;
        private void SetClose(int[] selectedRows)
        {
            if (selectedRows.Length == 0)
            {
                XtraMessageBox.Show("마감하고자 하는 행을 선택해주세요.");
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                #region// 1.IPCHULGO 테이블에 데이터 생성 하고 나서 J_ID 가져와서 나머지 테이블과 SYNC 맞춤(J_ID AUTO INCREMENT)
                // 2.INLIST 테이블 데이터 생성
                // 3.MESURING TABLE 수정 J_CHECK, IPCHULGO_MAIPID, IPCHULGO_MACHULID
                // 4.전표 데이터 생성

                #endregion

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < selectedRows.Length; i++)
                {
                    Cursor = Cursors.WaitCursor;

                    int iRowIdx = selectedRows[i]; 
                    string sCloseYN = GridViewRetr.GetRowCellValue(iRowIdx, "CLO").ToString();
                    
                    string sJunpyoid = GridViewRetr.GetRowCellValue(iRowIdx, "JUNPYOID").ToString();
                    string sKeraType = GridViewRetr.GetRowCellValue(iRowIdx, "GB").ToString();
                    /*
                     * 2020-08-27 부가세 포함여부
                     * INLIST / SEAKPOHAM에 들어갈 값이며
                     * N일 경우 포함 / Y일 경우 미포함 처리
                     */
                    string sSeakPoham = GridViewRetr.GetRowCellValue(iRowIdx, GridColVatYn)?.ToString();
                    
                    string sSlipKind = string.Empty;
                    string sExecDept = "2000"; // 집행부서
                    string sBsnsGb = string.Empty;
                    string sKey1 = string.Empty;
                    string sKey2 = string.Empty;
                    string sKey3 = string.Empty;
                    string sid = FmMainToolBar2.UserID;
                    string sNewSlipNo = string.Empty;
                    string sJid = string.Empty;
                    string sJLater = Convert.ToString(DtpClose.EditValue).Substring(0, 10);

                    string sCarryCost = string.Empty;

                    sCarryCost = GridViewRetr.GetRowCellValue(iRowIdx, "COST")?.ToString();
                    double dCarryCost = string.IsNullOrEmpty(sCarryCost) ? 0 : Convert.ToDouble(sCarryCost);

                    //string sCloseGb = RdgbMesuringGB.EditValue?.ToString(); //WEIGHT_GUBUN

                    string sCloseGb = GridViewRetr.GetRowCellValue(iRowIdx, "WEIGHT_GUBUN")?.ToString();

                    string sLandedWeight = string.Empty;
                    string sCompWeight = string.Empty;
                    string sLooseWeight = string.Empty;
                    string sLossWeight = string.Empty;

                    //sLandedWeight = TxtLandedWeight.EditValue?.ToString();
                    sLandedWeight = GridViewRetr.GetRowCellValue(iRowIdx, "OWN")?.ToString();
                    double dLandedWeight = string.IsNullOrEmpty(sLandedWeight) ? 0 : Convert.ToDouble(sLandedWeight);

                    //sCompWeight = TxtCompWeight.EditValue?.ToString();
                    sCompWeight = GridViewRetr.GetRowCellValue(iRowIdx, "CUSTOM")?.ToString();
                    double dCompWeight = string.IsNullOrEmpty(sCompWeight) ? 0 : Convert.ToDouble(sCompWeight);

                    //sLooseWeight = TxtLooseWeight.EditValue?.ToString();
                    sLooseWeight = GridViewRetr.GetRowCellValue(iRowIdx, "REDUCE")?.ToString();
                    double dLooseWeight = string.IsNullOrEmpty(sLooseWeight) ? 0 : Convert.ToDouble(sLooseWeight);

                    //sLossWeight = TxtLossWeight.EditValue?.ToString();
                    sLossWeight = GridViewRetr.GetRowCellValue(iRowIdx, "LOSS")?.ToString();
                    double dLossWeight = string.IsNullOrEmpty(sLossWeight) ? 0 : Convert.ToDouble(sLossWeight);
                    
                    double dResultWeight = 0;
                    double dResultLossWeight = 0;
                    if (sCloseGb.Equals("0")) //당사, 업체에 따라 결과값 결정
                    {
                        //dResultLossWeight = dLandedWeight - dResultWeight;
                        dResultWeight = dLandedWeight - dLooseWeight; //감량중량
                        dResultLossWeight = dLandedWeight - dCompWeight;
                    }
                    else  //업체 시
                    {
                        //dResultWeight = dLandedWeight - dLossWeight; //Loss중량(업체기준)
                        //dResultLossWeight = dLandedWeight - dResultWeight;
                        //dResultLossWeight = dCompWeight - dLossWeight;
                        //dResultWeight = dCompWeight - dLossWeight;
                        dResultWeight = dCompWeight - dLooseWeight;
                        dResultLossWeight = dLandedWeight - dCompWeight;
                    }

                    DataTable dt = new DataTable();
                    string sJ_ID = GridViewRetr.GetRowCellValue(iRowIdx, "J_ID")?.ToString();
                    
                    //마감승인된 자료 재마감 불가 (정은영)
                    strSql.Clear();
                    strSql.AppendLine(" SELECT APRV_YN ");
                    strSql.AppendLine("   FROM INLIST  ");
                    strSql.AppendLine("  WHERE J_ID = (");
                    strSql.AppendLine(" SELECT CASE WHEN KERATYPE = '입고' THEN IPCHULGO_MAIPID ");
                    strSql.AppendLine(" 	   		WHEN KERATYPE = '출고' THEN IPCHULGO_MACHULID END");
                    strSql.AppendLine("   FROM MESURING ");
                    strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "')");

                    DataTable dtAprvYN = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    string sRemark = string.Empty;
                    bool bReCloseGb = false;
                    if (dtAprvYN.Rows.Count > 0)
                    {
                        string sAprvYN = dtAprvYN.Rows[0][0].ToString();

                        //마감승인 시
                        if (sAprvYN.Equals("Y"))
                        {
                            XtraMessageBox.Show("마감승인된 자료는 재마감이 불가능합니다.");
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            Cursor = Cursors.Default;
                            return;
                        }
                        else
                        {
                            bReCloseGb = true; //재마감 구분
                            IN06001F01 frm = new IN06001F01();
                            frm.PARENT_FORM = this;
                            frm.RESULT_WEIGHT = dResultWeight;
                            frm.ROW_INFO = GridViewRetr.GetDataRow(iRowIdx);
                            Cursor = Cursors.Default;
                            if (frm.ShowDialog() == DialogResult.OK)
                            {

                            }
                        }
                    }

                    Cursor = Cursors.WaitCursor;

                    /*
                     * 2021-03-09
                     * 현업요청
                     * 변경된 마감 건에 대하여 LOG 메시지 로직추가
                     * Reference : #00001
                     */
                    #region[LOG INSERT]

                    DataTable dtPrv = null;
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT CAST(AGREE_DATE AS VARCHAR) AS AGREE_DATE ");
                    strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN ICHAGAM ELSE OCHAGAM END AS CHAGAM ");
                    strSql.AppendLine("      , CUSTOMWEIGHT ");
                    strSql.AppendLine("      , LOSSWEIGHT ");
                    strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN IWEIGHT ELSE OWEIGHT END AS WEIGHT ");
                    strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN IDANGA ELSE ODANGA END AS DANGA ");
                    strSql.AppendLine("      , TRANSPORTKUMAK ");
                    strSql.AppendLine("      , WEIGHT_GUBUN ");
                    strSql.AppendLine("      , SEAK_POHAM ");
                    strSql.AppendLine("   FROM MESURING ");
                    strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoid + " ");

                    dtPrv = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    strSql.Clear();

                    string sLog_Msg = string.Empty;
                    string sReCloseGb = bReCloseGb ? "[재마감]" : "[마감]";

                    /*
                     * #00003
                     */
                    string sSTD_COLS = string.Empty;
                    string sREF_RMK = string.Empty;
                    sSTD_COLS = string.Format("{0}/{1}/순번:{2}/{3}/차번:{4}/{5}"
                                            , GridViewRetr.GetRowCellValue(iRowIdx, GridColYmd)?.ToString()
                                            , sKeraType
                                            , GridViewRetr.GetRowCellValue(iRowIdx, GridColSeq)?.ToString()
                                            , GridViewRetr.GetRowCellValue(iRowIdx, GridColDealer)?.ToString()
                                            , GridViewRetr.GetRowCellValue(iRowIdx, GridColCarNo)?.ToString()
                                            , GridViewRetr.GetRowCellValue(iRowIdx, GridColGrade)?.ToString());
                    int iLogCnt = 0;
                    int iCCNT = 0;
                    if (dtPrv.Rows.Count > 0)
                    {
                        sLog_Msg += sReCloseGb;

                        string sPrv_JLater = dtPrv.Rows[0]["AGREE_DATE"]?.ToString() ?? string.Empty;
                        string sPrv_WeightGubun = dtPrv.Rows[0]["WEIGHT_GUBUN"]?.ToString() ?? string.Empty;
                        string sPrv_SeakPoham = dtPrv.Rows[0]["SEAK_POHAM"]?.ToString();

                        double dPrv_Chagam = 0;
                        double dPrv_CustomWeight = 0;
                        double dPrv_LossWeight = 0;
                        double dPrv_Weight = 0;
                        double dPrv_TransportKumak = 0;

                        double dPrv_Danga = 0;

                        double.TryParse(dtPrv.Rows[0]["CHAGAM"]?.ToString(), out dPrv_Chagam);
                        double.TryParse(dtPrv.Rows[0]["CUSTOMWEIGHT"]?.ToString(), out dPrv_CustomWeight);
                        double.TryParse(dtPrv.Rows[0]["LOSSWEIGHT"]?.ToString(), out dPrv_LossWeight);
                        double.TryParse(dtPrv.Rows[0]["WEIGHT"]?.ToString(), out dPrv_Weight);
                        double.TryParse(dtPrv.Rows[0]["TRANSPORTKUMAK"]?.ToString(), out dPrv_TransportKumak);

                        double.TryParse(dtPrv.Rows[0]["DANGA"]?.ToString(), out dPrv_Danga);

                        iLogCnt = 0;

                        string sJLater_Tran = DtpClose.EditValue?.ToString().Substring(0, 10);
                        if (!sPrv_JLater.Equals(sJLater_Tran))
                        {
                            if (bReCloseGb)
                            {
                                iLogCnt++;
                                if(iCCNT++ != 0)
                                {
                                    sLog_Msg += " | ";
                                }
                                sLog_Msg += string.Format("마감일자 : {0} ▶ {1}", sPrv_JLater, sJLater_Tran);
                            }
                        }

                        if (!sPrv_WeightGubun.Equals(sCloseGb))
                        {
                            iLogCnt++;
                            string sWeightGubun_Tran_1 = sPrv_WeightGubun.Equals("0") ? "당사" : "업체";
                            string sWeightGubun_Tran_2 = sCloseGb.Equals("0") ? "당사" : "업체";

                            if (iCCNT++ != 0)
                            {
                                sLog_Msg += " | ";
                            }
                            sLog_Msg += string.Format("계근구분 : {0} ▶ {1}", sWeightGubun_Tran_1, sWeightGubun_Tran_2);
                        }

                        if (!sPrv_SeakPoham.Equals(sSeakPoham))
                        {
                            iLogCnt++;
                            string sSeakPoham_Tran_1 = sPrv_SeakPoham.Equals("N") ? "포함" : "미포함";
                            string sSeakPoham_Tran_2 = sSeakPoham.Equals("N") ? "포함" : "미포함";
                            if (iCCNT++ != 0)
                            {
                                sLog_Msg += " | ";
                            }
                            sLog_Msg += string.Format("부가세여부 : {0} ▶ {1}", sSeakPoham_Tran_1, sSeakPoham_Tran_2);
                        }

                        if (dPrv_Weight != dLandedWeight)
                        {
                            iLogCnt++;
                            if (iCCNT++ != 0)
                            {
                                sLog_Msg += " | ";
                            }
                            sLog_Msg += string.Format("대지중량 : {0} ▶ {1}", dPrv_Weight, dLandedWeight);
                        }

                        if (dPrv_Chagam != dLooseWeight)
                        {
                            iLogCnt++;
                            if (iCCNT++ != 0)
                            {
                                sLog_Msg += " | ";
                            }
                            sLog_Msg += string.Format("감량중량 : {0} ▶ {1}", dPrv_Chagam, dLooseWeight);
                        }

                        if (dPrv_CustomWeight != dCompWeight)
                        {
                            iLogCnt++;
                            if (iCCNT++ != 0)
                            {
                                sLog_Msg += " | ";
                            }
                            sLog_Msg += string.Format("업체중량 : {0} ▶ {1}", dPrv_CustomWeight, dCompWeight);
                        }

                        if (dPrv_LossWeight != dLossWeight)
                        {
                            iLogCnt++;
                            if (iCCNT++ != 0)
                            {
                                sLog_Msg += " | ";
                            }
                            sLog_Msg += string.Format("Loss중량 : {0} ▶ {1}", dPrv_LossWeight, dLossWeight);
                        }

                        if (dPrv_TransportKumak != dCarryCost)
                        {
                            iLogCnt++;
                            if (iCCNT++ != 0)
                            {
                                sLog_Msg += " | ";
                            }
                            sLog_Msg += string.Format("운반비정산 : {0} ▶ {1}", dPrv_TransportKumak, dCarryCost);
                        }

                        string sDanga = sKeraType.Equals("입고") ? GridViewRetr.GetRowCellValue(iRowIdx, GridColIDanGa)?.ToString() : GridViewRetr.GetRowCellValue(iRowIdx, GridColODanGa)?.ToString();
                        sDanga = sDanga ?? string.Empty;
                        double dDanga = 0;
                        double.TryParse(sDanga, out dDanga);
                        if (dPrv_Danga != dDanga)
                        {
                            iLogCnt++;
                            if (iCCNT++ != 0)
                            {
                                sLog_Msg += " | ";
                            }
                            sLog_Msg += string.Format("{0}단가 : {1} ▶ {2}", sKeraType, dPrv_Danga, dDanga);
                        }

                        string sAprvReason = this.REMARK_RESULT ?? string.Empty;
                        //재마감 시 사유입력값 로그에 입력
                        if (bReCloseGb && !string.IsNullOrEmpty(sAprvReason.Trim()))
                        {
                            iLogCnt++;
                            if (iCCNT++ != 0)
                            {
                                sLog_Msg += " | ";
                            }
                            sLog_Msg += string.Format("적용사유 : {0}", sAprvReason);
                        }

                        sREF_RMK += string.Format("TABLE : Mesuring / JunpyoID : {0}", sJunpyoid);

                        /*
                         * #00004
                         */
                        if (dPrv_Chagam != dLooseWeight)
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
                            sbPrv.AppendFormat("     VALUES( {0} ", sJunpyoid);
                            sbPrv.AppendFormat("           , ( SELECT ISNULL(MAX(X1.CHG_SEQ), 0) + 1 AS CHG_SEQ ");
                            sbPrv.AppendFormat("                 FROM MESURING_SEQ AS X1 ");
                            sbPrv.AppendFormat("                WHERE X1.JUNPYOID = {0} )  ", sJunpyoid);
                            sbPrv.AppendFormat("           , {0} ", dPrv_Chagam);
                            sbPrv.AppendFormat("           , CONVERT(VARCHAR(19),GETDATE(),20) ");
                            sbPrv.AppendFormat("           , {0} ); ", FmMainToolBar2.UserID);

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = sbPrv.ToString();
                            cmd.Parameters.Clear();
                            cmd.ExecuteNonQuery();
                        }
                    }

                    #endregion[LOG INSERT]
                    
                    //이차장님 요청으로 전표쪽 삭제 구현(고혜성)
                    #region[재마감 시  IPCHULGO,INLIST, ACC_AUTO_SLIP, ACC_CRT_SEQ, ACC_SLIP_MST, ACC_SLIP_DET, ACC_PAY_BACK, ACC_SLIP_EVDN, ACC_TAX_MGT, ACC_TAX,DTL 삭제]
                    if (sCloseYN.Equals("1") && !sJ_ID.Equals("0"))
                    {
                        //string sJ_ID = GridViewRetr.GetRowCellValue(iRowIdx, "J_ID")?.ToString();

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" DELETE ");
                        strSql.AppendLine("   FROM IPCHULGO ");
                        strSql.AppendLine("  WHERE J_ID = " + sJ_ID + " ");
                        
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" DELETE ");
                        strSql.AppendLine("   FROM INLIST ");
                        strSql.AppendLine("  WHERE J_ID = " + sJ_ID + " ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" DELETE ");
                        strSql.AppendLine("   FROM ACC_AUTO_SLIP ");
                        strSql.AppendLine("  WHERE PYBC_KEY = " + sJ_ID + " ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        //string sSlipYmd = string.Empty;
                        //string sSlipNo = string.Empty;

                        //sSlipYmd = sJ_ID.Substring(0, 8);
                        //sSlipNo = sJ_ID.Substring(8, 7);

                        //strSql.Clear();
                        //strSql.AppendLine(" ");
                        //strSql.AppendLine(" DELETE ");
                        //strSql.AppendLine("   FROM ACC_CRT_SEQ ");
                        //strSql.AppendLine("  WHERE BSNS_GB = 'SLIP_NO' ");
                        //strSql.AppendLine("    AND KEY1 = '" + sSlipYmd + "' ");
                        //strSql.AppendLine("    AND SEQ = '" + sSlipNo + "' ");

                        //cmd.CommandType = CommandType.Text;
                        //cmd.CommandText = strSql.ToString();
                        //cmd.ExecuteNonQuery();

                        //strSql.Clear();
                        //strSql.AppendLine(" ");
                        //strSql.AppendLine(" DELETE ");
                        //strSql.AppendLine("   FROM ACC_SLIP_MST ");
                        //strSql.AppendLine("  WHERE SLIP_YMD = '" + sSlipYmd + "' ");
                        //strSql.AppendLine("    AND SLIP_NO = '" + sSlipNo + "' ");

                        //cmd.CommandType = CommandType.Text;
                        //cmd.CommandText = strSql.ToString();
                        //cmd.ExecuteNonQuery();

                        //strSql.Clear();
                        //strSql.AppendLine(" ");
                        //strSql.AppendLine(" DELETE ");
                        //strSql.AppendLine("   FROM ACC_SLIP_DET ");
                        //strSql.AppendLine("  WHERE SLIP_YMD = '" + sSlipYmd + "' ");
                        //strSql.AppendLine("    AND SLIP_NO = '" + sSlipNo + "' ");

                        //cmd.CommandType = CommandType.Text;
                        //cmd.CommandText = strSql.ToString();
                        //cmd.ExecuteNonQuery();

                        //strSql.Clear();
                        //strSql.AppendLine(" ");
                        //strSql.AppendLine(" DELETE ");
                        //strSql.AppendLine("   FROM ACC_PAY_BACK ");
                        //strSql.AppendLine("  WHERE SLIP_YMD = '" + sSlipYmd + "' ");
                        //strSql.AppendLine("    AND SLIP_NO = '" + sSlipNo + "' ");

                        //cmd.CommandType = CommandType.Text;
                        //cmd.CommandText = strSql.ToString();
                        //cmd.ExecuteNonQuery();
                        
                    }

                    #endregion[재마감 시  IPCHULGO,INLIST, ACC_AUTO_SLIP, ACC_CRT_SEQ, ACC_SLIP_MST, ACC_SLIP_DET, ACC_PAY_BACK, ACC_SLIP_EVDN, ACC_TAX_MGT, ACC_TAX,DTL 삭제]

                    // 입고/ 직송(매입)일때
                    if (!sKeraType.Equals("출고"))  
                    {
                        sSlipKind = "0002";

                        #region // SLIP_NO 생성 

                        sBsnsGb = "SLIP_NO";
                        sKey1 = sJLater.Replace("-","");
                        sKey2 = "****";
                        sKey3 = "****";
                        sid = FmMainToolBar2.UserID;

                        strSql.Clear();
                        //strSql.AppendLine(" SELECT SET_CRT_SEQ('" + sBsnsGb + "', '" + sKey1 + "','" + sKey2 + "','" + sKey3 + "','" + sid + "')");//프로시저로 변경
                        strSql.AppendLine(" EXEC SET_CRT_SEQ '" + sBsnsGb + "', '" + sKey1 + "','" + sKey2 + "','" + sKey3 + "','" + sid + "' ");

                        dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                        sNewSlipNo = dt.Rows[0]["NEW_SEQ"].ToString();

                        sJid = sJLater.Replace("-", "") + sNewSlipNo;

                        #endregion

                        //MESURING 매입단가, 운송비 수정(단가 및 운송비 수정 삭제, 마감일자, 중량, 업체 수정 2020-05-29)
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE MESURING ");
                        strSql.AppendLine("    SET AGREE_DATE = '" + sJLater + "'");
                        strSql.AppendLine("      , ICHAGAM = " + dLooseWeight + " ");
                        strSql.AppendLine("      , CUSTOMWEIGHT = " + dCompWeight + " "); //업체중량일 시 대지중량 - 업체중량
                        strSql.AppendLine("      , LOSSWEIGHT = " + dResultLossWeight + " ");
                        strSql.AppendLine("      , IWEIGHT = " + dResultWeight + " ");
                        strSql.AppendLine("      , TRANSPORTKUMAK = " + dCarryCost + " ");
                        strSql.AppendLine("      , WEIGHT_GUBUN = " + sCloseGb + " ");
                        strSql.AppendLine("      , EDIT_GB = 'A' " ); //A : 마감 및 업로드, B : 계근프로그램, C : 단가입력
                        strSql.AppendLine("      , SEAK_POHAM = '" + sSeakPoham + "' ");
                        //strSql.AppendLine("      , CHRG_ID = ( SELECT X1.CHRG_ID FROM ACC_DEALER_CD X1 WHERE X1.DEALER_CD = MAIPCHERID )"); //2021-02-22 쿼리 업체담당자 관련 추가
                        strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "'  ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        // IPCHULGO
                        strSql.Clear();
                        strSql.AppendLine("SET IDENTITY_INSERT IPCHULGO ON                                                                    ");
                        strSql.AppendLine(" INSERT INTO IPCHULGO ");
                        strSql.AppendLine(" ( J_ID, J_TYPE, J_KIND, KERAGUBUN, J_TYPE1, J_TYPE2, J_DATE, J_LATER, J_ASSIGNID, J_COMPANY, J_BNUM ");
                        strSql.AppendLine(" , J_SUMSURYANG, J_SUMWEIGHT, J_SUMITEM, J_AMOUNT, J_KONGKEP, J_RAMOUNT, J_LAMOUNT, J_COUNT ");
                        strSql.AppendLine(" , J_BUGASE, J_CHECK, J_IPCHULGO, J_VATPOHAM, DAMDANG, CKONGKEP, P_ID, J_KYULMETHOD, FIRSTDATE ");
                        strSql.AppendLine(" , U_DATE, J_GARAGE, COSTBUGASE, J_SIGN1) ");
                        strSql.AppendLine(" SELECT '" + sJid + "' AS J_ID ");
                        strSql.AppendLine("      , '매입' AS J_TYPE ");
                        strSql.AppendLine("      , '일반' AS J_KIND ");
                        strSql.AppendLine("      , CASE WHEN A.KERATYPE <> '직송' THEN '계량'  ELSE A.KERATYPE END AS KERAGUBUN ");
                        strSql.AppendLine("      , '대체' AS J_TYPE1 ");
                        strSql.AppendLine("      , '일반' AS J_TYPE2 ");
                        strSql.AppendLine("      , CONVERT(DATE,J_DATE) ");
                        strSql.AppendLine("      , '" + sJLater + "' AS J_LATER ");
                        strSql.AppendLine("      , MAIPCHERID AS J_ASSIGNID ");
                        strSql.AppendLine("      , MAIPCHER AS J_COMPANY ");
                        strSql.AppendLine("      , J_BNUM AS J_BNUM ");
                        strSql.AppendLine("      , 1 AS J_SUMSURYANG ");
                        strSql.AppendLine("      , A.IWEIGHT AS J_SUMWEIGHT ");
                        strSql.AppendLine("      , CASE WHEN A.KERATYPE = '직송' THEN 0 ELSE A.IWEIGHT END AS J_SUMITEM ");
                        strSql.AppendLine("      , A.IWEIGHT * A.IDANGA * 1.1 AS J_AMOUNT ");
                        strSql.AppendLine("      , A.IWEIGHT * A.IDANGA  AS J_KONGKEP ");
                        strSql.AppendLine("      , A.IWEIGHT * A.IDANGA * 1.1 AS J_RAMOUNT ");
                        strSql.AppendLine("      , A.IWEIGHT * A.IDANGA * 1.1 AS J_LAMOUNT ");
                        strSql.AppendLine("      , 1 AS J_COUNT ");
                        strSql.AppendLine("      , A.IWEIGHT * A.IDANGA * 0.1 AS J_BUGASE ");
                        strSql.AppendLine("      , A.J_CHECK  ");
                        strSql.AppendLine("      , 1 AS J_IPCHULGO ");
                        strSql.AppendLine("      , '별도' AS J_VATPOHAM ");
                        strSql.AppendLine("      , '80142' AS DAMDANG ");
                        strSql.AppendLine("      , TRANSPORTKUMAK AS CKONGKEP ");
                        strSql.AppendLine("      , P_ID ");
                        strSql.AppendLine("      , '외상' AS J_KYULMETHOD ");
                        strSql.AppendLine("      , CONVERT(VARCHAR(19), GETDATE(),20) AS FIRSTDATE ");
                        strSql.AppendLine("      , CONVERT(VARCHAR(19), GETDATE(),20) AS U_DATE ");
                        strSql.AppendLine("      , J_GARAGE ");
                        strSql.AppendLine("      , A.IWEIGHT * A.IDANGA * 0.1 AS COSTBUGASE ");
                        strSql.AppendLine("      , J_DATE AS J_SIGN1       ");
                        strSql.AppendLine("   FROM MESURING A   ");
                        strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "'  ");
                        strSql.AppendLine("SET IDENTITY_INSERT IPCHULGO OFF                                                                    ");
                        strSql.AppendLine("  ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        // INLIST
                        strSql.Clear();
                        strSql.AppendLine(" INSERT INTO INLIST ");
                        strSql.AppendLine(" ( J_CHECK, J_KIND, J_ID, J_ID1, J_SERIAL, GUBUN1 , J_CUSTOM, K_NAME, J_DATE, J_LATER, KERATYPE, KERAGUBUN ");
                        strSql.AppendLine(" , MODEL, J_LOTNO, J_BOOKING, J_SPEC, ISURYANG, OSURYANG, HALIN, DANJUNG, IWEIGHT, OWEIGHT, NWEIGHT ");
                        strSql.AppendLine(" , J_UNIT, DANGA, MIDANGA, CDANGA, GDANGA, EDANGA, MAIPCHER, IKONGKEP, KONGKEP, BUGASE, CKONGKEP ");
                        strSql.AppendLine(" , MIKONGKEP, GKONGKEP, EKONGKEP, SEAKPOHAM, J_RID, J_STATE, J_GARAGE, P_ID, NICKGUBUN1 ");
                        strSql.AppendLine(" , CHRG_ID, CHRG_NM "); //2021-02-22 컬럼추가에 따라 쿼리수정
                        strSql.AppendLine(" ) ");
                        strSql.AppendLine(" SELECT '' AS J_CHECK  ");
                        strSql.AppendLine("      , '일반' AS J_KIND ");
                        strSql.AppendLine("      , '" + sJid + "' AS J_ID ");
                        strSql.AppendLine("      , MAIPCHERID AS J_ID1 ");
                        strSql.AppendLine("      , J_SERIAL ");
                        strSql.AppendLine("      , GUBUN1 ");
                        strSql.AppendLine("      , WEIGHT AS J_CUSTOM ");
                        strSql.AppendLine("      , K_NAME ");
                        strSql.AppendLine("      , CONVERT(DATE,AGREE_DATE)  AS J_DATE ");
                        //strSql.AppendLine("      , J_DATE ");
                        strSql.AppendLine("      , '" + sJLater + "' AS J_LATER ");
                        strSql.AppendLine("      , '매입' AS KERATYPE  ");
                        strSql.AppendLine("      , '계량' AS KERABUBUN ");
                        strSql.AppendLine("      , 'K1' AS MODEL ");
                        strSql.AppendLine("      , '' AS J_LOTNO ");
                        strSql.AppendLine("      , ISNULL(J_COMPANY, '') AS J_BOOKING ");
                        strSql.AppendLine("      , '당사' AS J_SPEC ");
                        strSql.AppendLine("      , 1 AS ISURYANG ");
                        strSql.AppendLine("      , 0 AS OSURYANG ");
                        strSql.AppendLine("      , ICHAGAM AS HALIN ");
                        strSql.AppendLine("      , IWEIGHT AS DANJUNG ");
                        strSql.AppendLine("      , IWEIGHT ");
                        strSql.AppendLine("      , 0 AS OWEIGHT ");
                        strSql.AppendLine("      , CUSTOMWEIGHT AS NWEIGHT ");
                        strSql.AppendLine("      , 'KG' AS J_UNIT ");
                        strSql.AppendLine("      , IDANGA AS DANGA  -- IDANGA ");
                        strSql.AppendLine("      , ( SELECT X.DANGA FROM JAJAE X WHERE X.J_SERIAL = A.J_SERIAL ) AS MIDANGA ");
                        //strSql.AppendLine("      , CASE WHEN KERATYPE = '직송' THEN IDANGA ELSE (SELECT MAX(IFNULL(X.DANGA, 0)) FROM JAJAE X WHERE X.GUBUN1 = A.GUBUN1) END AS MIDANGA ");
                        strSql.AppendLine("      , 0 AS CDANGA ");
                        strSql.AppendLine("      , 0 AS GDANGA ");
                        strSql.AppendLine("      , 0 AS EDANGA ");
                        strSql.AppendLine("      , MAIPCHERID AS MAIPCHER ");
                        strSql.AppendLine("      , IWEIGHT * IDANGA AS IKONGKEP ");
                        strSql.AppendLine("      , NULL AS KONGKEP ");
                        strSql.AppendLine("      , IWEIGHT * IDANGA * 0.1 AS BUGASE ");
                        strSql.AppendLine("      , TRANSPORTKUMAK AS CKONGKEP ");
                        strSql.AppendLine("      , CASE WHEN KERATYPE = '직송' THEN IWEIGHT * IDANGA ELSE 0 END AS MIKONGKEP ");
                        strSql.AppendLine("      , 0 AS GKONGKEP "); 
                        strSql.AppendLine("      , 0 AS EKONGKEP "); 
                        strSql.AppendLine("      , '" + sSeakPoham + "' AS SEAKPOHAM ");
                        strSql.AppendLine("      , JUNPYOID AS J_RID ");
                        strSql.AppendLine("      , LOSSWEIGHT AS J_STATE ");
                        strSql.AppendLine("      , J_GARAGE AS J_GARAGE ");
                        strSql.AppendLine("      , P_ID AS P_ID ");
                        strSql.AppendLine("      , J_STATE AS NICKGUBUN1 ");
                        strSql.AppendLine("      , ( SELECT X1.CHRG_ID FROM ACC_DEALER_CD X1 WHERE X1.DEALER_CD = A.MAIPCHERID ) ");
                        strSql.AppendLine("      , ( SELECT X2.EMP_NM   ");
                        strSql.AppendLine("            FROM ACC_DEALER_CD X1 ");
                        strSql.AppendLine("            LEFT JOIN HR_EMP_BASIS X2 ");
                        strSql.AppendLine("              ON X1.CHRG_ID = X2.EMP_ID ");
                        strSql.AppendLine("           WHERE X1.DEALER_CD = A.MAIPCHERID ) AS CHRG_NM");
                        strSql.AppendLine("   FROM MESURING A ");
                        strSql.AppendLine("   WHERE JUNPYOID = '" + sJunpyoid + "'  ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        // MESURING 수정 J_CHECK, IPCHULGO_MAIPID, -- IPCHULGO_MACHULID

                        strSql.Clear();
                        strSql.AppendLine(" UPDATE MESURING ");
                        strSql.AppendLine("    SET J_CHECK = '1' ");
                        strSql.AppendLine("      , IPCHULGO_MAIPID = '" + sJid + "'  ");
                        strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "'");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        //LogInsert
                        //Reference : #00001
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
                            strSql.AppendLine("           , 'U' ");
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
                            cmd.Parameters.AddWithValue("@EDIT_RMK", sLog_Msg);
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                        }
                        // ACC_AUTO_SLIP  테이블 추가
                        //strSql.Clear();
                        //strSql.AppendLine(" INSERT INTO ACC_AUTO_SLIP ");
                        //strSql.AppendLine("  (AUTO_SLIP_KIND, REQ_YMD, REQ_SEQ, SEQ, ACC_CD, EXEC_DEPT, SLIP_AMT, SUPL_AMT, VAT_AMT, EVDN_GB, EVDN_YMD, DBCR_GB, DEALER_CD, SLIP_ISSUE_YN, PYBC_KEY, NOTE, ENT_DT, ENT_ID, MFY_DT, MFY_ID )    ");
                        //strSql.AppendLine("   ");
                        //strSql.AppendLine(" WITH INFO AS ( ");
                        //strSql.AppendLine(" SELECT CASE WHEN B.COL = 'S' THEN A.COM_SUB_CD1 ");
                        //strSql.AppendLine("             WHEN B.COL = 'V' THEN A.COM_SUB_CD2 ");
                        //strSql.AppendLine("        ELSE A.COM_SUB_CD3 ");
                        //strSql.AppendLine("        END ACC_CD ");
                        //strSql.AppendLine("      , CASE WHEN B.COL = 'S' THEN A.CD_YN1 ");
                        //strSql.AppendLine("             WHEN B.COL = 'V' THEN A.CD_YN2 ");
                        //strSql.AppendLine("        ELSE A.CD_YN3 ");
                        //strSql.AppendLine("        END GB ");
                        //strSql.AppendLine("      , B.SEQ   ");
                        //strSql.AppendLine("   FROM COM_BASE_CD A ");
                        //strSql.AppendLine("  CROSS JOIN (SELECT 'S' AS COL, 1 AS SEQ UNION ALL SELECT 'V', 2 UNION ALL SELECT 'T', 3) B  ");
                        //strSql.AppendLine("  WHERE A.CD_GB = 'KERA_ACC_CD_IN' ");
                        //strSql.AppendLine("    AND A.COM_CD = '" + sKeraType + "' ");
                        //strSql.AppendLine(" )   ");
                        //strSql.AppendLine(" SELECT '" + sSlipKind + "' AS AUTO_SLIP_KIND ");
                        //strSql.AppendLine("      , '" + sJLater + "' AS REQ_YMD ");
                        //strSql.AppendLine("      , IFNULL((SELECT MAX(REQ_SEQ) FROM ACC_AUTO_SLIP WHERE AUTO_SLIP_KIND = '" + sSlipKind + "' AND REQ_YMD = '" + sJLater + "'), 0) + 1 AS REQ_SEQ ");
                        //strSql.AppendLine("      , B.SEQ AS SEQ ");
                        //strSql.AppendLine("      , B.ACC_CD ");
                        //strSql.AppendLine("      , '" + sExecDept + "' AS EXEC_DEPT ");
                        //strSql.AppendLine("      , CASE WHEN B.GB = 'S' THEN A.IKONGKEP  ");
                        //strSql.AppendLine("             WHEN B.GB = 'V' THEN A.IKONGKEP * 0.1  ");
                        //strSql.AppendLine("        ELSE A.IKONGKEP  * 1.1 END AS SLIP_AMT ");
                        //strSql.AppendLine("      , CASE WHEN B.GB = 'S' THEN A.IKONGKEP ELSE 0 END AS SUPL_AMT ");
                        //strSql.AppendLine("      , CASE WHEN B.GB = 'S' THEN A.IKONGKEP * 0.1 ELSE 0 END AS VAT_AMT  ");
                        //strSql.AppendLine("      , '01' AS EVDN_GB ");
                        //strSql.AppendLine("      , DATE_FORMAT(A.J_DATE, '%Y%m%d') AS EVDN_YMD ");
                        //strSql.AppendLine("      , 'D' AS DBCR_GB ");
                        //strSql.AppendLine("      , A.MAIPCHERID AS DEALER_CD ");
                        //strSql.AppendLine("      , 'N' AS SLIP_ISSUE_YN ");
                        //strSql.AppendLine("      , '" + sJid + "' AS PYBC_KEY ");
                        //strSql.AppendLine("      , CONCAT(A.MAIPCHER, ',', A.GUBUN1, ',', A.J_BNUM, ',', A.IDANGA,  ',' , A.IWEIGHT) AS NOTE  ");
                        //strSql.AppendLine("      , NOW() AS ENT_DT ");
                        //strSql.AppendLine("      , '" + sid + "' AS ENT_ID ");
                        //strSql.AppendLine("      , NOW() AS MFY_DT ");
                        //strSql.AppendLine("      , '" + sid + "' AS MFY_ID      ");
                        //strSql.AppendLine("   FROM MESURING A ");
                        //strSql.AppendLine("  INNER JOIN INFO B ");
                        //strSql.AppendLine("     ON A.JUNPYOID = '" + sJunpyoid + "'");

                        //cmd.CommandType = CommandType.Text;
                        //cmd.CommandText = strSql.ToString();
                        ////cmd.BeginExecuteNonQuery();
                        //cmd.ExecuteNonQuery();

                        //// 전표관련내역 생성
                        //CreateSlipData(cmd, sJLater, sNewSlipNo, sExecDept, sSlipKind, sJid, sJunpyoid, sKeraType);
                    }

                    // 출고/ 직송(매출) 처리
                    if (!sKeraType.Equals("입고"))
                    {
                        sSlipKind = "0003";

                        #region // SLIP_NO 생성 

                        sBsnsGb = "SLIP_NO";
                        sKey1 = sJLater.Replace("-", "");
                        sKey2 = "****";
                        sKey3 = "****";
                        sid = FmMainToolBar2.UserID;

                        strSql = new StringBuilder();
                        //strSql.AppendLine(" SELECT SET_CRT_SEQ('" + sBsnsGb + "', '" + sKey1 + "','" + sKey2 + "','" + sKey3 + "','" + sid + "')");
                        strSql.AppendLine(" EXEC SET_CRT_SEQ '" + sBsnsGb + "', '" + sKey1 + "','" + sKey2 + "','" + sKey3 + "','" + sid + "' ");

                        dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                        sNewSlipNo = dt.Rows[0]["NEW_SEQ"].ToString();

                        sJid = sJLater.Replace("-", "") + sNewSlipNo;

                        #endregion

                        //MESURING 매출단가, 운송비 수정
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE MESURING ");
                        strSql.AppendLine("    SET AGREE_DATE = '" + sJLater + "'");
                        strSql.AppendLine("      , OCHAGAM = " + dLooseWeight + " ");
                        strSql.AppendLine("      , CUSTOMWEIGHT = " + dCompWeight + " "); //업체중량일 시 대지중량 - LossWeight
                        strSql.AppendLine("      , LOSSWEIGHT = " + dResultLossWeight + " ");
                        strSql.AppendLine("      , OWEIGHT = " + dResultWeight + " ");
                        strSql.AppendLine("      , TRANSPORTKUMAK = " + dCarryCost + " ");
                        strSql.AppendLine("      , WEIGHT_GUBUN = " + sCloseGb + " ");
                        strSql.AppendLine("      , EDIT_GB = 'A' "); //A : 마감 및 업로드, B : 계근프로그램, C : 단가입력
                        strSql.AppendLine("      , SEAK_POHAM = '" + sSeakPoham + "' ");
                        //strSql.AppendLine("      , CHRG_ID = ( SELECT X1.CHRG_ID FROM ACC_DEALER_CD X1 WHERE DEALER_CD = J_ASSIGNID )");
                        strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "'  ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                        
                        // ACC.IPCHULGO
                        strSql.Clear();
                        #region mariaDB
                        //strSql.AppendLine(" INSERT INTO IPCHULGO ");
                        //strSql.AppendLine(" ( J_ID, J_TYPE, J_KIND, KERAGUBUN, J_TYPE1, J_TYPE2, J_DATE, J_LATER, J_ASSIGNID, J_COMPANY, J_BNUM ");
                        //strSql.AppendLine(" , J_SUMSURYANG, J_SUMWEIGHT, J_SUMITEM, J_AMOUNT, J_KONGKEP, J_RAMOUNT, J_LAMOUNT, J_COUNT ");
                        //strSql.AppendLine(" , J_BUGASE, J_CHECK, J_IPCHULGO, J_VATPOHAM, DAMDANG, CKONGKEP, P_ID, J_KYULMETHOD, FIRSTDATE ");
                        //strSql.AppendLine(" , U_DATE, J_GARAGE, COSTBUGASE, J_SIGN1) ");
                        //strSql.AppendLine(" SELECT '" + sJid + "' AS J_ID ");
                        //strSql.AppendLine("      , '매출' AS J_TYPE ");
                        //strSql.AppendLine("      , '일반' AS J_KIND ");
                        //strSql.AppendLine("      , CASE WHEN A.KERATYPE <> '직송' THEN '계량'  ELSE A.KERATYPE END AS KERAGUBUN ");
                        //strSql.AppendLine("      , '대체' AS J_TYPE1 ");
                        //strSql.AppendLine("      , '일반' AS J_TYPE2 ");
                        //strSql.AppendLine("      , AGREE_DATE  AS J_DATE");
                        ////strSql.AppendLine("      , J_DATE  ");
                        //strSql.AppendLine("      , '" + sJLater + "' AS J_LATER ");
                        //strSql.AppendLine("      , J_ASSIGNID AS J_ASSIGNID ");
                        //strSql.AppendLine("      , MAIPCHER AS J_COMPANY ");
                        //strSql.AppendLine("      , J_BNUM AS J_BNUM ");
                        //strSql.AppendLine("      , 1 AS J_SUMSURYANG ");
                        //strSql.AppendLine("      , A.OWEIGHT AS J_SUMWEIGHT ");
                        //strSql.AppendLine("      , CASE WHEN A.KERATYPE = '직송' THEN 0 ELSE A.OWEIGHT END AS J_SUMITEM ");
                        //strSql.AppendLine("      , A.OWEIGHT * A.ODANGA * 1.1 AS J_AMOUNT ");
                        //strSql.AppendLine("      , A.OWEIGHT * A.ODANGA AS J_KONGKEP ");
                        //strSql.AppendLine("      , A.OWEIGHT * A.ODANGA * 1.1 AS J_RAMOUNT ");
                        //strSql.AppendLine("      , A.OWEIGHT * A.ODANGA * 1.1 AS J_LAMOUNT ");
                        //strSql.AppendLine("      , 1 AS J_COUNT ");
                        //strSql.AppendLine("      , A.OWEIGHT * A.ODANGA * 0.1 AS J_BUGASE ");
                        //strSql.AppendLine("      , A.J_CHECK  ");
                        //strSql.AppendLine("      , 1 AS J_IPCHULGO ");
                        //strSql.AppendLine("      , '별도' AS J_VATPOHAM ");
                        //strSql.AppendLine("      , '38126' AS DAMDANG ");
                        //strSql.AppendLine("      , TRANSPORTKUMAK AS CKONGKEP ");
                        //strSql.AppendLine("      , P_ID ");
                        //strSql.AppendLine("      , '외상' AS J_KYULMETHOD ");
                        //strSql.AppendLine("      , NOW() AS FIRSTDATE ");
                        //strSql.AppendLine("      , NOW() AS U_DATE ");
                        //strSql.AppendLine("      , J_GARAGE ");
                        //strSql.AppendLine("      , A.OWEIGHT * A.ODANGA * 0.1 AS COSTBUGASE ");
                        //strSql.AppendLine("      , J_DATE AS J_SIGN1 ");
                        //strSql.AppendLine("   FROM MESURING A   ");
                        //strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "'  ");
                        #endregion

                        strSql.AppendLine("SET IDENTITY_INSERT IPCHULGO ON                                                                    ");
                        strSql.AppendLine("INSERT INTO IPCHULGO                                                                               ");
                        strSql.AppendLine(" (J_ID, J_TYPE, J_KIND, KERAGUBUN, J_TYPE1, J_TYPE2, J_DATE, J_LATER, J_ASSIGNID, J_COMPANY, J_BNUM");
                        strSql.AppendLine(" , J_SUMSURYANG, J_SUMWEIGHT, J_SUMITEM, J_AMOUNT, J_KONGKEP, J_RAMOUNT, J_LAMOUNT, J_COUNT        ");
                        strSql.AppendLine(" , J_BUGASE, J_CHECK, J_IPCHULGO, J_VATPOHAM, DAMDANG, CKONGKEP, P_ID, J_KYULMETHOD, FIRSTDATE     ");
                        strSql.AppendLine(" , U_DATE, J_GARAGE, COSTBUGASE, J_SIGN1)                                                          ");
                        strSql.AppendLine(" SELECT '" + sJid + "' AS J_ID                                                                  ");
                        strSql.AppendLine("      , '매출' AS J_TYPE                                                                           ");
                        strSql.AppendLine("      , '일반' AS J_KIND                                                                           ");
                        strSql.AppendLine("      , CASE WHEN A.KERATYPE <> '직송' THEN '계량'  ELSE A.KERATYPE END AS KERAGUBUN               ");
                        strSql.AppendLine("      , '대체' AS J_TYPE1                                                                          ");
                        strSql.AppendLine("      , '일반' AS J_TYPE2                                                                          ");
                        strSql.AppendLine("      , CONVERT(DATE,AGREE_DATE)  AS J_DATE                                                                      ");
                        strSql.AppendLine("      , '" + sJLater + "' AS J_LATER                                                                      ");
                        strSql.AppendLine("      , J_ASSIGNID AS J_ASSIGNID                                                                   ");
                        strSql.AppendLine("      , MAIPCHER AS J_COMPANY                                                                      ");
                        strSql.AppendLine("      , J_BNUM AS J_BNUM                                                                           ");
                        strSql.AppendLine("      , 1 AS J_SUMSURYANG                                                                          ");
                        strSql.AppendLine("      , A.OWEIGHT AS J_SUMWEIGHT                                                                   ");
                        strSql.AppendLine("      , CASE WHEN A.KERATYPE = '직송' THEN 0 ELSE A.OWEIGHT END AS J_SUMITEM                       ");
                        strSql.AppendLine("      , A.OWEIGHT* A.ODANGA * 1.1 AS J_AMOUNT                                                      ");
                        strSql.AppendLine("     , A.OWEIGHT* A.ODANGA AS J_KONGKEP                                                            ");
                        strSql.AppendLine("      , A.OWEIGHT* A.ODANGA * 1.1 AS J_RAMOUNT                                                     ");
                        strSql.AppendLine("     , A.OWEIGHT* A.ODANGA * 1.1 AS J_LAMOUNT                                                      ");
                        strSql.AppendLine("    , 1 AS J_COUNT                                                                                 ");
                        strSql.AppendLine("    , A.OWEIGHT* A.ODANGA * 0.1 AS J_BUGASE                                                        ");
                        strSql.AppendLine("   , A.J_CHECK                                                                                     ");
                        strSql.AppendLine("      , 1 AS J_IPCHULGO                                                                            ");
                        strSql.AppendLine("      , '별도' AS J_VATPOHAM                                                                       ");
                        strSql.AppendLine("      , '38126' AS DAMDANG                                                                         ");
                        strSql.AppendLine("      , TRANSPORTKUMAK AS CKONGKEP                                                                 ");
                        strSql.AppendLine("      , P_ID                                                                                       ");
                        strSql.AppendLine("      , '외상' AS J_KYULMETHOD                                                                     ");
                        strSql.AppendLine("      , CONVERT(VARCHAR(19), GETDATE(), 20) AS FIRSTDATE                                           ");
                        strSql.AppendLine("        , CONVERT(VARCHAR(19), GETDATE(), 20) AS U_DATE                                            ");
                        strSql.AppendLine("          , J_GARAGE                                                                               ");
                        strSql.AppendLine("          , A.OWEIGHT* A.ODANGA * 0.1 AS COSTBUGASE                                                ");
                        strSql.AppendLine("         , J_DATE AS J_SIGN1                                                                       ");
                        strSql.AppendLine("   FROM MESURING A                                                                                 ");
                        strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "'                                                                        ");
                        strSql.AppendLine("SET IDENTITY_INSERT IPCHULGO OFF                                                                   ");
                                                                                                                                              
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                         

                        // ACC.INLIST
                        strSql.Clear();
                        #region mariaDB
                        //strSql.AppendLine(" INSERT INTO INLIST ");
                        //strSql.AppendLine(" ( J_CHECK, J_KIND, J_ID, J_ID1, J_SERIAL, GUBUN1 , J_CUSTOM, K_NAME, J_DATE, J_LATER, KERATYPE, KERAGUBUN ");
                        //strSql.AppendLine(" , MODEL, J_LOTNO, J_BOOKING, J_SPEC, ISURYANG, OSURYANG, HALIN, DANJUNG, IWEIGHT, OWEIGHT, NWEIGHT ");
                        //strSql.AppendLine(" , J_UNIT, DANGA, MIDANGA, CDANGA, GDANGA, EDANGA, MAIPCHER, IKONGKEP, KONGKEP, BUGASE, CKONGKEP ");
                        //strSql.AppendLine(" , MIKONGKEP, GKONGKEP, EKONGKEP, SEAKPOHAM, J_RID, J_STATE, J_GARAGE, P_ID, NICKGUBUN1 ");
                        //strSql.AppendLine(" , CHRG_ID, CHRG_NM ");
                        //strSql.AppendLine(" ) ");
                        //strSql.AppendLine(" SELECT '' AS J_CHECK  ");
                        //strSql.AppendLine("      , '일반' AS J_KIND ");
                        //strSql.AppendLine("      , '" + sJid + "' AS J_ID ");
                        //strSql.AppendLine("      , J_AssignID AS J_ID1 ");
                        //strSql.AppendLine("      , J_SERIAL ");
                        //strSql.AppendLine("      , GUBUN1 ");
                        //strSql.AppendLine("      , WEIGHT AS J_CUSTOM ");
                        //strSql.AppendLine("      , K_NAME ");
                        //strSql.AppendLine("      , AGREE_DATE AS J_DATE ");
                        //strSql.AppendLine("      , J_DATE AS J_LATER ");
                        //strSql.AppendLine("      , '매출' AS KERATYPE  ");
                        //strSql.AppendLine("      , '계량' AS KERAGUBUN ");
                        //strSql.AppendLine("      , 'K1' AS MODEL ");
                        //strSql.AppendLine("      , '' AS J_LOTNO ");
                        //strSql.AppendLine("      , IFNULL(MaipCher, '') AS J_BOOKING ");
                        //strSql.AppendLine("      , CASE WHEN KERATYPE = '직송' THEN '당사' ELSE '업체' END AS J_SPEC ");
                        //strSql.AppendLine("      , 0 AS ISURYANG ");
                        //strSql.AppendLine("      , 1 AS OSURYANG ");
                        //strSql.AppendLine("      , OCHAGAM AS HALIN ");
                        //strSql.AppendLine("      , OWEIGHT AS DANJUNG ");
                        //strSql.AppendLine("      , 0 AS IWEIGHT ");
                        //strSql.AppendLine("      , OWEIGHT ");
                        //strSql.AppendLine("      , CUSTOMWEIGHT AS NWEIGHT ");
                        //strSql.AppendLine("      , 'KG' AS J_UNIT ");
                        //strSql.AppendLine("      , ODANGA AS DANGA  ");
                        //strSql.AppendLine("      , ( SELECT X.SELLPRC1 FROM JAJAE X WHERE X.J_SERIAL = A.J_SERIAL ) AS MIDANGA ");
                        ////strSql.AppendLine("      , CASE WHEN KERATYPE = '직송' THEN IDANGA ELSE (SELECT MAX(IFNULL(X.DANGA, 0)) FROM JAJAE X WHERE X.GUBUN1 = A.GUBUN1) END AS MIDANGA ");
                        //strSql.AppendLine("      , 0 AS CDANGA ");
                        //strSql.AppendLine("      , 0 AS GDANGA ");
                        //strSql.AppendLine("      , 0 AS EDANGA ");
                        //strSql.AppendLine("      , J_AssignID AS MAIPCHER ");
                        //strSql.AppendLine("      , NULL AS IKONGKEP ");
                        //strSql.AppendLine("      , OWEIGHT * ODANGA AS KONGKEP ");
                        //strSql.AppendLine("      , OWEIGHT * ODANGA * 0.1 AS BUGASE ");
                        //strSql.AppendLine("      , TRANSPORTKUMAK AS CKONGKEP ");
                        //strSql.AppendLine("      , OWEIGHT * ODANGA AS MIKONGKEP ");
                        //strSql.AppendLine("      , OWEIGHT * ODANGA * -1 AS GKONGKEP ");
                        //strSql.AppendLine("      , 0 AS EKONGKEP ");
                        //strSql.AppendLine("      , '" + sSeakPoham + "' AS SEAKPOHAM ");
                        //strSql.AppendLine("      , JUNPYOID AS J_RID           ");
                        //strSql.AppendLine("      , LOSSWEIGHT AS J_STATE ");
                        //strSql.AppendLine("      , J_GARAGE AS J_GARAGE ");
                        //strSql.AppendLine("      , P_ID AS P_ID ");
                        //strSql.AppendLine("      , J_STATE AS NICKGUBUN1 ");
                        //strSql.AppendLine("      , ( SELECT X1.CHRG_ID FROM ACC_DEALER_CD X1 WHERE X1.DEALER_CD = A.MAIPCHERID ) ");
                        //strSql.AppendLine("      , ( SELECT X2.EMP_NM   ");
                        //strSql.AppendLine("            FROM ACC_DEALER_CD X1 ");
                        //strSql.AppendLine("            LEFT JOIN HR_EMP_BASIS X2 ");
                        //strSql.AppendLine("              ON X1.CHRG_ID = X2.EMP_ID ");
                        //strSql.AppendLine("           WHERE X1.DEALER_CD = A.MAIPCHERID ) AS CHRG_NM");
                        //strSql.AppendLine("   FROM MESURING A ");
                        //strSql.AppendLine("   WHERE JUNPYOID = '" + sJunpyoid + "'  ");
                        #endregion

                        strSql.AppendLine("INSERT INTO INLIST                                                                                           ");
                        strSql.AppendLine("     (J_CHECK, J_KIND, J_ID, J_ID1, J_SERIAL, GUBUN1, J_CUSTOM, K_NAME, J_DATE, J_LATER, KERATYPE, KERAGUBUN ");
                        strSql.AppendLine("     , MODEL, J_LOTNO, J_BOOKING, J_SPEC, ISURYANG, OSURYANG, HALIN, DANJUNG, IWEIGHT, OWEIGHT, NWEIGHT      ");
                        strSql.AppendLine("     , J_UNIT, DANGA, MIDANGA, CDANGA, GDANGA, EDANGA, MAIPCHER, IKONGKEP, KONGKEP, BUGASE, CKONGKEP         ");
                        strSql.AppendLine("     , MIKONGKEP, GKONGKEP, EKONGKEP, SEAKPOHAM, J_RID, J_STATE, J_GARAGE, P_ID, NICKGUBUN1                  ");
                        strSql.AppendLine("     , CHRG_ID, CHRG_NM                                                                                      ");
                        strSql.AppendLine("     )                                                                                                       ");
                        strSql.AppendLine("     SELECT '' AS J_CHECK                                                                                    ");
                        strSql.AppendLine("          , '일반' AS J_KIND                                                                                 ");
                        strSql.AppendLine("          , '" + sJid + "' AS J_ID                                                                        ");
                        strSql.AppendLine("          , J_AssignID AS J_ID1                                                                              ");
                        strSql.AppendLine("          , J_SERIAL                                                                                         ");
                        strSql.AppendLine("          , GUBUN1                                                                                           ");
                        strSql.AppendLine("          , WEIGHT AS J_CUSTOM                                                                               ");
                        strSql.AppendLine("          , K_NAME                                                                                           ");
                        strSql.AppendLine("          , AGREE_DATE AS J_DATE                                                                             ");
                        strSql.AppendLine("          , CONVERT(DATE,J_DATE) AS J_LATER                                                                                ");
                        strSql.AppendLine("          , '매출' AS KERATYPE                                                                               ");
                        strSql.AppendLine("          , '계량' AS KERAGUBUN                                                                              ");
                        strSql.AppendLine("          , 'K1' AS MODEL                                                                                    ");
                        strSql.AppendLine("          , '' AS J_LOTNO                                                                                    ");
                        strSql.AppendLine("          , ISNULL(MaipCher, '') AS J_BOOKING                                                                ");
                        strSql.AppendLine("          , CASE WHEN KERATYPE = '직송' THEN '당사' ELSE '업체' END AS J_SPEC                                ");
                        strSql.AppendLine("          , 0 AS ISURYANG                                                                                    ");
                        strSql.AppendLine("          , 1 AS OSURYANG                                                                                    ");
                        strSql.AppendLine("          , OCHAGAM AS HALIN                                                                                 ");
                        strSql.AppendLine("          , OWEIGHT AS DANJUNG                                                                               ");
                        strSql.AppendLine("          , 0 AS IWEIGHT                                                                                     ");
                        strSql.AppendLine("          , OWEIGHT                                                                                          ");
                        strSql.AppendLine("          , CUSTOMWEIGHT AS NWEIGHT                                                                          ");
                        strSql.AppendLine("          , 'KG' AS J_UNIT                                                                                   ");
                        strSql.AppendLine("          , ODANGA AS DANGA                                                                                  ");
                        //#00008
                        strSql.AppendLine("          , ( SELECT X.SELLPRC2 FROM JAJAE X WHERE X.J_SERIAL = A.J_SERIAL ) AS MIDANGA                      ");
                        //strSql.AppendLine("          , ( SELECT X.SELLPRC1 FROM JAJAE X WHERE X.J_SERIAL = A.J_SERIAL ) AS MIDANGA                      ");
                        strSql.AppendLine("          , 0 AS CDANGA                                                                                      ");
                        strSql.AppendLine("          , 0 AS GDANGA                                                                                      ");
                        strSql.AppendLine("          , 0 AS EDANGA                                                                                      ");
                        strSql.AppendLine("          , J_AssignID AS MAIPCHER                                                                           ");
                        strSql.AppendLine("          , NULL AS IKONGKEP                                                                                 ");
                        strSql.AppendLine("          , OWEIGHT *ODANGA AS KONGKEP                                                                       ");
                        strSql.AppendLine("          , OWEIGHT *ODANGA * 0.1 AS BUGASE                                                                  ");
                        strSql.AppendLine("          , TRANSPORTKUMAK AS CKONGKEP                                                                       ");
                        strSql.AppendLine("          , OWEIGHT *ODANGA AS MIKONGKEP                                                                     ");
                        strSql.AppendLine("          , OWEIGHT *ODANGA * -1 AS GKONGKEP                                                                 ");
                        strSql.AppendLine("          , 0 AS EKONGKEP                                                                                    ");
                        strSql.AppendLine("          , '" + sSeakPoham + "' AS SEAKPOHAM                                                                                 ");
                        strSql.AppendLine("          , JUNPYOID AS J_RID                                                                                ");
                        strSql.AppendLine("          , LOSSWEIGHT AS J_STATE                                                                            ");
                        strSql.AppendLine("          , J_GARAGE AS J_GARAGE                                                                             ");
                        strSql.AppendLine("          , P_ID AS P_ID                                                                                     ");
                        strSql.AppendLine("          , J_STATE AS NICKGUBUN1                                                                            ");
                        strSql.AppendLine("          , ( SELECT X1.CHRG_ID FROM ACC_DEALER_CD X1 WHERE X1.DEALER_CD = A.MAIPCHERID )                    ");
                        strSql.AppendLine("          , ( SELECT X2.EMP_NM                                                                               ");
                        strSql.AppendLine("                FROM ACC_DEALER_CD X1                                                                        ");
                        strSql.AppendLine("                LEFT JOIN HR_EMP_BASIS X2                                                                    ");
                        strSql.AppendLine("                  ON X1.CHRG_ID = X2.EMP_ID                                                                  ");
                        strSql.AppendLine("               WHERE X1.DEALER_CD = A.MAIPCHERID ) AS CHRG_NM                                                ");
                        strSql.AppendLine("       FROM MESURING A                                                                                       ");
                        strSql.AppendLine("      WHERE JUNPYOID = '" + sJunpyoid + "'                                                                              ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        // ACC.MESURING 수정 J_CHECK, IPCHULGO_MACHULID, 

                        strSql = new StringBuilder();

                        strSql.AppendLine(" UPDATE MESURING ");
                        strSql.AppendLine("    SET J_CHECK = '1' ");
                        strSql.AppendLine("      , IPCHULGO_MACHULID = '" + sJid + "'  ");
                        strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "'");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        //LogInsert
                        //Reference : #00001
                        //재마감일시에만 적용
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
                            strSql.AppendLine("           , 'U' ");
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
                            cmd.Parameters.AddWithValue("@EDIT_RMK", sLog_Msg);
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();
                        }
                        // ACC.ACC_AUTO_SLIP  테이블 추가
                        //strSql.Clear();
                        //strSql.AppendLine(" INSERT INTO ACC_AUTO_SLIP ");
                        //strSql.AppendLine("  (AUTO_SLIP_KIND, REQ_YMD, REQ_SEQ, SEQ, ACC_CD, EXEC_DEPT, SLIP_AMT, SUPL_AMT, VAT_AMT, EVDN_GB, EVDN_YMD, DBCR_GB, DEALER_CD, SLIP_ISSUE_YN, PYBC_KEY, NOTE, ENT_DT, ENT_ID, MFY_DT, MFY_ID )    ");
                        //strSql.AppendLine("   ");
                        //strSql.AppendLine(" WITH ACC_INFO AS ( ");
                        //strSql.AppendLine(" SELECT CASE WHEN B.COL = 'S' THEN A.COM_SUB_CD1 ");
                        //strSql.AppendLine("             WHEN B.COL = 'V' THEN A.COM_SUB_CD2 ");
                        //strSql.AppendLine("        ELSE A.COM_SUB_CD3 ");
                        //strSql.AppendLine("        END ACC_CD ");
                        //strSql.AppendLine("      , CASE WHEN B.COL = 'S' THEN A.CD_YN1 ");
                        //strSql.AppendLine("             WHEN B.COL = 'V' THEN A.CD_YN2 ");
                        //strSql.AppendLine("        ELSE A.CD_YN3 ");
                        //strSql.AppendLine("        END GB ");
                        //strSql.AppendLine("      , B.SEQ   ");
                        //strSql.AppendLine("   FROM COM_BASE_CD A ");
                        //strSql.AppendLine("  CROSS JOIN (SELECT 'S' AS COL, 1 AS SEQ UNION ALL SELECT 'V', 2 UNION ALL SELECT 'T', 3) B  ");
                        //strSql.AppendLine("  WHERE A.CD_GB = 'KERA_ACC_CD_OUT' ");
                        //strSql.AppendLine("    AND A.COM_CD = '" + sKeraType + "' ");
                        //strSql.AppendLine(" )   ");
                        //strSql.AppendLine(" SELECT '" + sSlipKind + "' AS AUTO_SLIP_KIND ");
                        //strSql.AppendLine("      , '" + sJLater + "' AS REQ_YMD ");
                        //strSql.AppendLine("      , IFNULL((SELECT MAX(REQ_SEQ) FROM ACC_AUTO_SLIP WHERE AUTO_SLIP_KIND = '" + sSlipKind + "' AND REQ_YMD = '" + sJLater + "'), 0) + 1 AS REQ_SEQ ");
                        //strSql.AppendLine("      , B.SEQ AS SEQ ");
                        //strSql.AppendLine("      , B.ACC_CD ");
                        //strSql.AppendLine("      , '" + sExecDept + "' AS EXEC_DEPT ");
                        //strSql.AppendLine("      , CASE WHEN B.GB = 'S' THEN A.OKONGKEP  ");
                        //strSql.AppendLine("             WHEN B.GB = 'V' THEN A.OKONGKEP * 0.1  ");
                        //strSql.AppendLine("        ELSE A.OKONGKEP  * 1.1 END AS SLIP_AMT ");
                        //strSql.AppendLine("      , CASE WHEN B.GB = 'S' THEN A.OKONGKEP ELSE 0 END AS SUPL_AMT ");
                        //strSql.AppendLine("      , CASE WHEN B.GB = 'S' THEN A.OKONGKEP * 0.1 ELSE 0 END AS VAT_AMT  ");
                        //strSql.AppendLine("      , '01' AS EVDN_GB ");
                        //strSql.AppendLine("      , DATE_FORMAT(A.J_DATE, '%Y%m%d') AS EVDN_YMD ");
                        //strSql.AppendLine("      , 'D' AS DBCR_GB ");
                        //strSql.AppendLine("      , A.J_ASSIGNID AS DEALER_CD ");
                        //strSql.AppendLine("      , 'N' AS SLIP_ISSUE_YN ");
                        //strSql.AppendLine("      , '" + sJid + "' AS PYBC_KEY ");
                        //strSql.AppendLine("      , CONCAT(A.J_COMPANY,',', A.GUBUN1, ',', A.J_BNUM, ',', A.ODANGA, ',' , A.OWEIGHT) AS NOTE  ");
                        //strSql.AppendLine("      , NOW() AS ENT_DT ");
                        //strSql.AppendLine("      , '" + sid + "' AS ENT_ID ");
                        //strSql.AppendLine("      , NOW() AS MFY_DT ");
                        //strSql.AppendLine("      , '" + sid + "' AS MFY_ID ");
                        //strSql.AppendLine("   FROM MESURING A ");
                        //strSql.AppendLine("  INNER JOIN ACC_INFO B ");
                        //strSql.AppendLine("     ON A.JUNPYOID = '" + sJunpyoid + "'");

                        //cmd.CommandType = CommandType.Text;
                        //cmd.CommandText = strSql.ToString();
                        //cmd.ExecuteNonQuery();

                        //// 전표관련내역 생성
                        //CreateSlipData(cmd, sJLater, sNewSlipNo, sExecDept, sSlipKind, sJid, sJunpyoid, sKeraType);
                    }

                    if(bReCloseGb)
                    {
                        strSql.Clear();
                        strSql.AppendLine(" UPDATE INLIST ");
                        strSql.AppendLine("    SET REMARK = @REMARK ");
                        strSql.AppendLine("      , RECLOSE_GB = @RECLOSE_GB");
                        strSql.AppendLine("  WHERE J_ID = @J_ID   ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@REMARK", REMARK_RESULT);
                        cmd.Parameters.AddWithValue("@RECLOSE_GB", 'Y');
                        cmd.Parameters.AddWithValue("@J_ID", sJid);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                    }

                    strSql.Clear();
                    strSql.AppendLine(" ");
                }

                Cursor = Cursors.Default;

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");

                BtnRetr_Click(null, null);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        //private void CreateSlipData(SqlCommand cmd,  string sJLater, string sNewSlipNo, string sExecDept, string sSlipKind, string sJid, string sJunpyoid, string sKeraType)
        //{
        //    string sid = FmMainToolBar2.UserID;
        //    StringBuilder strSql = new StringBuilder();
            
        //    #region// ACC.ACC_SLIP_DET  테이블 추가
        //    strSql.Clear();
        //    strSql.AppendLine(" INSERT INTO ACC_SLIP_DET ");
        //    strSql.AppendLine("  (SLIP_YMD, SLIP_NO, SLIP_SEQ, ACC_CD, ACC_NM, RSLTN_DEPT_CD, RSLTN_DEPT_NM, EXEC_DEPT_CD, EXEC_DEPT_NM, DEBT_AMT ");
        //    strSql.AppendLine("  , CRDT_AMT, RMK_CD, ACC_RMK, DEALER_CD, IDT_NO, DEALER_NM, MGMT_GB1, MGMT_CD1, APRV_YN, ENT_DT, ENT_ID, MFY_DT, MFY_ID) ");
        //    strSql.AppendLine(" SELECT A.REQ_YMD AS SLIP_YMD  ");
        //    strSql.AppendLine("      , '" + sNewSlipNo + "' AS SLIP_NO  ");
        //    strSql.AppendLine("      , A.SEQ AS SLIP_SEQ  ");
        //    strSql.AppendLine("      , A.ACC_CD AS ACC_CD ");
        //    strSql.AppendLine("      , (SELECT X.ACC_NM FROM ACC_ACC_CD X WHERE X.ACC_CD = A.ACC_CD) AS ACC_NM  ");
        //    strSql.AppendLine("      , '2000' AS RSLTN_DEPT_CD ");
        //    strSql.AppendLine("      , (SELECT X.DEPT_NM FROM ACC_DEPT_CD X WHERE X.DEPT_CD = '" + sExecDept + "') AS RSLTN_DEPT_NM ");
        //    strSql.AppendLine("      , A.EXEC_DEPT AS EXEC_DEPT_CD ");
        //    strSql.AppendLine("      , (SELECT X.DEPT_NM FROM ACC_DEPT_CD X WHERE X.DEPT_CD = A.EXEC_DEPT) AS EXEC_DEPT_NM ");
        //    strSql.AppendLine("      , CASE WHEN B.DBCR_GB = 'D' THEN A.SLIP_AMT ELSE 0 END DEBT_AMT ");
        //    strSql.AppendLine("      , CASE WHEN B.DBCR_GB = 'C' THEN A.SLIP_AMT ELSE 0 END CRDT_AMT ");
        //    strSql.AppendLine("      , B.RMK_CD ");
        //    strSql.AppendLine("      , A.NOTE AS ACC_RMK ");
        //    strSql.AppendLine("      , A.DEALER_CD AS DEALER_CD ");
        //    strSql.AppendLine("      , (SELECT X.IDT_NO FROM ACC_DEALER_CD X WHERE X.DEALER_CD = A.DEALER_CD) AS IDT_NO ");
        //    strSql.AppendLine("      , (SELECT X.DEALER_NM FROM ACC_DEALER_CD X WHERE X.DEALER_CD = A.DEALER_CD) AS DEALER_NM ");
        //    strSql.AppendLine("      , CASE WHEN B.DBCR_GB = 'D' THEN (SELECT X.DEBT_MGMT_GB1 FROM ACC_ACC_CD X WHERE X.ACC_CD = A.ACC_CD)  ");
        //    strSql.AppendLine("        ELSE (SELECT X.CRDT_MGMT_GB1 FROM ACC_ACC_CD X WHERE X.ACC_CD = A.ACC_CD) END AS MGMT_GB1  ");
        //    strSql.AppendLine("      , A.MGMT_CD1 ");
        //    strSql.AppendLine("      , 'N' AS APRV_YN ");
        //    strSql.AppendLine("      , NOW() AS ENT_DT ");
        //    strSql.AppendLine("      , '" + sid + "' AS ENT_ID ");
        //    strSql.AppendLine("      , NOW() AS MFY_DT ");
        //    strSql.AppendLine("      , '" + sid + "' AS MFY_ID   ");
        //    strSql.AppendLine("   FROM ACC_AUTO_SLIP A ");
        //    strSql.AppendLine("   LEFT OUTER JOIN ACC_SLIP_TYPE B ");
        //    strSql.AppendLine("     ON A.AUTO_SLIP_KIND = B.AUTO_SLIP_KIND ");
        //    strSql.AppendLine("    AND A.ACC_CD = B.ACC_CD  ");
        //    strSql.AppendLine("  WHERE A.AUTO_SLIP_KIND = '" + sSlipKind + "' ");
        //    strSql.AppendLine("    AND A.REQ_YMD = '" + sJLater + "' ");
        //    strSql.AppendLine("    AND A.PYBC_KEY = '" + sJid + "' ");

        //    cmd.CommandType = CommandType.Text;
        //    cmd.CommandText = strSql.ToString();
        //    cmd.ExecuteNonQuery();
        //    #endregion

        //    #region// ACC.ACC_PAY_BACK  테이블 추가

        //    //strSql.Clear();
        //    //strSql.AppendLine(" INSERT INTO ACC_PAY_BACK ");
        //    //strSql.AppendLine(" (PYBC_KEY, PYBC_SEQ, ACC_CD, SLIP_YMD, SLIP_NO, SLIP_SEQ, DEALER_CD, IDT_NO, DEALER_NM, PYBC_TRGT_AMT, PYBC_AMT, APRV_YN, NOTE, ENT_DT, ENT_ID, MFY_DT, MFY_ID) ");
        //    //strSql.AppendLine(" SELECT CONCAT(A.SLIP_YMD, A.SLIP_NO, RIGHT(CONCAT('0000', A.SLIP_SEQ), 4)) AS PYBC_KEY ");
        //    //strSql.AppendLine("      , 1 AS PYBC_SEQ ");
        //    //strSql.AppendLine("      , A.ACC_CD ");
        //    //strSql.AppendLine("      , A.SLIP_YMD ");
        //    //strSql.AppendLine("      , A.SLIP_NO ");
        //    //strSql.AppendLine("      , A.SLIP_SEQ ");
        //    //strSql.AppendLine("      , A.DEALER_CD ");
        //    //strSql.AppendLine("      , A.IDT_NO ");
        //    //strSql.AppendLine("      , A.DEALER_NM ");
        //    //strSql.AppendLine("      , A.CRDT_AMT + A.DEBT_AMT AS PYBC_TRGT_AMT ");
        //    //strSql.AppendLine("      , 0 AS PYBC_AMT ");
        //    //strSql.AppendLine("      , 'Y' AS APRV_YN ");
        //    //strSql.AppendLine("      , A.ACC_RMK AS NOTE ");
        //    //strSql.AppendLine("      , NOW() AS ENT_DT ");
        //    //strSql.AppendLine("      , '" + sid + "' AS ENT_ID ");
        //    //strSql.AppendLine("      , NOW() AS MFY_DT ");
        //    //strSql.AppendLine("      , '" + sid + "' AS MFY_ID ");
        //    //strSql.AppendLine("   FROM ACC_SLIP_DET A  ");
        //    //strSql.AppendLine("  INNER JOIN ACC_ACC_CD B  ");
        //    //strSql.AppendLine("     ON A.ACC_CD = B.ACC_CD  ");
        //    //strSql.AppendLine("    AND B.PYBC_YN = 'Y' ");
        //    //strSql.AppendLine("  WHERE A.SLIP_YMD = '" + sJLater + "' ");
        //    //strSql.AppendLine("    AND A.SLIP_NO = '" + sNewSlipNo + "' ");

        //    //cmd.CommandType = CommandType.Text;
        //    //cmd.CommandText = strSql.ToString();
        //    //cmd.ExecuteNonQuery();

        //    #endregion

        //    #region// ACC.ACC_SLIP_Mst  테이블 추가

        //    strSql.Clear();
        //    strSql.AppendLine(" INSERT INTO ACC_SLIP_MST ");
        //    strSql.AppendLine(" (SLIP_YMD, SLIP_NO, SLIP_KIND, RSLTN_DEPT_CD, RSLTN_DEPT_NM, SLIP_RMK, PMNT_REQ_YMD, PMNT_YMD, SLIP_AMT, APRV_YN, ENT_DT, ENT_ID, MFY_DT, MFY_ID) ");
        //    strSql.AppendLine("  ");
        //    strSql.AppendLine(" SELECT A.SLIP_YMD ");
        //    strSql.AppendLine("      , A.SLIP_NO ");
        //    strSql.AppendLine("      , '" + sSlipKind + "' AS SLIP_KIND  ");
        //    strSql.AppendLine("      , MAX(A.RSLTN_DEPT_CD) AS RSLTN_DEPT_CD ");
        //    strSql.AppendLine("      , MAX(A.RSLTN_DEPT_NM) AS RSLTN_DEPT_NM ");
        //    strSql.AppendLine("      , MAX(A.ACC_RMK) AS SLIP_RMK ");
        //    strSql.AppendLine("      , A.SLIP_YMD AS PMNT_REQ_YMD ");
        //    strSql.AppendLine("      , A.SLIP_YMD AS PMNT_YMD ");
        //    strSql.AppendLine("      , SUM(A.DEBT_AMT) AS SLIP_AMT ");
        //    strSql.AppendLine("      , 'N' AS APRV_YN ");
        //    strSql.AppendLine("      , NOW() AS ENT_DT ");
        //    strSql.AppendLine("      , '" + sid + "' AS ENT_ID ");
        //    strSql.AppendLine("      , NOW() AS MFY_DT ");
        //    strSql.AppendLine("      , '" + sid + "' AS MFY_ID   ");
        //    strSql.AppendLine("   FROM ACC_SLIP_DET A ");
        //    strSql.AppendLine("  WHERE A.SLIP_YMD = '" + sJLater + "' ");
        //    strSql.AppendLine("    AND A.SLIP_NO = '" + sNewSlipNo + "' ");
        //    strSql.AppendLine("  GROUP BY A.SLIP_YMD, A.SLIP_NO  ");

        //    cmd.CommandType = CommandType.Text;
        //    cmd.CommandText = strSql.ToString();
        //    cmd.ExecuteNonQuery();

        //    #endregion

        //    #region  // 세금계산서 Header

        //    string sIssueYy = sJLater.Substring(0, 4);

        //    strSql.Clear();
        //    //strSql.AppendLine(" SELECT SET_CRT_SEQ( 'BILL_SEQ', '" + sIssueYy + "' , '****', '****', '" + sid + "') ");
        //    strSql.AppendLine(" EXEC SET_CRT_SEQ 'BILL_SEQ', '" + sIssueYy + "' , '****', '****', '" + sid + "'");

        //    cmd.CommandType = CommandType.Text;
        //    cmd.CommandText = strSql.ToString();

        //    string sBillSeq = Convert.ToString(cmd.ExecuteScalar());
        //    string sPurSaleGb = sSlipKind.Equals("0002") ? "P" : "S";
        //    string sTaxGb = sSlipKind.Equals("0002") ? "21" : "11";
        //    string sCdGb = sSlipKind.Equals("0002") ? "KERA_ACC_CD_IN" : "KERA_ACC_CD_OUT";

        //    strSql.Clear();
        //    strSql.AppendLine(" INSERT INTO ACC_TAX_MGT ");
        //    strSql.AppendLine(" ( BILL_GB, PURCH_SALE_GB, ISSUE_YY, BILL_SEQ, BILL_ISSUE_YMD, JURIDICAL_GB, DEALER_CD, IDT_NO ");
        //    strSql.AppendLine(" , DEALER_NM, BIZ_NM, TYPE_NM, REP_NM, DEALER_ADDR, CHRG_NM, SUPL_AMT, VAT_AMT, TOT_AMT, TAX_GB ");
        //    strSql.AppendLine(" , NONDDCT_RESN, ASK_RECPT_GB, VAT_INCLUDE_YN, BILL_PMNT_GB, SLIP_ISSUE_YN, SLIP_YMD, SLIP_NO ");
        //    strSql.AppendLine(" , BILL_KEY, BILL_ISSUE_DEPT, DEALER_ISSUE_YN, NOTE, ENT_DT, ENT_ID, MFY_DT, MFY_ID)     ");
        //    strSql.AppendLine(" SELECT '1' AS BILL_GB ");

        //    strSql.AppendLine("      , '" + sPurSaleGb + "' AS PURCH_SALE_GB ");

        //    strSql.AppendLine("      , LEFT(A.SLIP_YMD, 4) AS ISSUE_YY ");
        //    strSql.AppendLine("      , '" + sBillSeq + "' AS BILL_SEQ ");
        //    strSql.AppendLine("      , '" + sJLater + "' AS BILL_ISSUE_YMD ");
        //    strSql.AppendLine("      , MAX(CASE WHEN LENGTH(B.IDT_NO) = 13 THEN 'S' ELSE 'C' END) AS JURIDICAL_GB ");
        //    strSql.AppendLine("      , MAX(A.DEALER_CD) AS DEALER_CD ");
        //    strSql.AppendLine("      , MAX(B.IDT_NO) AS IDT_NO ");
        //    strSql.AppendLine("      , MAX(B.DEALER_NM) AS DEALER_NM ");
        //    strSql.AppendLine("      , MAX(B.BIZ_NM) AS BIZ_NM ");
        //    strSql.AppendLine("      , MAX(B.TYPE_NM) AS TYPE_NM ");
        //    strSql.AppendLine("      , MAX(B.REP_NM) AS REP_NM ");
        //    strSql.AppendLine("      , MAX(CONCAT(B.ADDR, ' ', B.DTL_ADDR)) AS DEALER_ADDR ");
        //    strSql.AppendLine("      , MAX(B.CHRG_NM) AS CHRG_NM ");
        //    strSql.AppendLine("      , SUM(CASE WHEN A.ACC_CD = C.COM_SUB_CD1 THEN A.DEBT_AMT + A.CRDT_AMT ELSE 0 END) AS SUPL_AMT ");
        //    strSql.AppendLine("      , SUM(CASE WHEN A.ACC_CD = C.COM_SUB_CD2 THEN A.DEBT_AMT + A.CRDT_AMT ELSE 0 END) AS VAT_AMT ");
        //    strSql.AppendLine("      , SUM(CASE WHEN A.ACC_CD = C.COM_SUB_CD3 THEN A.DEBT_AMT + A.CRDT_AMT ELSE 0 END) AS TOT_AMT ");

        //    strSql.AppendLine("      , '" + sTaxGb + "' AS TAX_GB ");

        //    strSql.AppendLine("      , '' AS NONDDCT_RESN ");
        //    strSql.AppendLine("      , 'A' AS ASK_RECPT_GB ");
        //    strSql.AppendLine("      , 'N' AS VAT_INCLUDE_YN ");
        //    strSql.AppendLine("      , '4' AS BILL_PMNT_GB ");
        //    strSql.AppendLine("      , 'Y' AS SLIP_ISSUE_YN ");
        //    strSql.AppendLine("      , A.SLIP_YMD ");
        //    strSql.AppendLine("      , A.SLIP_NO ");
        //    strSql.AppendLine("      , CONCAT(LEFT(A.SLIP_YMD, 4), '" + sBillSeq + "') AS BILL_KEY ");
        //    strSql.AppendLine("      , '2000' AS BILL_ISSUE_DEPT ");
        //    strSql.AppendLine("      , 'N' AS DEALER_ISSUE_YN ");
        //    strSql.AppendLine("      , '매입/매출 세금계산서 임시 생성' AS NOTE ");
        //    strSql.AppendLine("      , NOW() AS ENT_DT ");
        //    strSql.AppendLine("      , '" + sid + "' AS ENT_ID ");
        //    strSql.AppendLine("      , NOW() AS MFY_DT ");
        //    strSql.AppendLine("      , '" + sid + "' AS MFY_ID ");
        //    strSql.AppendLine("   FROM ACC_SLIP_DET A  ");
        //    strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD B ");
        //    strSql.AppendLine("     ON A.DEALER_CD = B.DEALER_CD  ");
        //    strSql.AppendLine("  INNER JOIN COM_BASE_CD C ");

        //    strSql.AppendLine("     ON C.CD_GB = '" + sCdGb + "' ");

        //    strSql.AppendLine("    AND C.COM_CD = '" + sKeraType + "' ");
        //    strSql.AppendLine("  WHERE A.SLIP_YMD = '" + sJLater + "' ");
        //    strSql.AppendLine("    AND A.SLIP_NO = '" + sNewSlipNo + "' ");
        //    strSql.AppendLine("  GROUP BY A.SLIP_YMD, A.SLIP_NO  ");

        //    cmd.CommandType = CommandType.Text;
        //    cmd.CommandText = strSql.ToString();
        //    cmd.ExecuteNonQuery();

        //    #endregion

        //    #region  // 세금계산서 Detail ACC.ACC_TAX_DTL
        //    //string sMesId = sJLater + sNewSlipNo;

        //    //strSql.Clear();
        //    //strSql.AppendLine(" INSERT INTO ACC_TAX_DTL ");
        //    //strSql.AppendLine(" ( BILL_GB, PURCH_SALE_GB, ISSUE_YY, BILL_SEQ, DTL_SEQ, BILL_ISSUE_YMD, ITEM_NM, ITEM_SPEC ");
        //    //strSql.AppendLine(" , ITEM_AMOUNT, ITEM_UNPR, SUPL_AMT, VAT_AMT, ITEM_NOTE, ENT_DT, ENT_ID, MFY_DT, MFY_ID) ");
        //    //strSql.AppendLine("  SELECT '1' AS BILL_GB ");
        //    //strSql.AppendLine("       , '" + sPurSaleGb + "' AS PURCH_SALE_GB ");
        //    //strSql.AppendLine("       , '" + sIssueYy + "' AS ISSUE_YY ");
        //    //strSql.AppendLine("       , '" + sBillSeq + "' AS BILL_SEQ ");
        //    //strSql.AppendLine("       , 1 AS DTL_SEQ ");
        //    //strSql.AppendLine("       , '" + sJLater + "' AS BILL_ISSUE_YMD ");
        //    //strSql.AppendLine("       , B.GUBUN1 AS ITEM_NM ");
        //    //strSql.AppendLine("       , '' AS ITEM_SPEC ");

        //    //if(sPurSaleGb.Equals("P"))
        //    //{
        //    //    strSql.AppendLine("       , B.IWEIGHT AS ITEM_AMOUNT ");
        //    //    strSql.AppendLine("       , B.IDANGA AS ITEM_UNPR ");
        //    //    strSql.AppendLine("       , B.IKONGKEP AS SUPL_AMT ");
        //    //    strSql.AppendLine("       , B.IKONGKEP * 0.1 AS VAT_AMT ");
        //    //}
        //    //else
        //    //{
        //    //    strSql.AppendLine("       , B.OWEIGHT AS ITEM_AMOUNT ");
        //    //    strSql.AppendLine("       , B.ODANGA AS ITEM_UNPR ");
        //    //    strSql.AppendLine("       , B.OKONGKEP AS SUPL_AMT ");
        //    //    strSql.AppendLine("       , B.OKONGKEP * 0.1 AS VAT_AMT ");
        //    //}

        //    //strSql.AppendLine("       , B.GUMSUBIGO AS ITEM_NOTE ");
        //    //strSql.AppendLine("       , NOW() AS ENT_DT ");
        //    //strSql.AppendLine("       , '" + sid + "' AS ENT_ID ");
        //    //strSql.AppendLine("       , NOW() AS MFY_DT ");
        //    //strSql.AppendLine("       , '" + sid + "' AS MFY_ID ");
        //    //strSql.AppendLine("    FROM MESURING B ");
        //    //if (sPurSaleGb.Equals("P"))
        //    //{
        //    //    strSql.AppendLine("   WHERE B.IPCHULGO_MAIPID = '" + sMesId + "' ");
        //    //}
        //    //else
        //    //{
        //    //    strSql.AppendLine("   WHERE B.IPCHULGO_MACHULID = '" + sMesId + "' ");
        //    //}

        //    //cmd.CommandType = CommandType.Text;
        //    //cmd.CommandText = strSql.ToString();
        //    //cmd.ExecuteNonQuery();

        //    #endregion

        //    #region  // 증빙내역 ACC.ACC_SLIP_EVDN

        //    //strSql.Clear();
        //    //strSql.AppendLine(" INSERT INTO ACC_SLIP_EVDN ");
        //    //strSql.AppendLine(" ( SLIP_YMD, SLIP_NO, EVDN_SEQ, EVDN_GB, EVDN_YMD, EVDN_AMT, SUPL_AMT, VAT_AMT ");
        //    //strSql.AppendLine(" , DEALER_CD, IDT_NO, DEALER_NM, EVDN_KEY, NOTE, ENT_DT, ENT_ID, MFY_DT, MFY_ID) ");
        //    //strSql.AppendLine(" SELECT A.SLIP_YMD  ");
        //    //strSql.AppendLine("      , A.SLIP_NO ");
        //    //strSql.AppendLine("      , 1 AS EVDN_SEQ ");
        //    //strSql.AppendLine("      , '01' AS EVDN_GB ");
        //    //strSql.AppendLine("      , A.BILL_ISSUE_YMD AS EVDN_YMD ");
        //    //strSql.AppendLine("      , A.TOT_AMT AS EVDN_AMT ");
        //    //strSql.AppendLine("      , A.SUPL_AMT ");
        //    //strSql.AppendLine("      , A.VAT_AMT ");
        //    //strSql.AppendLine("      , A.DEALER_CD ");
        //    //strSql.AppendLine("      , A.IDT_NO ");
        //    //strSql.AppendLine("      , A.DEALER_NM ");
        //    //strSql.AppendLine("      , CONCAT(A.ISSUE_YY, A.BILL_SEQ) AS EVDN_KEY ");
        //    //strSql.AppendLine("      , A.NOTE AS NOTE ");
        //    //strSql.AppendLine("      , NOW() AS ENT_DT ");
        //    //strSql.AppendLine("      , '" + sid + "' AS ENT_ID ");
        //    //strSql.AppendLine("      , NOW() AS MFY_DT ");
        //    //strSql.AppendLine("      , '" + sid + "' AS MFY_ID ");
        //    //strSql.AppendLine("   FROM ACC_TAX_MGT A ");
        //    //strSql.AppendLine("  WHERE ISSUE_YY = '" + sIssueYy + "' ");
        //    //strSql.AppendLine("    AND BILL_SEQ = '" + sBillSeq + "' ");

        //    //cmd.CommandType = CommandType.Text;
        //    //cmd.CommandText = strSql.ToString();
        //    //cmd.ExecuteNonQuery();

        //    #endregion
        //}

        int unCheckedCount;

        private void GridViewRetr_CustomSummaryCalculate(object sender, CustomSummaryEventArgs e)
        {
            GridView view = sender as GridView;

            int summaryID = Convert.ToInt32((e.Item as GridSummaryItem).Tag);

            if (e.SummaryProcess == CustomSummaryProcess.Start)
            {
                unCheckedCount = 0;
            }

            if (e.SummaryProcess == CustomSummaryProcess.Calculate)
            {
                if (e.FieldValue.Equals("0")) unCheckedCount++;
            }

            if (e.SummaryProcess == CustomSummaryProcess.Finalize)
            {
                e.TotalValue = unCheckedCount;
            }
        }

        private void RdGrClose_SelectedIndexChanged(object sender, EventArgs e)
        {
            int iCloseGbIdx = RdGrClose.SelectedIndex;

            if (iCloseGbIdx == 1)
            {
                BtnCancle.Enabled = true;
                BtnSave.Enabled = false;
            }
            else
            {
                BtnCancle.Enabled = false;
                BtnSave.Enabled = true;
            }

            BtnRetr_Click(null, null);
        }

        //private void GridViewRetr_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        //{
        //    if (e.Column.FieldName.Equals("IDANGA") || e.Column.FieldName.Equals("ODANGA"))
        //    {
        //        DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
        //        SqlCommand cmd = DBConn.dbCon.CreateCommand();
        //        cmd.Transaction = DBConn.dbTran;

        //        string sJunpyoid = GridViewRetr.GetRowCellValue(e.RowHandle, "JUNPYOID").ToString();
        //        double dIDanGa = Convert.ToDouble(GridViewRetr.GetRowCellValue(e.RowHandle, "IDANGA"));
        //        double dODanGa = Convert.ToDouble(GridViewRetr.GetRowCellValue(e.RowHandle, "ODANGA"));
        //        StringBuilder strSql = new StringBuilder();

        //        try
        //        {
        //            strSql = new StringBuilder();

        //            strSql.AppendLine(" UPDATE ESURING ");
        //            strSql.AppendLine("    SET IDANGA = " + dIDanGa);
        //            strSql.AppendLine("      , ODANGA = " + dODanGa);
        //            strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "'");

        //            cmd.CommandType = CommandType.Text;
        //            cmd.CommandText = strSql.ToString();
        //            cmd.ExecuteNonQuery();

        //            DBConn.dbTran.Commit();
        //            DBConn.dbTran = null;
        //            MessageBox.Show("저장을 완료했습니다.");
        //        }
        //        catch (Exception ex)
        //        {
        //            DBConn.dbTran.Rollback();
        //            DBConn.dbTran = null;
        //            MessageBox.Show(ex.Message);
        //        }
        //    }
        //}

        private void GridViewRetr_RowUpdated(object sender, DevExpress.XtraGrid.Views.Base.RowObjectEventArgs e)
        {

        }

        private void GridViewRetr_ShowingPopupEditForm(object sender, ShowingPopupEditFormEventArgs e)
        {
            e.EditForm.ImeMode = ImeMode.Hangul;
            e.EditForm.StartPosition = FormStartPosition.CenterParent;

            foreach (var button in e.EditForm.Controls.OfType<EditFormContainer>()
                        .SelectMany(control => Enumerable.Cast<Control>(control.Controls)).OfType<PanelControl>()
                        .SelectMany(nestedControl => Enumerable.Cast<Control>(nestedControl.Controls)).OfType<SimpleButton>())
            {
                switch (button.Text)
                {
                    case "Update":
                        //button.Visible = false;
                        button.Text = "수정";
                        button.Click += EditFormUpdateButton_Click;
                        break;

                    case "Cancel":
                        button.Text = "취소";
                        break;
                }
            }
        }

        private void EditFormUpdateButton_Click(object sender, EventArgs e)
        {
            //DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            //SqlCommand cmd = DBConn.dbCon.CreateCommand();
            //cmd.Transaction = DBConn.dbTran;

            //string sJunpyoid = GridViewRetr.GetRowCellValue(iUpdtRowIdx, "JUNPYOID").ToString();
            //double dIDanGa = Convert.ToDouble(GridViewRetr.GetRowCellValue(iUpdtRowIdx, "IDANGA"));
            //double dODanGa = Convert.ToDouble(GridViewRetr.GetRowCellValue(iUpdtRowIdx, "ODANGA"));
            //StringBuilder strSql = new StringBuilder();

            //try
            //{
            //    strSql = new StringBuilder();

            //    strSql.AppendLine(" UPDATE ESURING ");
            //    strSql.AppendLine("    SET IDANGA = " + dIDanGa);
            //    strSql.AppendLine("      , ODANGA = " + dODanGa);
            //    strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "'");

            //    cmd.CommandType = CommandType.Text;
            //    cmd.CommandText = strSql.ToString();
            //    cmd.ExecuteNonQuery();

            //    DBConn.dbTran.Commit();
            //    DBConn.dbTran = null;
            //    MessageBox.Show("저장을 완료했습니다.");
            //}
            //catch (Exception ex)
            //{
            //    DBConn.dbTran.Rollback();
            //    DBConn.dbTran = null;
            //    MessageBox.Show(ex.Message);
            //}

            string sJunpyoid = GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "JUNPYOID").ToString();
            string sKeraType = GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "GB").ToString();
            double dIDanGa = Convert.ToDouble(GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "IDANGA"));
            double dODanGa = Convert.ToDouble(GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "ODANGA"));
            double dAmount = Convert.ToDouble(GridViewRetr.GetRowCellValue(GridViewRetr.FocusedRowHandle, "AMOUNT"));

            double dIKongKep = 0;
            double dOKongKep = 0;

            if (!sKeraType.Equals("출고"))  // 매입금액 수정
            {
                dIKongKep = dIDanGa * dAmount;
            }

            if (!sKeraType.Equals("입고")) // 매출금액 수정
            {
                dOKongKep = dODanGa * dAmount;
            }

            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            StringBuilder strSql = new StringBuilder();

            try
            {
                strSql = new StringBuilder();

                strSql.AppendLine(" UPDATE MESURING ");
                strSql.AppendLine("    SET IDANGA = " + dIDanGa);
                strSql.AppendLine("      , ODANGA = " + dODanGa);
                strSql.AppendLine("      , IKONGKEP = " + dIKongKep);
                strSql.AppendLine("      , OKONGKEP = " + dOKongKep);
                strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoid + "'");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }
        
        private void ChkComboGb_EditValueChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void AccMeasureCloseDev_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                //BtnCrete_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {
                BtnModify_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnFile_Click(null, null);
            }
            else if (e.KeyCode == Keys.F12)
            {
                BtnExcel_Click(null, null);
            }
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (e.FocusedRowHandle >= 0)
            {
                bReCloseYn = false;

                string sGb = GridViewRetr.GetFocusedRowCellValue("GB")?.ToString();
                if (sGb.Equals("직송"))
                {
                    GridColMaipcher.Visible = true;
                    GridColMaipcher.VisibleIndex = 7;
                }
                else
                {
                    GridColMaipcher.Visible = false;
                }

                string sCloseYN = GridViewRetr.GetFocusedRowCellValue("CLO")?.ToString();
                string sInOutGb = GridViewRetr.GetFocusedRowCellValue("GB")?.ToString();

                if (clickBtnModifyTF)
                {
                    if (GridViewRetr.IsRowSelected(e.FocusedRowHandle))
                    {
                        SetControlsUnReadOnly();
                        BtnSave.Enabled = false;
                    }
                    else
                    {
                        SetControlsReadOnly();
                    }
                }
                else
                {
                    if (sCloseYN.Equals("1"))   //마감 시 전체 컨트롤 READ ONLY
                    {
                        SetControlsReadOnly();
                        BtnSave.Enabled = false;
                    }
                    else
                    {
                        /*
                         * 2021-03-12
                         * 현업요청
                         * #00003
                         * 출고일 경우 마감및업로드에서 감량중량을 조절하므로 해당 Control은 별도로 관리
                         */
                        if (sGb.Equals("출고"))
                        {
                            TxtLooseWeight.ReadOnly = false;
                        }
                        else
                        {
                            TxtLooseWeight.ReadOnly = true;
                        }
                        SetControlsUnReadOnly();
                        BtnSave.Enabled = true;
                    }
                }
                clickBtnModifyTF = false;
                BtnModify.Text = "재마감(F4)";
                #region[계근상세내역]

                //운행정보
                string sCarNo = string.Empty;
                string sArrvDstn = string.Empty;
                string sDepatureDstn = string.Empty;
                string sCarComp = string.Empty;
                string sCarryCostRemark = string.Empty;

                sCarNo = GridViewRetr.GetFocusedRowCellValue("CARNO")?.ToString();
                TxtCarNo.EditValue = sCarNo;

                //계근정보
                string sTotalWeight = string.Empty;
                string sMesuringTime1 = string.Empty;
                string sEmptyWeight = string.Empty;
                string sMesuringTime2 = string.Empty;
                string sLandedWeight = string.Empty;
                string sLooseWeight = string.Empty;
                string sCompWeight = string.Empty;
                string sLossWeight = string.Empty;
                string sCarryCostPmnt = string.Empty;
                string sInUnitPrc = string.Empty;
                string sOutUnitPrc = string.Empty;

                sTotalWeight = GridViewRetr.GetFocusedRowCellValue("TOT")?.ToString();
                sMesuringTime1 = GridViewRetr.GetFocusedRowCellValue("TIME2")?.ToString();
                sEmptyWeight = GridViewRetr.GetFocusedRowCellValue("EMPAMT")?.ToString();
                sMesuringTime2 = GridViewRetr.GetFocusedRowCellValue("TIME1")?.ToString();
                sLandedWeight = GridViewRetr.GetFocusedRowCellValue("OWN")?.ToString();
                sLooseWeight = GridViewRetr.GetFocusedRowCellValue("REDUCE")?.ToString();
                sCompWeight = GridViewRetr.GetFocusedRowCellValue("CUSTOM")?.ToString();
                sLossWeight = GridViewRetr.GetFocusedRowCellValue("LOSS")?.ToString();
                sCarryCostPmnt = GridViewRetr.GetFocusedRowCellValue("COST")?.ToString();
                sInUnitPrc = GridViewRetr.GetFocusedRowCellValue("IDANGA")?.ToString();
                sOutUnitPrc = GridViewRetr.GetFocusedRowCellValue("ODANGA")?.ToString();

                TxtTotWeight.EditValue = sTotalWeight;
                TimeEditTot.EditValue = sMesuringTime1;
                TxtEmptyWeight.EditValue = sEmptyWeight;
                TimeEditEmpty.EditValue = sMesuringTime2;
                TxtLandedWeight.EditValue = sLandedWeight;
                TxtLooseWeight.EditValue = sLooseWeight;
                TxtCompWeight.EditValue = sCompWeight;
                TxtLossWeight.EditValue = sLossWeight;
                TxtPmntCarryCost.EditValue = sCarryCostPmnt;
                TxtInUnitPrc.EditValue = sInUnitPrc;
                TxtOutPrc.EditValue = sOutUnitPrc;

                ChkVatPoham.EditValue = GridViewRetr.GetFocusedRowCellValue(GridColVatYn);

                string sKeraType = GridViewRetr.GetFocusedRowCellValue("GB")?.ToString();
                
                if (sKeraType.Equals("직송"))
                {
                    LayoutGroupDirectSend.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                    //직송매출정보
                    string sSaleDealerCd = string.Empty;
                    string sSaleDealerNm = string.Empty;
                    string sSaleLossWeight = string.Empty;
                    string sSaleWeight = string.Empty;
                    string sSaleUnitPrc = string.Empty;
                    string sSaleAmt = string.Empty;
                    string sCashUnitPrcProfit = string.Empty;

                    sSaleDealerCd = GridViewRetr.GetFocusedRowCellValue("PURCHASER_CD")?.ToString();
                    sSaleDealerNm = GridViewRetr.GetFocusedRowCellValue("PURCHASER")?.ToString();
                    sSaleLossWeight = GridViewRetr.GetFocusedRowCellValue("OUT_LOOSE_WEIGHT")?.ToString();
                    sSaleWeight = GridViewRetr.GetFocusedRowCellValue("OUT_WEIGHT")?.ToString();
                    sSaleUnitPrc = GridViewRetr.GetFocusedRowCellValue("ODANGA")?.ToString();
                    sSaleAmt = GridViewRetr.GetFocusedRowCellValue("OUT_AMT")?.ToString();

                    TxtSaleDealerCd.EditValue = sSaleDealerCd;
                    TxtSaleDealerNM.EditValue = sSaleDealerNm;
                    TxtSaleLoss.EditValue = sSaleLossWeight;
                    TxtSaleWeight.EditValue = sSaleWeight;
                    TxtSaleUnitPrc.EditValue = sSaleUnitPrc;
                    TxtSaleAmt.EditValue = sSaleAmt;
                }
                else
                {
                    LayoutGroupDirectSend.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                }

                //당사, 업체중량 구분
                string sOurTheirWeightGb = GridViewRetr.GetFocusedRowCellValue("WEIGHT_GUBUN")?.ToString();
                RdgbMesuringGB.EditValue = sOurTheirWeightGb;

                //기타비용
                string sJunpyoid = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" SELECT B.DEALER_NM AS ETC_DEALER_NM ");
                strSql.AppendLine(" 	 , A.ETC_REMARK1 AS ETC_REMARK     ");
                strSql.AppendLine(" 	 , A.ETC_COST1 AS ETC_COST         ");
                strSql.AppendLine("   FROM MESURING A                       ");
                strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD B     ");
                strSql.AppendLine("  	ON A.ETC_DEALER_CD1 = CONVERT(VARCHAR,B.DEALER_CD)");
                strSql.AppendLine("  WHERE A.JUNPYOID =  " + sJunpyoid + "           ");
                strSql.AppendLine("  UNION ALL                             ");
                strSql.AppendLine(" SELECT B.DEALER_NM AS ETC_DEALER_NM");
                strSql.AppendLine(" 	 , A.ETC_REMARK2 AS ETC_REMARK     ");
                strSql.AppendLine(" 	 , A.ETC_COST2 AS ETC_COST         ");
                strSql.AppendLine("   FROM MESURING A                       ");
                strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD B     ");
                strSql.AppendLine("  	ON A.ETC_DEALER_CD2 = CONVERT(VARCHAR,B.DEALER_CD)");
                strSql.AppendLine("  WHERE A.JUNPYOID =  " + sJunpyoid + "           ");

                DataTable dtEtc = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                GridEtcCost.DataSource = dtEtc;

                #endregion[계근상세내역]
            }
        }

        private bool bReCloseYn = false;
        private void BtnModify_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (clickBtnModifyTF)
            {
                BtnSave_Click(null, null);
                clickBtnModifyTF = false;
                BtnModify.Text = "재마감(F4)";
                return;
            }

            if(GridViewRetr.RowCount == 0)
            {
                XtraMessageBox.Show("리스트에 입출고 건이 존재하지 않습니다.");
                return;
            }

            string sCloseYN = GridViewRetr.GetFocusedRowCellValue("CLO")?.ToString();
            string sJ_ID = GridViewRetr.GetFocusedRowCellValue("J_ID")?.ToString();
            string sJunpyoId = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();
            if (sCloseYN.Equals("1"))
            {
                if (XtraMessageBox.Show("재마감을 하시겠습니까?", "재마감 진행 여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    clickBtnModifyTF = false;
                    return;
                }
                else
                {
                    bReCloseYn = true;
                    if (!GridViewRetr.GetFocusedRowCellValue("GB").ToString().Equals("직송"))
                    {
                        clickBtnModifyTF = true;

                        //string sSlipYmd = sJ_ID.Substring(0, 8);
                        //string sSlipNo = sJ_ID.Substring(8, 7);

                        StringBuilder strSql = new StringBuilder();

                        //INLIST 체크
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" SELECT JUNPYOID ");
                        strSql.AppendLine("   FROM INLIST A ");
                        strSql.AppendLine("  WHERE A.J_ID = ( SELECT CASE WHEN KERATYPE = '입고' THEN IPCHULGO_MAIPID ");
                        strSql.AppendLine("                   	   		  WHEN KERATYPE = '출고' THEN IPCHULGO_MACHULID END");
                        strSql.AppendLine("                     FROM MESURING ");
                        strSql.AppendLine("                    WHERE JUNPYOID = '" + sJunpyoId + "')");

                        DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                        if (dtChk.Rows.Count > 0)
                        {
                            string sInlistJunpyoId = dtChk.Rows[0]["JUNPYOID"].ToString();

                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" SELECT COUNT(*) AS CNT ");
                            strSql.AppendLine("   FROM ACTRAN A ");
                            strSql.AppendLine("  WHERE A.REF1 = '" + sInlistJunpyoId + "' ");

                            dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                            int iCnt = Convert.ToInt32(dtChk.Rows[0]["CNT"]);
                            if (iCnt > 0)
                            {
                                Cursor = Cursors.Default;
                                XtraMessageBox.Show("해당 건은 전표발행 상태입니다.\r\n전표취소 후 다시 시도하세요.");
                                return;
                            }
                        }

                        ////INLIST 체크
                        //strSql.Clear();
                        //strSql.AppendLine(" SELECT CASE WHEN B.JUNPYOID IS NOT NULL THEN B.JUNPYOID ELSE NULL END AS JUNPYOID ");
                        //strSql.AppendLine("      , COUNT(1) AS CNT ");
                        //strSql.AppendLine("   FROM MESURING A ");
                        //strSql.AppendLine("   LEFT OUTER JOIN INLIST B ");
                        //strSql.AppendLine("     ON A.JunpyoID = B.J_RID ");
                        //strSql.AppendLine("  WHERE A.JunpyoID = " + sJunpyoId + " ");
                        
                        //DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                        //if(dt.Rows.Count > 0)
                        //{
                        //    //전표 개수 체크
                        //    double dCnt = Convert.ToDouble(dt.Rows[0]["CNT"]);
                        //    string sInlistJunpyoId = string.Empty;
                        //    if (dCnt > 0)
                        //    {
                        //        sInlistJunpyoId = dt.Rows[0]["JUNPYOID"]?.ToString();

                        //        strSql.Clear();
                        //        strSql.AppendLine(" SELECT COUNT(1) AS CNT ");
                        //        strSql.AppendLine("   FROM ACTRAN A ");
                        //        strSql.AppendLine("  WHERE REF1 = '" + sInlistJunpyoId + "' ");

                        //        dt.Clear();
                        //        dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                        //        dCnt = 0;
                        //        dCnt = Convert.ToDouble(dt.Rows[0]["CNT"]);
                        //        if(dCnt > 0)
                        //        {
                        //            XtraMessageBox.Show("해당 마감 건은 전표발행 상태입니다.\r\n 승인취소 후 다시 마감해주세요.");
                        //            return;
                        //        }
                        //    }
                        //}

                        SetControlsUnReadOnly();
                        if (GridViewRetr.GetFocusedRowCellValue("GB").ToString().Equals("출고"))
                        {
                            TxtLooseWeight.ReadOnly = false;
                        }
                        else
                        {
                            TxtLooseWeight.ReadOnly = true;
                        }
                        BtnSave.Enabled = false;
                        BtnModify.Text = "완료";
                    }
                    else
                    {
                        BtnSave_Click(null, null);
                    }
                }
            }
            else
            {
                SetControlsUnReadOnly();
            }
        }

        private void BtnCancle_Click(object sender, EventArgs e)
        {
            string sKeraType = GridViewRetr.GetFocusedRowCellValue(GridColGb)?.ToString();
            string sDealerNm = GridViewRetr.GetFocusedRowCellValue(GridColDealer)?.ToString();
            string sGrade = GridViewRetr.GetFocusedRowCellValue(GridColGrade)?.ToString();
            string sAmount = GridViewRetr.GetFocusedRowCellValue(GridColAmont)?.ToString();
            string sJunpyoId = GridViewRetr.GetFocusedRowCellValue(GridColJunPyoId)?.ToString();
            string sDate = GridViewRetr.GetFocusedRowCellValue(GridColYmd)?.ToString();

            if (string.IsNullOrEmpty(sJunpyoId))
            {
                XtraMessageBox.Show("마감취소 할 항목을 선택하세요.");
            }

            if (sKeraType.Equals("직송"))
            {
                XtraMessageBox.Show("직송은 취소할 수 없습니다.");
                return;
            }

            string sMsg = string.Format("거래구분 : {0}\r\n거래처명 : {1}\r\n등급 : {2}\r\n인수량 : {3}\r\n해당 건에 대하여 마감취소를 진행하시겠습니까?", sKeraType, sDealerNm, sGrade, sAmount, sDate);
            if (XtraMessageBox.Show(sMsg, "마감취소여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                bool bYn = true;
                StringBuilder strSql = new StringBuilder();

                //INLIST 체크
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT JUNPYOID ");
                strSql.AppendLine("   FROM INLIST A ");
                strSql.AppendLine("  WHERE A.J_ID = ( SELECT CASE WHEN KERATYPE = '입고' THEN IPCHULGO_MAIPID ");
                strSql.AppendLine("                   	   		  WHEN KERATYPE = '출고' THEN IPCHULGO_MACHULID END");
                strSql.AppendLine("                     FROM MESURING ");
                strSql.AppendLine("                    WHERE JUNPYOID = '" + sJunpyoId + "')");
                
                DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                if(dtChk.Rows.Count > 0)
                {
                    string sInlistJunpyoId = dtChk.Rows[0]["JUNPYOID"].ToString();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT COUNT(*) AS CNT ");
                    strSql.AppendLine("   FROM ACTRAN A ");
                    strSql.AppendLine("  WHERE A.REF1 = '" + sInlistJunpyoId + "' " );

                    dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    int iCnt = Convert.ToInt32(dtChk.Rows[0]["CNT"]);
                    if(iCnt > 0)
                    {
                        Cursor = Cursors.Default;
                        XtraMessageBox.Show("해당 건은 전표발행 상태입니다.\r\n전표취소 후 다시 시도하세요.");
                        return;
                    }
                    else if (iCnt == 0)
                    {
                        bYn = false;
                    }
                }
                else //INLIST에 없을 경우(마이그로 인해 J_CHECK 가 1인 데이터이면서 INLIST에 존재하지 않는 데이터의 경우) 그대로 DELETE 로직 처리
                {
                    bYn = false;
                }

                if (bYn)
                    return;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT CASE WHEN KERATYPE = '입고' THEN IPCHULGO_MAIPID ELSE IPCHULGO_MACHULID END AS J_ID ");
                strSql.AppendLine("   FROM MESURING ");
                strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoId + " ");

                DataTable dtJID = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                string sJ_ID = string.Empty;
                if(dtJID.Rows.Count > 0)
                {
                    sJ_ID = dtJID.Rows[0]["J_ID"]?.ToString();

                    //IPCHULGO 삭제처리
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" DELETE FROM IPCHULGO ");
                    strSql.AppendLine("       WHERE J_ID = " + sJ_ID + " ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }
                
                //INLIST 삭제처리
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" DELETE FROM INLIST ");
                strSql.AppendLine("       WHERE J_RID = " + sJunpyoId + " ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();
                
                //MESURING 미마감 처리
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" UPDATE MESURING ");
                strSql.AppendLine("    SET J_CHECK = '' ");
                strSql.AppendLine("      , AGREE_DATE = NULL ");
                strSql.AppendLine("      , IPCHULGO_MAIPID = 0 ");
                strSql.AppendLine("      , IPCHULGO_MACHULID = 0 ");
                strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoId + " ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                /*
                 * 2021-03-09
                 * (현업요청)
                 * LOG 적용에 따라 로직추가
                 * Reference : #00001
                 */
                DataTable dtPrv = null;
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT JUNPYOID ");
                strSql.AppendLine("      , SUN ");
                strSql.AppendLine("      , J_DATE ");
                strSql.AppendLine("      , J_BNUM ");
                strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN MAIPCHER ELSE J_COMPANY END AS DEALER ");
                strSql.AppendLine("      , CASE WHEN KERATYPE = '입고' THEN IWEIGHT ELSE OWEIGHT END AS WEIGHT ");
                strSql.AppendLine("   FROM MESURING ");
                strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoId + " ");

                dtPrv = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                strSql.Clear();

                /*
                 * #00006
                 */
                string sSTD_COLS = string.Empty;
                string sEDIT_RMK = string.Empty;
                string sREF_RMK = string.Empty;
                string sLog_Msg = string.Format("[마감취소]");
                if (dtPrv.Rows.Count > 0)
                {
                    string sJUNPYOID = dtPrv.Rows[0]["JUNPYOID"]?.ToString();
                    string sJ_DATE = dtPrv.Rows[0]["J_DATE"]?.ToString().Substring(0, 10);
                    string sSUN = dtPrv.Rows[0]["SUN"]?.ToString();
                    string sDEALER = dtPrv.Rows[0]["DEALER"]?.ToString();
                    string sJ_BNUM = dtPrv.Rows[0]["J_BNUM"]?.ToString();
                    string sWEIGHT = dtPrv.Rows[0]["WEIGHT"]?.ToString();

                    sSTD_COLS += string.Format("{0}/순번:{1}/{2}/차번:{3}/인수량:{4}"
                        , sJ_DATE
                        , sSUN
                        , sDEALER
                        , sJ_BNUM
                        , sWEIGHT );

                    sREF_RMK += string.Format("Table : Mesuring, JunpyoID : {0}", sJUNPYOID);
                }

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
                cmd.Parameters.AddWithValue("@EDIT_RMK", sLog_Msg);
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                Cursor = Cursors.Default;
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("마감취소를 완료했습니다.");

                int idx = GridViewRetr.FocusedRowHandle;

                BtnRetr.PerformClick();

                GridViewRetr.FocusedRowHandle = idx - 1;
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        #region[컨트롤 ReadOnly 설정]

        private void SetControlsReadOnly()
        {
            TxtCarNo.ReadOnly = true;
            TxtArrv.ReadOnly = true;
            TxtDepature.ReadOnly = true;
            TxtCarComp.ReadOnly = true;
            TxtCrryCostNote.ReadOnly = true;

            TxtTotWeight.ReadOnly = true;
            TimeEditTot.ReadOnly = true;
            TxtEmptyWeight.ReadOnly = true;
            TimeEditEmpty.ReadOnly = true;
            TxtLandedWeight.ReadOnly = true;
            TxtLooseWeight.ReadOnly = true;
            TxtCompWeight.ReadOnly = true;
            TxtLossWeight.ReadOnly = true;
            RdgbMesuringGB.ReadOnly = true;

            TxtSaleDealerCd.ReadOnly = true;
            TxtSaleDealerNM.ReadOnly = true;
            TxtSaleLoss.ReadOnly = true;
            TxtSaleWeight.ReadOnly = true;
            TxtSaleUnitPrc.ReadOnly = true;
            TxtSaleAmt.ReadOnly = true;
            TxtCashUntPrcPrft.ReadOnly = true;

            ChkVatPoham.ReadOnly = true;
        }

        private void SetControlsUnReadOnly()
        {
            TxtCarNo.ReadOnly = false;
            TxtArrv.ReadOnly = false;
            TxtDepature.ReadOnly = false;
            TxtCarComp.ReadOnly = false;
            TxtCrryCostNote.ReadOnly = false;

            //총중량, 계근시각, 공차중량은 계근마감에서 입력하므로 ReadOnly 해제에서 제외시킴
            //TxtTotWeight.ReadOnly = false;
            //TimeEditTot.ReadOnly = false;
            //TxtEmptyWeight.ReadOnly = false;
            //TimeEditEmpty.ReadOnly = false;
            TxtLandedWeight.ReadOnly = false;

            /*
             * 2021-03-11
             * Reference Key : #00002
             * 감량중량은 계근프로그램에서 입력하므로 Readonly 해제에서 제외시킴
             * 
             */
            //TxtLooseWeight.ReadOnly = false;

            TxtCompWeight.ReadOnly = false;
            TxtLossWeight.ReadOnly = false;
            RdgbMesuringGB.ReadOnly = false;

            TxtSaleDealerCd.ReadOnly = false;
            TxtSaleDealerNM.ReadOnly = false;
            TxtSaleLoss.ReadOnly = false;
            TxtSaleWeight.ReadOnly = false;
            TxtSaleUnitPrc.ReadOnly = false;
            TxtSaleAmt.ReadOnly = false;
            TxtCashUntPrcPrft.ReadOnly = false;

            ChkVatPoham.ReadOnly = false;
        }

        #endregion[컨트롤 ReadOnly 설정]

        #region[값 입력 시 해당 ROW에 입력 이벤트]
        private void TxtInUnitPrc_EditValueChanged(object sender, EventArgs e)
        {
            string sInUnitPrc = TxtInUnitPrc.EditValue?.ToString();
            GridViewRetr.SetFocusedRowCellValue("IDANGA", sInUnitPrc);
        }

        private void TxtOutPrc_EditValueChanged(object sender, EventArgs e)
        {
            string sOutUnitPrc = TxtOutPrc.EditValue?.ToString();
            GridViewRetr.SetFocusedRowCellValue("ODANGA", sOutUnitPrc);
        }

        private void TxtPmntCarryCost_EditValueChanged(object sender, EventArgs e)
        {
            string sCarryCost = TxtPmntCarryCost.EditValue?.ToString();
            GridViewRetr.SetFocusedRowCellValue("COST", sCarryCost);
        }

        private void RdgbMesuringGB_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sWeightGb = RdgbMesuringGB.EditValue?.ToString();
            GridViewRetr.SetFocusedRowCellValue("WEIGHT_GUBUN", sWeightGb);
        }

        private void TxtLandedWeight_EditValueChanged(object sender, EventArgs e)
        {
            string sLandedWeight = TxtLandedWeight.EditValue?.ToString();
            GridViewRetr.SetFocusedRowCellValue("OWN", sLandedWeight);
        }

        private void TxtLooseWeight_EditValueChanged(object sender, EventArgs e)
        {
            string sLooseWeight = TxtLooseWeight.EditValue?.ToString();
            GridViewRetr.SetFocusedRowCellValue("REDUCE", sLooseWeight);
        }

        private void TxtCompWeight_EditValueChanged(object sender, EventArgs e)
        {
            string sCompWeight = TxtCompWeight.EditValue?.ToString();
            GridViewRetr.SetFocusedRowCellValue("CUSTOM", sCompWeight);
        }

        private void TxtLossWeight_EditValueChanged(object sender, EventArgs e)
        {
            string sLossWeight = TxtLossWeight.EditValue?.ToString();
            GridViewRetr.SetFocusedRowCellValue("LOSS", sLossWeight);
        }
        #endregion[값 입력 시 해당 ROW에 입력 이벤트]

        string _sFileName;
        private void BtnFile_Click(object sender, EventArgs e)
        {
            clickBtnModifyTF = false;
            BtnModify.Text = "재마감(F4)";

            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 EXCEL 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }
            
            try
            {
                OpenFileDialog fileDlg = new OpenFileDialog();
                
                string filePath = Application.StartupPath + @"\VersionCheck.ini";
                iniUtil ini = new iniUtil(filePath);
                string sInitPath = ini.GetIniValue("DIRECT_SEND_FILE_PATH", "path").Replace(" ", "");

                if (string.IsNullOrEmpty(sInitPath))
                    fileDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                else
                    fileDlg.InitialDirectory = sInitPath;

                fileDlg.Filter = "Excel Files (.xlsx)|*.xlsx|All Files (*.*)|*.*";
                fileDlg.FilterIndex = 1;
                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    ini.SetIniValue("DIRECT_SEND_FILE_PATH", "path", fileDlg.FileName.Replace(fileDlg.SafeFileName, ""));
                    _sFileName = fileDlg.FileName;
                }
                fileDlg.Dispose();

                if (!string.IsNullOrEmpty(_sFileName))
                {
                    DevExpress.DataAccess.Excel.ExcelDataSource excelDataSource = new DevExpress.DataAccess.Excel.ExcelDataSource
                    {
                        FileName = _sFileName
                    };

                    ExcelWorksheetSettings workSheetSettings = new ExcelWorksheetSettings("sheet1");
                    excelDataSource.SourceOptions = new ExcelSourceOptions(workSheetSettings)
                    {
                        SkipEmptyRows = true,
                        UseFirstRowAsHeader = true
                    };

                    excelDataSource.Fill();
                    GridFile.DataSource = excelDataSource;

                    UploadData(GridViewFile);
                    _sFileName = string.Empty;
                }
                _sFileName = string.Empty;
            }
            catch (Exception ex)
            {
                if (ex.Message.Equals("Error Opening file"))
                {
                    //파일이 열려있음 체크
                    XtraMessageBox.Show(((ex.InnerException).InnerException).Message);
                }
            }
        }

        private void UploadData(DevExpress.XtraGrid.Views.Grid.GridView gridView)
        {
            if(gridView.RowCount == 0)
            {
                return;
            }


            //string keratype = string.Empty;
            //string maipcher = string.Empty;
            //string jCompany = string.Empty;
            //string jBnum = string.Empty;
            //string agreeDate = string.Empty;
            //string gubun1 = string.Empty;
            //double weight = 0;
            //double idanga = 0;
            //double odanga = 0;
            //double ikongkep = 0;
            //double okongkep = 0;
            //string jState2 = string.Empty;
            //double secondWeight = 0;
            //double firstWeight = 0;
            //double iweight = 0;
            //double oweight = 0;
            //double ichagam = 0;
            //double transportKumak = 0;
            //string transportDanga = string.Empty; //운반업체명 #00007

            Cursor = Cursors.WaitCursor;
            
            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql_Test = new StringBuilder();
                string sAGREE_DATE = Convert.ToString(gridView.GetRowCellValue(0, "Agree_Date")).Substring(0, 10);
                if (string.IsNullOrEmpty(sAGREE_DATE))
                {
                    XtraMessageBox.Show("업로드 자료의 첫 행의 Agree_Date의 일자가 입력되지 않았습니다.\r\n업로드 자료를 다시한번 확인해주세요.");
                    return;
                }

                if (gridView.RowCount > 0)
                {
                    /*
                     * 2020-11-06 현업요청
                     * 직송업로드 시 AGREE_DATE 일자에 해당하는 전표내역이 존재할 경우 뒷단 업로드 취소되도록 구현
                     */
                    #region[2020-11-06 이전 체크로직]

                    ////직송 매입출ID 조회
                    //strSql_Test.Clear();
                    //strSql_Test.AppendLine(" SELECT X1.J_ID, A.JUNPYOID ");
                    //strSql_Test.AppendLine("   FROM ( ");
                    //strSql_Test.AppendLine("          SELECT IPCHULGO_MAIPID AS J_ID ");
                    //strSql_Test.AppendLine("            FROM MESURING  ");
                    //strSql_Test.AppendLine("           WHERE KERATYPE = '직송' ");
                    //strSql_Test.AppendLine("             AND J_DATE = '" + sAGREE_DATE + "' ");
                    //strSql_Test.AppendLine("           UNION ALL ");
                    //strSql_Test.AppendLine("          SELECT IPCHULGO_MACHULID AS J_ID ");
                    //strSql_Test.AppendLine("            FROM MESURING  ");
                    //strSql_Test.AppendLine("           WHERE KERATYPE = '직송' ");
                    //strSql_Test.AppendLine("             AND J_DATE = '" + sAGREE_DATE + "' ");
                    //strSql_Test.AppendLine("        ) X1 ");
                    //strSql_Test.AppendLine("   LEFT JOIN INLIST A  ");
                    //strSql_Test.AppendLine("     ON X1.J_ID = A.J_ID ");

                    //DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql_Test.ToString());
                    ////INLIST, ACTRAN에 해당 자료 삭제
                    //foreach(DataRow row in dt.Rows)
                    //{
                    //    string sJ_ID = row["J_ID"]?.ToString();
                    //    string sJunpyoId = row["JUNPYOID"]?.ToString();

                    //    strSql_Test.Clear();
                    //    strSql_Test.AppendLine(" DELETE FROM INLIST ");
                    //    strSql_Test.AppendLine("  WHERE J_ID = " + sJ_ID + " ");

                    //    cmd.CommandType = CommandType.Text;
                    //    cmd.CommandText = strSql_Test.ToString();
                    //    cmd.ExecuteNonQuery();

                    //    strSql_Test.Clear();
                    //    strSql_Test.AppendLine(" DELETE FROM IPCHULGO ");
                    //    strSql_Test.AppendLine("  WHERE J_ID = " + sJ_ID + " ");

                    //    cmd.CommandType = CommandType.Text;
                    //    cmd.CommandText = strSql_Test.ToString();
                    //    cmd.ExecuteNonQuery();
                    //}

                    #endregion[2020-11-06 이전 체크로직]
                    
                    Dictionary<string, string> dicParams = new Dictionary<string, string>();
                    dicParams.Clear();
                    dicParams.Add("TDATE", sAGREE_DATE.Replace("-","").Substring(0, 8));

                    strSql_Test.Clear();
                    strSql_Test.AppendLine(" ");
                    strSql_Test.AppendLine(" SELECT COUNT(*) AS CNT ");
                    strSql_Test.AppendLine("   FROM ACTRAN ");
                    strSql_Test.AppendLine("  WHERE TDATE = @TDATE ");
                    strSql_Test.AppendLine("    AND AAUTO = 'A02' --직송 HARDCODING");

                    DataTable dt = ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql_Test.ToString(), dicParams);
                    if(dt.Rows.Count > 0)
                    {
                        int iCnt = 0;
                        bool b = int.TryParse(dt.Rows[0]["CNT"]?.ToString(), out iCnt);

                        if (!b)
                        {
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            Cursor = Cursors.Default;
                            XtraMessageBox.Show("[시스템 에러]관리자에게 문의해주세요." +
                                "\r\n직송데이터 관련 Int.Parsing 부분에러");
                            return;
                        }
                        else
                        {
                            if(iCnt > 0)
                            {
                                DBConn.dbTran.Rollback();
                                DBConn.dbTran = null;
                                Cursor = Cursors.Default;
                                string sMsg = string.Format("일자 : {0}" +
                                "\r\n" + "해당 일자에 직송 관련 전표데이터가 존재합니다." +
                                "\r\n" + "매입출 자료승인에서 해당 일자의 전표데이터의 마감취소를 해주세요."
                                , sAGREE_DATE);
                                XtraMessageBox.Show(sMsg);
                                return;
                            }
                        }
                    }
                    else
                    {
                        DBConn.dbTran.Rollback();
                        DBConn.dbTran = null;
                        Cursor = Cursors.Default;
                        XtraMessageBox.Show("[시스템 에러]관리자에게 문의해주세요.");
                        return;
                    }
                    
                    //INLIST 및 IPCHULGO 삭제
                    strSql_Test.Clear();
                    #region mariaDB
                    //strSql_Test.AppendLine(" SELECT GROUP_CONCAT(SALE) AS DEL_J_ID ");
                    //strSql_Test.AppendLine("   FROM ( SELECT GROUP_CONCAT(ipchulgo_machulid) AS SALE ");
                    //strSql_Test.AppendLine("            FROM MESURING ");
                    //strSql_Test.AppendLine("           WHERE J_DATE = '" + sAGREE_DATE + "' ");
                    //strSql_Test.AppendLine("             AND KERATYPE = '직송' ");
                    //strSql_Test.AppendLine("           GROUP BY KERATYPE ");
                    //strSql_Test.AppendLine("           UNION ALL ");
                    //strSql_Test.AppendLine("           SELECT GROUP_CONCAT(ipchulgo_maipid)  ");
                    //strSql_Test.AppendLine("             FROM MESURING ");
                    //strSql_Test.AppendLine("            WHERE J_DATE = '" + sAGREE_DATE + "' ");
                    //strSql_Test.AppendLine("              AND KERATYPE = '직송' ");
                    //strSql_Test.AppendLine("            GROUP BY KERATYPE ) X1 ");
                    #endregion

                    strSql_Test.AppendLine("SELECT STRING_AGG(SALE,',') AS DEL_J_ID                                   ");
                    strSql_Test.AppendLine("  FROM(SELECT STRING_AGG(CONVERT(NUMERIC, ipchulgo_machulid), ',') AS SALE");
                    strSql_Test.AppendLine("           FROM MESURING                                                  ");
                    strSql_Test.AppendLine("          WHERE J_DATE = '" + sAGREE_DATE + "'                                     ");
                    strSql_Test.AppendLine("            AND KERATYPE = '직송'                                         ");
                    strSql_Test.AppendLine("          UNION ALL                                                       ");
                    strSql_Test.AppendLine("          SELECT STRING_AGG(CONVERT(NUMERIC, ipchulgo_maipid), ',')       ");
                    strSql_Test.AppendLine("            FROM MESURING                                                 ");
                    strSql_Test.AppendLine("           WHERE J_DATE = '" + sAGREE_DATE + "'                                    ");
                    strSql_Test.AppendLine("             AND KERATYPE = '직송'                                        ");
                    strSql_Test.AppendLine("       ) X1                                                               ");

                    DataTable dtJ_ID = DBConn.GetDataTable(DBConn.dbCon, strSql_Test.ToString());
                    if(dtJ_ID.Rows.Count > 0)
                    {
                        string sGROUP_JID = dtJ_ID.Rows[0]["DEL_J_ID"]?.ToString();

                        if (!string.IsNullOrEmpty(sGROUP_JID))
                        {
                            strSql_Test.Clear();
                            strSql_Test.AppendLine("  ");
                            strSql_Test.AppendLine(" DELETE FROM INLIST ");
                            strSql_Test.AppendLine("       WHERE J_ID IN (" + sGROUP_JID + ") ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql_Test.ToString();
                            cmd.ExecuteNonQuery();

                            strSql_Test.Clear();
                            strSql_Test.AppendLine("  ");
                            strSql_Test.AppendLine(" DELETE FROM IPCHULGO ");
                            strSql_Test.AppendLine("       WHERE J_ID IN (" + sGROUP_JID + ") ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql_Test.ToString();
                            cmd.ExecuteNonQuery();
                        }
                    }

                    strSql_Test.Clear();
                    strSql_Test.AppendLine(" DELETE FROM MESURING  ");
                    strSql_Test.AppendLine("  WHERE KERATYPE = '직송' ");
                    strSql_Test.AppendLine("    AND J_DATE = '" + sAGREE_DATE + "'  ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql_Test.ToString();
                    cmd.ExecuteNonQuery();

                    //strSql_Test.AppendLine(" ");
                    //strSql_Test.AppendLine(" SELECT COUNT(*) AS CNT ");
                    //strSql_Test.AppendLine("   FROM MESURING ");
                    //strSql_Test.AppendLine("  WHERE J_DATE = '" + sAGREE_DATE + "' ");
                    //strSql_Test.AppendLine("    AND KERATYPE = '직송' ");

                    //DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql_Test.ToString());
                    //double dCnt = Convert.ToDouble(dt.Rows[0]["CNT"]);
                    //if (dCnt > 0)
                    //{
                    //    Cursor = Cursors.Default;
                    //    XtraMessageBox.Show(string.Format("{0}에 이미 직송데이터가 존재합니다.", sAGREE_DATE));
                    //    return;
                    //}
                }

                for (int i = 0; i < gridView.RowCount; i++)
                {
                    string keratype = Convert.ToString(gridView.GetRowCellValue(i, "keratype"));
                    string maipcher = Convert.ToString(gridView.GetRowCellValue(i, "Maipcher"));
                    string jCompany = Convert.ToString(gridView.GetRowCellValue(i, "J_Company"));
                    string jBnum = Convert.ToString(gridView.GetRowCellValue(i, "J_Bnum"));
                    string agreeDate = Convert.ToString(gridView.GetRowCellValue(i, "Agree_Date"));
                    string sjTime = string.Empty;
                    if(DateTime.TryParse(agreeDate, out DateTime dateResult))
                    {
                        sjTime = dateResult.ToString("yyyy-MM-dd HH:mm;ss");
                        agreeDate = agreeDate.Substring(0, 10);
                    }
                    string gubun1 = Convert.ToString(gridView.GetRowCellValue(i, "Gubun1"));
                    double weight = Convert.ToDouble(gridView.GetRowCellValue(i, "weight"));
                    double idanga = Convert.ToDouble(gridView.GetRowCellValue(i, "idanga"));
                    double odanga = Convert.ToDouble(gridView.GetRowCellValue(i, "ODanga"));
                    double ikongkep = Convert.ToDouble(gridView.GetRowCellValue(i, "ikongkep"));
                    double okongkep = Convert.ToDouble(gridView.GetRowCellValue(i, "OKongkep"));
                    string jState2 = Convert.ToString(gridView.GetRowCellValue(i, "J_State2"));
                    double secondWeight = Convert.ToDouble(gridView.GetRowCellValue(i, "SecondWeight"));
                    double firstWeight = Convert.ToDouble(gridView.GetRowCellValue(i, "FirstWeight"));
                    double iweight = Convert.ToDouble(gridView.GetRowCellValue(i, "iweight"));
                    double oweight = Convert.ToDouble(gridView.GetRowCellValue(i, "OWeight"));
                    double ichagam = Convert.ToDouble(gridView.GetRowCellValue(i, "ichagam"));
                    double transportKumak = Convert.ToDouble(gridView.GetRowCellValue(i, "TransportKumak"));
                    string transportDanga = gridView.GetRowCellValue(i, "TransportDanga")?.ToString();//#00007

                    string sBsnsGb = string.Empty;
                    string sKey1 = string.Empty;
                    string sKey2 = string.Empty;
                    string sKey3 = string.Empty;
                    string sid = FmMainToolBar2.UserID;
                    string sNewSlipNo = string.Empty;
                    string sJid1 = string.Empty;
                    string sJid2 = string.Empty;

                    StringBuilder strSql = new StringBuilder();

                    #region // SLIP_NO 생성 

                    sBsnsGb = "SLIP_NO";
                    sKey1 = agreeDate.Replace("-","");
                    sKey2 = "****";
                    sKey3 = "****";
                    sid = FmMainToolBar2.UserID;

                    strSql.Clear();
                    //strSql.AppendLine(" SELECT SET_CRT_SEQ('" + sBsnsGb + "', '" + sKey1 + "','" + sKey2 + "','" + sKey3 + "','" + sid + "')");
                    strSql.AppendLine(" EXEC SET_CRT_SEQ '" + sBsnsGb + "', '" + sKey1 + "','" + sKey2 + "','" + sKey3 + "','" + sid + "' ");

                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    sNewSlipNo = dt.Rows[0]["NEW_SEQ"].ToString();

                    sJid1 = agreeDate.Replace("-", "") + sNewSlipNo;

                    #endregion

                    #region // SLIP_NO 생성 

                    sBsnsGb = "SLIP_NO";
                    sKey1 = agreeDate.Replace("-", "");
                    sKey2 = "****";
                    sKey3 = "****";
                    sid = FmMainToolBar2.UserID;

                    strSql = new StringBuilder();
                    //strSql.AppendLine(" SELECT SET_CRT_SEQ('" + sBsnsGb + "', '" + sKey1 + "','" + sKey2 + "','" + sKey3 + "','" + sid + "')");
                    strSql.AppendLine(" EXEC SET_CRT_SEQ '" + sBsnsGb + "', '" + sKey1 + "','" + sKey2 + "','" + sKey3 + "','" + sid + "' ");

                    dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    sNewSlipNo = dt.Rows[0]["NEW_SEQ"].ToString();

                    sJid2 = agreeDate.Replace("-", "") + sNewSlipNo;

                    #endregion

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" INSERT INTO MESURING ");
                    strSql.AppendLine("           ( J_CHECK,        KERATYPE,          MAIPCHERID,   MAIPCHER,   J_ASSIGNID,       J_COMPANY,    SUN,          J_DATE,       FIRSTTIME, SECONDTIME ");
                    strSql.AppendLine("           , FIRSTWEIGHT,    SECONDWEIGHT,      WEIGHT,       J_SERIAL,   GUBUN1,           J_BNUM,       K_NAME,       ICHAGAM,      OCHAGAM,   IGAMGA ");
                    strSql.AppendLine("           , OGAMGA,         IWEIGHT,           OWEIGHT,      IDANGA,     ODANGA,           IKONGKEP,     OKONGKEP,     FUSERCODE,    USERCODE,  FBUSEOCODE ");
                    strSql.AppendLine("           , BUSEOCODE,      P_ID,              J_GARAGE,     U_DATE,     J_ID,             KYERYANG12,   DRIVER_INOUT, AGREE_DATE,   J_STATE2,  TRANSPORTDANGA ");
                    strSql.AppendLine("           , TRANSPORTKUMAK, TRANSPORTC_SERIAL, CUSTOMWEIGHT, LOSSWEIGHT, TRANSPORTJUNGSAN, IPCHULGOJ_ID, MAGAM_FLAG,   WEIGHT_GUBUN, LENGTHSID, DAMAGE ");
                    strSql.AppendLine("           , GUMSU_SERIAL,   HALINYUL,          SURYANG  ");
                    strSql.AppendLine("           , IPCHULGO_MAIPID,IPCHULGO_MACHULID, ETC_DEALER_CD1, ETC_COST1) ");
                    strSql.AppendLine(" SELECT '1' AS J_CHECK ");
                    strSql.AppendLine("      , '" + keratype + "' AS KERATYPE ");
                    strSql.AppendLine("      , A.DEALER_CD AS MAIPCHERID ");
                    strSql.AppendLine("      , A.DEALER_NM AS MAIPCHER ");
                    strSql.AppendLine("      , B.DEALER_CD AS J_ASSIGNID ");
                    strSql.AppendLine("      , B.DEALER_NM AS J_COMPANY ");
                    strSql.AppendLine("      , (SELECT ISNULL(MAX(SUN), 0) + 1 FROM MESURING WHERE KERATYPE = '" + keratype + "' AND J_DATE = '" + agreeDate + "') AS SUN ");
                    strSql.AppendLine("      , '" + agreeDate + "'  AS J_DATE ");
                    strSql.AppendLine("      , '" + sjTime + "' AS FIRSTTIME ");
                    strSql.AppendLine("      , '" + sjTime + "' AS SECONDTIME ");
                    strSql.AppendLine("      , " + firstWeight + " AS FIRSTWEIGHT ");
                    strSql.AppendLine("      , " + secondWeight + " AS SECONDWEIGHT ");
                    strSql.AppendLine("      , " + weight + " AS WEIGHT ");
                    strSql.AppendLine("      , ( SELECT J_SERIAL                    ");
                    strSql.AppendLine("            FROM JAJAE                       ");
                    strSql.AppendLine("           WHERE GUBUN1 = '"+ gubun1 + "') AS J_SERIAL");
                    strSql.AppendLine("      , '" + gubun1 + "' AS GUBUN1 ");
                    strSql.AppendLine("      , '" + jBnum + "' AS J_BNUM ");
                    strSql.AppendLine("      , '상품' AS K_NAME  -- HARD CODING ");
                    strSql.AppendLine("      , " + ichagam + " AS ICHAGAM ");
                    strSql.AppendLine("      , 0 AS OCHAGAM -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS IGAMGA -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS OGAMGA -- HARD CODING ");
                    strSql.AppendLine("      , " + iweight + " AS IWEIGHT ");
                    strSql.AppendLine("      , " + oweight + " AS OWEIGHT ");
                    strSql.AppendLine("      , " + idanga + " AS IDANGA ");
                    strSql.AppendLine("      , " + odanga + " AS ODANGA ");
                    strSql.AppendLine("      , " + ikongkep + " AS IKONGKEP ");
                    strSql.AppendLine("      , " + okongkep + " AS OKONGKEP ");
                    strSql.AppendLine("      , 0 AS FUSERCODE -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS USERCODE -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS FBUSEOCODE -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS BUSEOCODE -- HARD CODING ");
                    strSql.AppendLine("      , 400 AS P_ID -- HARD CODING ");
                    strSql.AppendLine("      , 4 AS J_GARAGE -- HARD CODING ");
                    strSql.AppendLine("      , '" + agreeDate + "' AS U_DATE ");
                    strSql.AppendLine("      , 0 AS J_ID -- HARD CODING ");
                    strSql.AppendLine("      , 1 AS KYERYANG12 -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS DRIVER_INOUT -- HARD CODING ");
                    strSql.AppendLine("      , '" + agreeDate + "' AS AGREE_DATE ");
                    strSql.AppendLine("      , '" + jState2 + "' AS J_STATE2 ");
                    strSql.AppendLine("      , 0 AS TRANSPORTDANGA  -- HARD CODING ");
                    strSql.AppendLine("      , " + transportKumak + " AS TRANSPORTKUMAK ");
                    strSql.AppendLine("      , 0 AS TRANSPORTC_SERIAL -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS CUSTOMWEIGHT  -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS LOSSWEIGHT  -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS TRANSPORTJUNGSAN  -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS IPCHULGOJ_ID -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS MAGAM_FLAG -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS WEIGHT_GUBUN -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS LENGTHSID -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS DAMAGE -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS GUMSU_SERIAL  -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS HALINYUL -- HARD CODING ");
                    strSql.AppendLine("      , 0 AS SURYANG -- HARD CODING ");
                    strSql.AppendLine("      , "+ sJid1 + " AS IPCHULGO_MAIPID");
                    strSql.AppendLine("      , "+ sJid2 + " AS IPCHULGO_MACHULID ");
                    strSql.AppendLine("      , C.DEALER_CD AS ETC_DEALER_CD1");//#00007
                    strSql.AppendLine("      , " + transportKumak + " AS ETC_COST1");//#00007
                    strSql.AppendLine("   FROM ACC_DEALER_CD A ");
                    strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD B ");
                    strSql.AppendLine("     ON B.DEALER_NM = '" + jCompany + "'");
                    strSql.AppendLine("   LEFT JOIN ACC_DEALER_CD C ");
                    strSql.AppendLine("     ON C.DEALER_NM = '" + transportDanga + "'");
                    strSql.AppendLine("  WHERE A.DEALER_NM = '" + maipcher + "'");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    #region [매입 IPCHULGO, INLIST]
                    // IPCHULGO
                    strSql.Clear();
                    strSql.AppendLine(" SET IDENTITY_INSERT IPCHULGO ON");
                    strSql.AppendLine(" INSERT INTO IPCHULGO ");
                    strSql.AppendLine(" ( J_ID, J_TYPE, J_KIND, KERAGUBUN, J_TYPE1, J_TYPE2, J_DATE, J_LATER, J_ASSIGNID, J_COMPANY, J_BNUM ");
                    strSql.AppendLine(" , J_SUMSURYANG, J_SUMWEIGHT, J_SUMITEM, J_AMOUNT, J_KONGKEP, J_RAMOUNT, J_LAMOUNT, J_COUNT ");
                    strSql.AppendLine(" , J_BUGASE, J_CHECK, J_IPCHULGO, J_VATPOHAM, DAMDANG, CKONGKEP, P_ID, J_KYULMETHOD, FIRSTDATE ");
                    strSql.AppendLine(" , U_DATE, J_GARAGE, COSTBUGASE, J_SIGN1 ");
                    strSql.AppendLine(" ) ");
                    strSql.AppendLine(" SELECT '" + sJid1 + "' AS J_ID ");
                    strSql.AppendLine("      , '매입' AS J_TYPE ");
                    strSql.AppendLine("      , '일반' AS J_KIND ");
                    strSql.AppendLine("      , CASE WHEN A.KERATYPE <> '직송' THEN '계량'  ELSE A.KERATYPE END AS KERAGUBUN ");
                    strSql.AppendLine("      , '대체' AS J_TYPE1 ");
                    strSql.AppendLine("      , '일반' AS J_TYPE2 ");
                    strSql.AppendLine("      , J_DATE  ");
                    strSql.AppendLine("      , '" + agreeDate + "' AS J_LATER ");
                    strSql.AppendLine("      , MAIPCHERID AS J_ASSIGNID ");
                    strSql.AppendLine("      , MAIPCHER AS J_COMPANY ");
                    strSql.AppendLine("      , J_BNUM AS J_BNUM ");
                    strSql.AppendLine("      , 1 AS J_SUMSURYANG ");
                    strSql.AppendLine("      , A.IWEIGHT AS J_SUMWEIGHT ");
                    strSql.AppendLine("      , CASE WHEN A.KERATYPE = '직송' THEN 0 ELSE A.IWEIGHT END AS J_SUMITEM ");
                    strSql.AppendLine("      , A.IWEIGHT * A.IDANGA * 1.1 AS J_AMOUNT ");
                    strSql.AppendLine("      , A.IWEIGHT * A.IDANGA  AS J_KONGKEP ");
                    strSql.AppendLine("      , A.IWEIGHT * A.IDANGA * 1.1 AS J_RAMOUNT ");
                    strSql.AppendLine("      , A.IWEIGHT * A.IDANGA * 1.1 AS J_LAMOUNT ");
                    strSql.AppendLine("      , 1 AS J_COUNT ");
                    strSql.AppendLine("      , A.IWEIGHT * A.IDANGA * 0.1 AS J_BUGASE ");
                    strSql.AppendLine("      , A.J_CHECK  ");
                    strSql.AppendLine("      , 1 AS J_IPCHULGO ");
                    strSql.AppendLine("      , '별도' AS J_VATPOHAM ");
                    strSql.AppendLine("      , '80142' AS DAMDANG ");
                    strSql.AppendLine("      , TRANSPORTKUMAK AS CKONGKEP ");
                    strSql.AppendLine("      , P_ID ");
                    strSql.AppendLine("      , '외상' AS J_KYULMETHOD ");
                    strSql.AppendLine("      , CONVERT(VARCHAR(19),GETDATE(),20) AS FIRSTDATE ");
                    strSql.AppendLine("      , CONVERT(VARCHAR(19),GETDATE(),20) AS U_DATE ");
                    strSql.AppendLine("      , J_GARAGE ");
                    strSql.AppendLine("      , A.IWEIGHT * A.IDANGA * 0.1 AS COSTBUGASE ");
                    strSql.AppendLine("      , J_DATE AS J_SIGN1       ");
                    strSql.AppendLine("   FROM MESURING A   ");
                    strSql.AppendLine("  WHERE IPCHULGO_MAIPID = '" + sJid1 + "'  ");
                    strSql.AppendLine("  ");
                    strSql.AppendLine(" SET IDENTITY_INSERT IPCHULGO OFF");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    // INLIST
                    strSql.Clear();
                    strSql.AppendLine(" INSERT INTO INLIST ");
                    strSql.AppendLine(" ( J_CHECK, J_KIND, J_ID, J_ID1, J_SERIAL, GUBUN1 , J_CUSTOM, K_NAME, J_DATE, J_LATER, KERATYPE, KERAGUBUN ");
                    strSql.AppendLine(" , MODEL, J_LOTNO, J_BOOKING, J_SPEC, ISURYANG, OSURYANG, HALIN, DANJUNG, IWEIGHT, OWEIGHT, NWEIGHT ");
                    strSql.AppendLine(" , J_UNIT, DANGA, MIDANGA, CDANGA, GDANGA, EDANGA, MAIPCHER, IKONGKEP, KONGKEP, BUGASE, CKONGKEP ");
                    strSql.AppendLine(" , MIKONGKEP, GKONGKEP, EKONGKEP, SEAKPOHAM, J_RID, J_STATE, J_GARAGE, P_ID, NICKGUBUN1 ");
                    strSql.AppendLine(" , CHRG_ID, CHRG_NM ");
                    strSql.AppendLine(" ) ");
                    strSql.AppendLine(" SELECT '' AS J_CHECK  ");
                    strSql.AppendLine("      , '일반' AS J_KIND ");
                    strSql.AppendLine("      , '" + sJid1 + "' AS J_ID ");
                    strSql.AppendLine("      , MAIPCHERID AS J_ID1 ");
                    strSql.AppendLine("      , J_SERIAL ");
                    strSql.AppendLine("      , GUBUN1 ");
                    strSql.AppendLine("      , WEIGHT AS J_CUSTOM ");
                    strSql.AppendLine("      , K_NAME ");
                    strSql.AppendLine("      , J_DATE ");
                    strSql.AppendLine("      , '" + agreeDate + "' AS J_LATER ");
                    strSql.AppendLine("      , '매입' AS KERATYPE  ");
                    strSql.AppendLine("      , '계량' AS KERABUBUN ");
                    strSql.AppendLine("      , 'K1' AS MODEL ");
                    strSql.AppendLine("      , '4' AS J_LOTNO ");
                    strSql.AppendLine("      , ISNULL(J_COMPANY, '') AS J_BOOKING ");
                    strSql.AppendLine("      , '당사' AS J_SPEC ");
                    strSql.AppendLine("      , 1 AS ISURYANG ");
                    strSql.AppendLine("      , 0 AS OSURYANG ");
                    strSql.AppendLine("      , ICHAGAM AS HALIN ");
                    strSql.AppendLine("      , IWEIGHT AS DANJUNG ");
                    strSql.AppendLine("      , IWEIGHT ");
                    strSql.AppendLine("      , 0 AS OWEIGHT ");
                    strSql.AppendLine("      , CUSTOMWEIGHT AS NWEIGHT ");
                    strSql.AppendLine("      , 'KG' AS J_UNIT ");
                    strSql.AppendLine("      , IDANGA AS DANGA  -- IDANGA ");
                    strSql.AppendLine("      , CASE WHEN KERATYPE = '직송' THEN IDANGA ELSE (SELECT MAX(ISNULL(X.DANGA, 0)) FROM JAJAE X WHERE X.GUBUN1 = A.GUBUN1) END AS MIDANGA ");
                    strSql.AppendLine("      , 0 AS CDANGA ");
                    strSql.AppendLine("      , 0 AS GDANGA ");
                    strSql.AppendLine("      , 0 AS EDANGA ");
                    strSql.AppendLine("      , MAIPCHERID AS MAIPCHER ");
                    strSql.AppendLine("      , IKONGKEP AS IKONGKEP "); //2020-08-06 직송의 경우 인수량 단가를 고려하지 않고 금액만 변경하는 경우가 있어 금액을 바로 올릴 수 있도록 수정
                    //strSql.AppendLine("      , IWEIGHT * IDANGA AS IKONGKEP ");
                    strSql.AppendLine("      , NULL AS KONGKEP ");
                    strSql.AppendLine("      , IKONGKEP * 0.1 AS BUGASE ");
                    strSql.AppendLine("      , TRANSPORTKUMAK AS CKONGKEP ");
                    strSql.AppendLine("      , CASE WHEN KERATYPE = '직송' THEN IWEIGHT * IDANGA ELSE 0 END AS MIKONGKEP ");
                    strSql.AppendLine("      , 0 AS GKONGKEP ");
                    strSql.AppendLine("      , 0 AS EKONGKEP ");
                    strSql.AppendLine("      , 'N' AS SEAKPOHAM ");
                    strSql.AppendLine("      , JUNPYOID AS J_RID ");
                    strSql.AppendLine("      , LOSSWEIGHT AS J_STATE ");
                    strSql.AppendLine("      , J_GARAGE AS J_GARAGE ");
                    strSql.AppendLine("      , P_ID AS P_ID ");
                    strSql.AppendLine("      , J_STATE AS NICKGUBUN1 ");
                    strSql.AppendLine("      , ( SELECT X1.CHRG_ID  ");
                    strSql.AppendLine("            FROM ACC_DEALER_CD X1 ");
                    strSql.AppendLine("           WHERE X1.DEALER_CD = A.MAIPCHERID ) AS CHRG_ID ");
                    strSql.AppendLine("      , ( SELECT X2.EMP_NM ");
                    strSql.AppendLine("            FROM ACC_DEALER_CD X1 ");
                    strSql.AppendLine("            LEFT JOIN HR_EMP_BASIS X2  ");
                    strSql.AppendLine("              ON X1.CHRG_ID = X2.EMP_ID ");
                    strSql.AppendLine("           WHERE X1.DEALER_CD = A.MAIPCHERID ) AS CHRG_NM ");
                    strSql.AppendLine("   FROM MESURING A ");
                    strSql.AppendLine("   WHERE IPCHULGO_MAIPID = '" + sJid1 + "'  ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    #endregion

                    #region [매출 IPCHULGO, INLIST]

                    // ACC.IPCHULGO
                    strSql.Clear();
                    strSql.AppendLine(" SET IDENTITY_INSERT IPCHULGO ON");
                    strSql.AppendLine(" INSERT INTO IPCHULGO ");
                    strSql.AppendLine(" ( J_ID, J_TYPE, J_KIND, KERAGUBUN, J_TYPE1, J_TYPE2, J_DATE, J_LATER, J_ASSIGNID, J_COMPANY, J_BNUM ");
                    strSql.AppendLine(" , J_SUMSURYANG, J_SUMWEIGHT, J_SUMITEM, J_AMOUNT, J_KONGKEP, J_RAMOUNT, J_LAMOUNT, J_COUNT ");
                    strSql.AppendLine(" , J_BUGASE, J_CHECK, J_IPCHULGO, J_VATPOHAM, DAMDANG, CKONGKEP, P_ID, J_KYULMETHOD, FIRSTDATE ");
                    strSql.AppendLine(" , U_DATE, J_GARAGE, COSTBUGASE, J_SIGN1) ");
                    strSql.AppendLine(" SELECT '" + sJid2 + "' AS J_ID ");
                    strSql.AppendLine("      , '매출' AS J_TYPE ");
                    strSql.AppendLine("      , '일반' AS J_KIND ");
                    strSql.AppendLine("      , CASE WHEN A.KERATYPE <> '직송' THEN '계량'  ELSE A.KERATYPE END AS KERAGUBUN ");
                    strSql.AppendLine("      , '대체' AS J_TYPE1 ");
                    strSql.AppendLine("      , '일반' AS J_TYPE2 ");
                    strSql.AppendLine("      , J_DATE  ");
                    strSql.AppendLine("      , J_DATE AS J_LATER ");
                    strSql.AppendLine("      , J_ASSIGNID AS J_ASSIGNID ");
                    strSql.AppendLine("      , MAIPCHER AS J_COMPANY ");
                    strSql.AppendLine("      , J_BNUM AS J_BNUM ");
                    strSql.AppendLine("      , 1 AS J_SUMSURYANG ");
                    strSql.AppendLine("      , A.OWEIGHT AS J_SUMWEIGHT ");
                    strSql.AppendLine("      , CASE WHEN A.KERATYPE = '직송' THEN 0 ELSE A.OWEIGHT END AS J_SUMITEM ");
                    strSql.AppendLine("      , A.OWEIGHT * A.ODANGA * 1.1 AS J_AMOUNT ");
                    strSql.AppendLine("      , A.OWEIGHT * A.ODANGA AS J_KONGKEP ");
                    strSql.AppendLine("      , A.OWEIGHT * A.ODANGA * 1.1 AS J_RAMOUNT ");
                    strSql.AppendLine("      , A.OWEIGHT * A.ODANGA * 1.1 AS J_LAMOUNT ");
                    strSql.AppendLine("      , 1 AS J_COUNT ");
                    strSql.AppendLine("      , A.OWEIGHT * A.ODANGA * 0.1 AS J_BUGASE ");
                    strSql.AppendLine("      , A.J_CHECK  ");
                    strSql.AppendLine("      , 1 AS J_IPCHULGO ");
                    strSql.AppendLine("      , '별도' AS J_VATPOHAM ");
                    strSql.AppendLine("      , '38126' AS DAMDANG ");
                    strSql.AppendLine("      , TRANSPORTKUMAK AS CKONGKEP ");
                    strSql.AppendLine("      , P_ID ");
                    strSql.AppendLine("      , '외상' AS J_KYULMETHOD ");
                    strSql.AppendLine("      , CONVERT(VARCHAR(19),GETDATE(),20) AS FIRSTDATE ");
                    strSql.AppendLine("      , CONVERT(VARCHAR(19),GETDATE(),20) AS U_DATE ");
                    strSql.AppendLine("      , J_GARAGE ");
                    strSql.AppendLine("      , A.OWEIGHT * A.ODANGA * 0.1 AS COSTBUGASE ");
                    strSql.AppendLine("      , J_DATE AS J_SIGN1 ");
                    strSql.AppendLine("   FROM MESURING A   ");
                    strSql.AppendLine("  WHERE IPCHULGO_MACHULID = '" + sJid2 + "'  ");
                    strSql.AppendLine(" SET IDENTITY_INSERT IPCHULGO OFF");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    // ACC.INLIST
                    strSql.Clear();
                    strSql.AppendLine(" INSERT INTO INLIST ");
                    strSql.AppendLine(" ( J_CHECK, J_KIND, J_ID, J_ID1, J_SERIAL, GUBUN1 , J_CUSTOM, K_NAME, J_DATE, J_LATER, KERATYPE, KERAGUBUN ");
                    strSql.AppendLine(" , MODEL, J_LOTNO, J_BOOKING, J_SPEC, ISURYANG, OSURYANG, HALIN, DANJUNG, IWEIGHT, OWEIGHT, NWEIGHT ");
                    strSql.AppendLine(" , J_UNIT, DANGA, MIDANGA, CDANGA, GDANGA, EDANGA, MAIPCHER, IKONGKEP, KONGKEP, BUGASE, CKONGKEP ");
                    strSql.AppendLine(" , MIKONGKEP, GKONGKEP, EKONGKEP, SEAKPOHAM, J_RID, J_STATE, J_GARAGE, P_ID, NICKGUBUN1 ");
                    strSql.AppendLine(" , CHRG_ID, CHRG_NM ");
                    strSql.AppendLine(" ) ");
                    strSql.AppendLine(" SELECT '' AS J_CHECK  ");
                    strSql.AppendLine("      , '일반' AS J_KIND ");
                    strSql.AppendLine("      , '" + sJid2 + "' AS J_ID ");
                    strSql.AppendLine("      , J_AssignID AS J_ID1 ");
                    strSql.AppendLine("      , J_SERIAL ");
                    strSql.AppendLine("      , GUBUN1 ");
                    strSql.AppendLine("      , WEIGHT AS J_CUSTOM ");
                    strSql.AppendLine("      , K_NAME ");
                    strSql.AppendLine("      , J_DATE ");
                    strSql.AppendLine("      , J_DATE AS J_LATER ");
                    strSql.AppendLine("      , '매출' AS KERATYPE  ");
                    strSql.AppendLine("      , '계량' AS KERAGUBUN ");
                    strSql.AppendLine("      , 'K1' AS MODEL ");
                    strSql.AppendLine("      , '4' AS J_LOTNO ");
                    strSql.AppendLine("      , ISNULL(MaipCher, '') AS J_BOOKING ");
                    strSql.AppendLine("      , CASE WHEN KERATYPE = '직송' THEN '당사' ELSE '업체' END AS J_SPEC ");
                    strSql.AppendLine("      , 0 AS ISURYANG ");
                    strSql.AppendLine("      , 1 AS OSURYANG ");
                    strSql.AppendLine("      , OCHAGAM AS HALIN ");
                    strSql.AppendLine("      , OWEIGHT AS DANJUNG ");
                    strSql.AppendLine("      , 0 AS IWEIGHT ");
                    strSql.AppendLine("      , OWEIGHT ");
                    strSql.AppendLine("      , CUSTOMWEIGHT AS NWEIGHT ");
                    strSql.AppendLine("      , 'KG' AS J_UNIT ");
                    strSql.AppendLine("      , ODANGA AS DANGA  ");
                    //#00008
                    strSql.AppendLine("      , ( SELECT X.SELLPRC2 FROM JAJAE X WHERE X.J_SERIAL = A.J_SERIAL ) AS MIDANGA                      ");
                    //strSql.AppendLine("      , CASE WHEN KERATYPE = '직송' THEN IDANGA ELSE (SELECT MAX(ISNULL(X.DANGA, 0)) FROM JAJAE X WHERE X.GUBUN1 = A.GUBUN1) END AS MIDANGA ");
                    strSql.AppendLine("      , 0 AS CDANGA ");
                    strSql.AppendLine("      , 0 AS GDANGA ");
                    strSql.AppendLine("      , 0 AS EDANGA ");
                    strSql.AppendLine("      , J_AssignID AS MAIPCHER ");
                    strSql.AppendLine("      , NULL AS IKONGKEP ");
                    strSql.AppendLine("      , OKONGKEP AS KONGKEP "); //2020-08-06 직송의 경우 인수량 단가를 고려하지 않고 금액만 변경하는 경우가 있어 금액을 바로 올릴 수 있도록 수정
                    //strSql.AppendLine("      , OWEIGHT * ODANGA AS KONGKEP ");
                    strSql.AppendLine("      , OKONGKEP * 0.1 AS BUGASE ");
                    strSql.AppendLine("      , TRANSPORTKUMAK AS CKONGKEP ");
                    strSql.AppendLine("      , IWEIGHT * IDANGA AS MIKONGKEP ");
                    strSql.AppendLine("      , IWEIGHT * IDANGA * -1 AS GKONGKEP ");
                    strSql.AppendLine("      , 0 AS EKONGKEP ");
                    strSql.AppendLine("      , 'N' AS SEAKPOHAM ");
                    strSql.AppendLine("      , JUNPYOID AS J_RID           ");
                    strSql.AppendLine("      , LOSSWEIGHT AS J_STATE ");
                    strSql.AppendLine("      , J_GARAGE AS J_GARAGE ");
                    strSql.AppendLine("      , P_ID AS P_ID ");
                    strSql.AppendLine("      , J_STATE AS NICKGUBUN1 ");
                    strSql.AppendLine("      , ( SELECT X1.CHRG_ID  ");
                    strSql.AppendLine("            FROM ACC_DEALER_CD X1 ");
                    strSql.AppendLine("           WHERE X1.DEALER_CD = A.MAIPCHERID ) AS CHRG_ID ");
                    strSql.AppendLine("      , ( SELECT X2.EMP_NM ");
                    strSql.AppendLine("            FROM ACC_DEALER_CD X1 ");
                    strSql.AppendLine("            LEFT JOIN HR_EMP_BASIS X2  ");
                    strSql.AppendLine("              ON X1.CHRG_ID = X2.EMP_ID ");
                    strSql.AppendLine("           WHERE X1.DEALER_CD = A.MAIPCHERID ) AS CHRG_NM ");
                    strSql.AppendLine("   FROM MESURING A ");
                    strSql.AppendLine("   WHERE IPCHULGO_MACHULID = '" + sJid2 + "'  ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                    #endregion
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                Cursor = Cursors.Default;

                if(gridView.RowCount > 0)
                {
                    XtraMessageBox.Show("정상적으로 저장을 완료했습니다.");

                    string agreeDate = Convert.ToString(gridView.GetRowCellValue(0, "Agree_Date")).Substring(0, 10);

                    DtpRetr.EditValue = agreeDate;
                    DtpRetrEnd.EditValue = agreeDate;

                    RdGrClose.SelectedIndex = 1;

                    BtnRetr.PerformClick();

                }
                else
                {
                    XtraMessageBox.Show("업로드할 직송자료가 없습니다.");
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }

        }

        private void AccMeasureCloseDev_FormClosed(object sender, FormClosedEventArgs e)
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
                string sFileNM = "마감 현황 리스트";
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
                    XtraMessageBox.Show(((ex.InnerException).InnerException).Message);
                }
            }
        }

        private void GridViewRetr_RowClick(object sender, RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                if (GridViewRetr.IsRowSelected(e.RowHandle))
                    GridViewRetr.UnselectRow(e.RowHandle);
                else
                    GridViewRetr.SelectRow(e.RowHandle);
            }
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
            //        GridViewRetr.FocusedRowHandle = 0;
            //    }
            //    else if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.Descending)
            //    {
            //        hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.Ascending;
            //        GridViewRetr.FocusedRowHandle = 0;
            //    }
            //    else if (hitInfo.Column.SortOrder == DevExpress.Data.ColumnSortOrder.Ascending)
            //    {
            //        hitInfo.Column.SortOrder = DevExpress.Data.ColumnSortOrder.None;
            //        GridViewRetr.FocusedRowHandle = 0;
            //    }
            //    // if ((ModifierKeys & Keys.Control) == Keys.Control) return;
            //    //if ((ModifierKeys & Keys.Shift) != Keys.Shift) view.ClearSorting();
            //}
        }

        #endregion

        private void GridViewRetr_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void ChkVatPoham_EditValueChanged(object sender, EventArgs e)
        {
            CheckEdit chk = (CheckEdit)sender;
            GridViewRetr.SetFocusedRowCellValue(GridColVatYn, chk.EditValue);
        }

        private void GridViewRetr_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridViewRetr.UpdateCurrentRow();
        }

        private void ChkComboGb_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}