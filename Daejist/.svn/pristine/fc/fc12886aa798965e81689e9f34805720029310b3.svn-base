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
using MySql.Data.MySqlClient;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.BandedGrid.ViewInfo;
using DevExpress.Data;
using System.Data.SqlClient;

/*
 * 작성일자 : 모름
 * 작성자 : 고혜성
 * ---------------------HISTORY-----------------------
 * 수정일자 : 2021-02-02
 * 수정자 : 고혜성
 * 수정내용 : (현업요청)
 *            저장 시 데이터바뀔시 로그로직 추가
 *            
 * 수정일자 : 2021-02-25 ~ 2021-02-26
 * 수정자   : 고혜성
 * 수정내용 : (현업요청)
 *            1. 그리드 폰트 설정
 *            2. 레이아웃 전체 저장 설정
 * 
 * 수정일자 : 2021-03-17
 * 수정자   : 고혜성
 * Reference Key : #0001
 * 수정내용 : (현업요청)
 *            1. 로그작업
 *              1-1) 기본사항/변경사항을 나누어 입력
 *            
 * 수정일자 : 2021-03-18
 * 수정자   : 고혜성
 * 수정내용 : (현업요청) #0001 관련
 *            1. 기본항목에 입력되는 값 단순화
 *            
 * 수정일자 : 2021-03-24
 * 수정자   : 고혜성
 * Reference Key : #0002
 * 수정내용 : (현업요청)
 *            스크랩반품, 슈레더반품(JAJAE)의 경우 리스트에 나타나지 않도록 쿼리수정
 *            
 * 수정일자 : 2022-04-13
 * 수정자   : 정은영
 * key      : #0003
 * 수정내용 : (현업요청)
 *            단가입력 후 저장 했을때 승인으로 인해 저장 실패 시 적혀있던값 원래 값으로 리프레쉬
 *            
 * 수정일자 : 2022-08-19
 * 수정자   : 정은영
 * key      : #0004
 * 수정내용 : (현업요청)
 *            단가입력안해도 수정된 내용 저장되도록 변경
 *            필터추가
 *            
 * 수정일자 : 2022-12-09
 * 수정자   : 정은영
 * key      : #0005
 * 수정내용 : (현업요청)
 *            1. 매출 기준 단가 현대제철로 변경
 *             
 */
namespace AccAdm
{
    public partial class IN07001F01 : DevExpress.XtraEditors.XtraForm
    {
        public IN07001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void IN07001F01_Load(object sender, EventArgs e)
        {
            //ComnEtcFunc.SetDateFromToValue(DateEditFrom, DateEditTo);
            DateEditFrom.EditValue = DateTime.Now;
            DateEditTo.EditValue = DateTime.Now;
            ComnGridFunc.SetInitGridRowColor(GridViewRetr);
            CboSubject.SelectedIndex = 0;

            DataTable LookupDt = GetLookupData("1", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGridLkupDealer1, LookupDt, GridRetr, GridColDealerCd1, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGridLkupDealer1, LookupDt, GridRetr, GridColDealerCd2, "CD", "NM", "");

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
            if( FmMainToolBar2.EmpID == "3002" || FmMainToolBar2.EmpID == "3003")
            {
                RdgbKeraType.EditValue = "입고";
                layoutControlItem10.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;

            }
        }

        private DataTable GetLookupData(string sGb, string sNullYn)
        {
            StringBuilder strSql = new StringBuilder();

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '' AS CD ");
                strSql.AppendLine("      , '' AS NM ");
                strSql.AppendLine("  UNION ALL ");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT CAST(DEALER_CD AS CHAR) AS CD     ");
                strSql.AppendLine(" 	 , DEALER_NM AS NM     ");
                strSql.AppendLine(" FROM ACC_DEALER_CD");
                strSql.AppendLine("WHERE EOB_YN = 'N'");
            }

            strSql.AppendLine("  ORDER BY CD ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                ComnEtcFunc.YmdFromToValuesCheck(DateEditFrom, DateEditTo);
                string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
                string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
                string sUtPrcYn = RdgbCloseYn.EditValue?.ToString();
                string sAddQuery = string.Empty;
                string sVal = LkupFindWord.EditValue?.ToString().Replace(" ", "");
                string sKeraType = RdgbKeraType.EditValue?.ToString();
                if(!string.IsNullOrEmpty(sVal))
                    sAddQuery = GetWhereQuery(CboSubject.EditValue?.ToString(), sVal).ToString();

                GridRetr.DataSource = GetMesuringData(sYmdFrom, sYmdTo, sUtPrcYn, sAddQuery, sKeraType);
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void BtnDataClose_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            //#0004
            DataTable dtRetr = GridRetr.DataSource as DataTable;
            DataTable dtModi = dtRetr.GetChanges(DataRowState.Modified);

            if (dtModi == null || dtModi.Rows.Count == 0)
            {
                XtraMessageBox.Show("편집한 데이터가 없습니다.");
                return;
            }

            if (XtraMessageBox.Show(string.Format("편집한 데이터를 모두 저장하시겠습니까?"), "저장확인", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                for(int i = 0; i < dtModi.Rows.Count; i++)
                {
                    //string sVal = dtModi.Rows[i]["CHK"]?.ToString();

                    Dictionary<string, string[]> dicCompare = new Dictionary<string, string[]>();

                    string sJunpyoId = dtModi.Rows[i]["JUNPYOID"]?.ToString();
                    string sKeratype = dtModi.Rows[i]["KERATYPE"]?.ToString();
                    //sUnitPrc = dt.Rows[i]["APPLY_UNIT_PRC"]?.ToString();
                    string sUnitPrc = dtModi.Rows[i]["APPLY_UNIT_PRC"]?.ToString();
                    string sCarryCost = dtModi.Rows[i]["TRANSPORTKUMAK"]?.ToString();

                    double dUnitPrc = string.IsNullOrEmpty(sUnitPrc) ? 0 : Convert.ToDouble(sUnitPrc);
                    double dCarryCost = string.IsNullOrEmpty(sCarryCost) ? 0 : Convert.ToDouble(sCarryCost);

                    string sEtcDealerCd1 = dtModi.Rows[i]["ETC_DEALER_CD1"]?.ToString();
                    double dEtcDealerCd1 = string.IsNullOrEmpty(sEtcDealerCd1) ? 0 : Convert.ToDouble(sEtcDealerCd1);

                    string sEtcRemark1 = dtModi.Rows[i]["ETC_REMARK1"]?.ToString();

                    string sEtcCost1 = dtModi.Rows[i]["ETC_COST1"]?.ToString();
                    double dEtcCost1 = string.IsNullOrEmpty(sEtcCost1) ? 0 : Convert.ToDouble(sEtcCost1);

                    string sEtcDealerCd2 = dtModi.Rows[i]["ETC_DEALER_CD2"]?.ToString();
                    double dEtcDealerCd2 = string.IsNullOrEmpty(sEtcDealerCd2) ? 0 : Convert.ToDouble(sEtcDealerCd2);

                    string sEtcRemark2 = dtModi.Rows[i]["ETC_REMARK2"]?.ToString();

                    string sEtcCost2 = dtModi.Rows[i]["ETC_COST2"]?.ToString();
                    double dEtcCost2 = string.IsNullOrEmpty(sEtcCost2) ? 0 : Convert.ToDouble(sEtcCost2);

                    string sJCheck = dtModi.Rows[i]["J_CHECK"]?.ToString();

                    string sAcptWeight = dtModi.Rows[i]["ACPT_WEIGHT"].ToString();
                    double dAcptWeight = string.IsNullOrEmpty(sAcptWeight) ? 0 : Convert.ToDouble(sAcptWeight);

                    string sEtcDealerNm1 = dtModi.Rows[i]["ETC_DEALER_NM1"]?.ToString();
                    string sEtcDealerNm2 = dtModi.Rows[i]["ETC_DEALER_NM2"]?.ToString();

                    strSql.Clear();
                    strSql.AppendLine(" SELECT APRV_YN ");
                    strSql.AppendLine("   FROM INLIST  ");
                    strSql.AppendLine("  WHERE J_ID = (");
                    strSql.AppendLine(" SELECT CASE WHEN KERATYPE = '입고' THEN IPCHULGO_MAIPID ");
                    strSql.AppendLine(" 	   		WHEN KERATYPE = '출고' THEN IPCHULGO_MACHULID END");
                    strSql.AppendLine("   FROM MESURING ");
                    strSql.AppendLine("  WHERE JUNPYOID = '" + sJunpyoId + "')");

                    DataTable dtAprvYN = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    if(dtAprvYN.Rows.Count > 0)
                    {
                        string sAprvYN = dtAprvYN.Rows[0][0].ToString();
                        if (sAprvYN.Equals("Y"))
                        {
                            XtraMessageBox.Show("마감승인된 자료는 단가수정이 불가능합니다.");
                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;

                            /*
                                * #0003
                                * 2022-04-13
                                * (현업요청)
                                * 승인 안되면 리프레쉬 추가
                                */
                            DataTable tempDt = GridRetr.DataSource as DataTable;

                            DataTable tempDtModi = tempDt.GetChanges(DataRowState.Modified);

                            if (tempDtModi != null)
                            {
                                foreach (DataRow dr in tempDtModi.Rows)
                                {
                                    if (dr.RowState == DataRowState.Modified)
                                    {
                                        string sJunpyoID = dr["JUNPYOID"].ToString();
                                        // 수정되기 전 값 
                                        var boforeValue1 = dr["APPLY_UNIT_PRC", DataRowVersion.Original];
                                        var boforeValue2 = dr["TRANSPORTKUMAK", DataRowVersion.Original];

                                        for (int j = 0; j < tempDt.Rows.Count; j++)
                                        {
                                            string sJunpyoID_1 = tempDt.Rows[j]["JUNPYOID"]?.ToString();
                                            if (sJunpyoID_1.Equals(sJunpyoID))
                                            {
                                                tempDt.Rows[j]["APPLY_UNIT_PRC"] = boforeValue1;
                                                tempDt.Rows[j]["TRANSPORTKUMAK"] = boforeValue2;
                                            }
                                        }
                                    }
                                }
                            }

                            return;
                        }
                    }

                    string sLogMsg = string.Empty;

                    if (sKeratype.Equals("입고"))
                    {
                        /*
                            * 2021-02-02
                            * 변경된 컬럼 색출하여 로그기록
                            */
                        dicCompare.Clear();
                        dicCompare.Add("IDANGA", new[] { dUnitPrc.ToString(), "적용단가" });
                        dicCompare.Add("TRANSPORTKUMAK", new[] { dCarryCost.ToString(), "운반비" });
                        dicCompare.Add("ETC_DEALER_NM1", new[] { sEtcDealerNm1.ToString(), "거래처명1" });
                        dicCompare.Add("ETC_COST1", new[] { dEtcCost1.ToString(), "기타단가1" });
                        dicCompare.Add("ETC_REMARK1", new[] { sEtcRemark1, "비고1" });
                        dicCompare.Add("ETC_DEALER_NM2", new[] { sEtcDealerNm2.ToString(), "거래처명2" });
                        dicCompare.Add("ETC_COST2", new[] { dEtcCost2.ToString(), "기타단가2" });
                        dicCompare.Add("ETC_REMARK2", new[] { sEtcRemark2, "비고2" });

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" SELECT A1.IDANGA ");
                        strSql.AppendLine("      , A1.TRANSPORTKUMAK ");
                        strSql.AppendLine("      , A1.ETC_DEALER_CD1 ");
                        strSql.AppendLine("      , B1.DEALER_NM AS ETC_DEALER_NM1 ");
                        strSql.AppendLine("      , A1.ETC_COST1 ");
                        strSql.AppendLine("      , A1.ETC_REMARK1 ");
                        strSql.AppendLine("      , A1.ETC_DEALER_CD2 ");
                        strSql.AppendLine("      , B2.DEALER_NM AS ETC_DEALER_NM2 ");
                        strSql.AppendLine("      , A1.ETC_COST2 ");
                        strSql.AppendLine("      , A1.ETC_REMARK2 ");
                        strSql.AppendLine("   FROM MESURING A1 ");
                        strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD B1          ");
                        strSql.AppendLine("    ON A1.ETC_DEALER_CD1 = CONVERT(VARCHAR,B1.DEALER_CD)");
                        strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD B2          ");
                        strSql.AppendLine("    ON A1.ETC_DEALER_CD2 = CONVERT(VARCHAR,B2.DEALER_CD)");
                        strSql.AppendLine("  WHERE A1.JUNPYOID = " + sJunpyoId + " ");

                        DataTable dtLog = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                        string sPrvODanga = dtLog.Rows[0]["IDANGA"]?.ToString();
                        int iCCnt = 0;
                        foreach (KeyValuePair<string, string[]> param in dicCompare)
                        {
                            if (!sPrvODanga.Equals("0")) //로그는 최초 0 일시에는 남기지 않음
                            {
                                if (dtLog.Columns.Contains(param.Key))
                                {
                                    string sPrvVal = dtLog.Rows[0][param.Key]?.ToString();
                                    if (!sPrvVal.Equals(param.Value[0]))
                                    {
                                        if (iCCnt++ == 0)
                                        {
                                            sLogMsg += string.Format("{0} : {1}  ▶ {2}", param.Value[1], sPrvVal, param.Value[0]);
                                        }
                                        else
                                        {
                                            sLogMsg += string.Format("| {0} : {1}  ▶ {2}", param.Value[1], sPrvVal, param.Value[0]);
                                        }
                                    }
                                }
                            }
                        }
                            
                        strSql.Clear();
                        strSql.AppendLine(" UPDATE MESURING  ");
                        strSql.AppendLine("    SET IDANGA = " + dUnitPrc + " ");
                        strSql.AppendLine("      , TRANSPORTKUMAK = " + dCarryCost + " ");
                        strSql.AppendLine("      , IKONGKEP = ROUND(" + dUnitPrc * dAcptWeight + ", 0) ");
                        strSql.AppendLine("      , UNIT_PRC_CHG_YN = 'Y' ");
                        strSql.AppendLine("      , ETC_DEALER_CD1 = '"+ sEtcDealerCd1 + "'");
                        strSql.AppendLine("      , ETC_COST1 = "+ dEtcCost1 + "");
                        strSql.AppendLine("      , ETC_REMARK1 = '"+ sEtcRemark1 + "'");
                        strSql.AppendLine("      , ETC_DEALER_CD2 = '"+ sEtcDealerCd2 + "'");
                        strSql.AppendLine("      , ETC_COST2 = "+ dEtcCost2 + "");
                        strSql.AppendLine("      , ETC_REMARK2 = '"+ sEtcRemark2 + "'");
                        strSql.AppendLine("      , EDIT_GB = 'C' "); //A : 마감 및 업로드, B : 계근프로그램, C : 단가입력
                        strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoId + " ");
                    }
                    else if (sKeratype.Equals("출고"))
                    {
                        /*
                            * 2021-02-02
                            * 변경된 컬럼 색출하여 로그기록
                            */
                        dicCompare.Clear();
                        dicCompare.Add("ODANGA", new[] { dUnitPrc.ToString(), "적용단가" });
                        dicCompare.Add("TRANSPORTKUMAK", new[] { dCarryCost.ToString(), "운반비" });
                        dicCompare.Add("ETC_DEALER_NM1", new[] { sEtcDealerNm1.ToString(), "거래처명1" });
                        dicCompare.Add("ETC_COST1", new[] { dEtcCost1.ToString(), "기타단가1" });
                        dicCompare.Add("ETC_REMARK1", new[] { sEtcRemark1, "비고1" });
                        dicCompare.Add("ETC_DEALER_NM2", new[] { sEtcDealerNm2.ToString(), "거래처명2" });
                        dicCompare.Add("ETC_COST2", new[] { dEtcCost2.ToString(), "기타단가2" });
                        dicCompare.Add("ETC_REMARK2", new[] { sEtcRemark2, "비고2" });

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" SELECT A1.ODANGA ");
                        strSql.AppendLine("      , A1.TRANSPORTKUMAK ");
                        strSql.AppendLine("      , A1.ETC_DEALER_CD1 ");
                        strSql.AppendLine("      , B1.DEALER_NM AS ETC_DEALER_NM1 ");
                        strSql.AppendLine("      , A1.ETC_COST1 ");
                        strSql.AppendLine("      , A1.ETC_REMARK1 ");
                        strSql.AppendLine("      , A1.ETC_DEALER_CD2 ");
                        strSql.AppendLine("      , B2.DEALER_NM AS ETC_DEALER_NM2 ");
                        strSql.AppendLine("      , A1.ETC_COST2 ");
                        strSql.AppendLine("      , A1.ETC_REMARK2 ");
                        strSql.AppendLine("   FROM MESURING A1 ");
                        strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD B1          ");
                        strSql.AppendLine("    ON A1.ETC_DEALER_CD1 = CONVERT(VARCHAR,B1.DEALER_CD)");
                        strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD B2          ");
                        strSql.AppendLine("    ON A1.ETC_DEALER_CD2 = CONVERT(VARCHAR,B2.DEALER_CD)");
                        strSql.AppendLine("  WHERE A1.JUNPYOID = " + sJunpyoId + " ");

                        DataTable dtLog = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                        string sPrvODanga = dtLog.Rows[0]["ODANGA"]?.ToString();
                        int iCCnt = 0;
                        if (!sPrvODanga.Equals("0")) //로그는 최초 0 일시에는 남기지 않음
                        {
                            foreach (KeyValuePair<string, string[]> param in dicCompare)
                            {
                                if (dtLog.Columns.Contains(param.Key))
                                {
                                    string sPrvVal = dtLog.Rows[0][param.Key]?.ToString();
                                    if (!sPrvVal.Equals(param.Value[0]))
                                    {
                                        if(iCCnt++ == 0)
                                        {
                                            sLogMsg += string.Format("{0} : {1}  ▶ {2}", param.Value[1], sPrvVal, param.Value[0]);
                                        }
                                        else
                                        {
                                            sLogMsg += string.Format("| {0} : {1}  ▶ {2}", param.Value[1], sPrvVal, param.Value[0]);
                                        }
                                    }
                                }
                            }
                        }
                            
                        strSql.Clear();
                        strSql.AppendLine(" UPDATE MESURING  ");
                        strSql.AppendLine("    SET ODANGA = " + dUnitPrc + " ");
                        strSql.AppendLine("      , TRANSPORTKUMAK = " + dCarryCost + " ");
                        strSql.AppendLine("      , OKONGKEP = ROUND(" + dUnitPrc * dAcptWeight + ", 0) ");
                        strSql.AppendLine("      , UNIT_PRC_CHG_YN = 'Y' ");
                        strSql.AppendLine("      , ETC_DEALER_CD1 = '" + sEtcDealerCd1 + "'");
                        strSql.AppendLine("      , ETC_COST1 = " + dEtcCost1 + "");
                        strSql.AppendLine("      , ETC_REMARK1 = '" + sEtcRemark1 + "'");
                        strSql.AppendLine("      , ETC_DEALER_CD2 = '" + sEtcDealerCd2 + "'");
                        strSql.AppendLine("      , ETC_COST2 = " + dEtcCost2 + "");
                        strSql.AppendLine("      , ETC_REMARK2 = '" + sEtcRemark2 + "'");
                        strSql.AppendLine("      , EDIT_GB = 'C' "); //A : 마감 및 업로드, B : 계근프로그램, C : 단가입력
                        strSql.AppendLine("  WHERE JUNPYOID = " + sJunpyoId + " ");
                    }

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    //LogInsert
                    if (!string.IsNullOrEmpty(sLogMsg))
                    {
                        /*
                            * #0001
                            */
                        StringBuilder sb = new StringBuilder();
                            
                        string sLogBasic = string.Format("{0}" +
                            "/{2}" +
                            "/{3}" +
                            "/순번:{1}" +
                            "/차번:{4}" +
                            "/{5}"
                            , dtModi.Rows[i]["J_DATE"]?.ToString().Substring(0, 10)
                            , dtModi.Rows[i]["SUN"]?.ToString()
                            , dtModi.Rows[i]["KERATYPE"]?.ToString()
                            , dtModi.Rows[i]["DEALER_NM"]?.ToString()
                            , dtModi.Rows[i]["J_BNUM"]?.ToString()
                            , dtModi.Rows[i]["GRADE_NM"]?.ToString());

                        string sREF_RMK = string.Format("TABLE : MESURING / JUNPYOID : {0}", sJunpyoId);

                        //sLogMsg = sLogBasic + sLogMsg;
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
                        cmd.Parameters.AddWithValue("@STD_COLS", sLogBasic);
                        cmd.Parameters.AddWithValue("@REF_RMK", sREF_RMK);
                        cmd.Parameters.AddWithValue("@EDIT_RMK", sLogMsg);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                    }

                    /*
                        * 2020-06-25 설명추가(고혜성)
                        * 해당 로직은 MESURING의 J_CHECK이며 만약 마감및업로드에서 단가미입력 데이터를 마감처리할 경우 해당 INLIST테이블에서 수정
                        */
                    if (sJCheck.Equals("1"))
                    {
                        strSql.Clear();
                        strSql.AppendLine(" UPDATE INLIST          ");
                        strSql.AppendLine("    SET DANGA = "+ dUnitPrc + "     ");
                        strSql.AppendLine("   	 , CKONGKEP = "+ dCarryCost + "");
                        if (sKeratype.Equals("입고"))
                        {
                            strSql.AppendLine("  , IKONGKEP = ROUND(" + dUnitPrc * dAcptWeight + ", 0) ");
                            strSql.AppendLine("  , BUGASE = ROUND(IWEIGHT * " + dUnitPrc + " * 0.1, 0) ");
                        }
                        else if (sKeratype.Equals("출고"))
                        {
                            strSql.AppendLine("  , KONGKEP = ROUND(" + dUnitPrc * dAcptWeight + ", 0) ");
                            strSql.AppendLine("  , BUGASE = ROUND(OWEIGHT * " + dUnitPrc + " * 0.1, 0) ");
                        }
                        strSql.AppendLine("  WHERE J_ID = (");
                        strSql.AppendLine(" SELECT CASE WHEN KERATYPE = '입고' THEN IPCHULGO_MAIPID ");
                        strSql.AppendLine(" 	   		WHEN KERATYPE = '출고' THEN IPCHULGO_MACHULID END");
                        strSql.AppendLine("   FROM MESURING ");
                        strSql.AppendLine("  WHERE JUNPYOID = '"+ sJunpyoId + "')");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");

                BtnRetr_Click(null, null);
            }
            catch(Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            ComnEtcFunc.ExportExcelFile(this.Name + "_", GridRetr);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private StringBuilder GetWhereQuery(string sText, string sVal)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();

            if (sText.Equals("담당자"))
            {
                strSql.AppendLine("    AND B.CHRG_ID = " + sVal + " ");
            }
            else if(sText.Equals("거래처명"))
            {
                strSql.AppendLine("    AND B.DEALER_CD = " + sVal + " ");
            }

            return strSql;
        }

        private DataTable GetMesuringData(string sYmdFrom, string sYmdTo, string sUtPrcYn, string sQuery, string sKeraType)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            #region mariaDB
            //strSql.AppendLine(" SELECT 'N' AS CHK ");
            //strSql.AppendLine(" 	 , A.JUNPYOID ");
            //strSql.AppendLine(" 	 , A.SUN ");
            //strSql.AppendLine(" 	 , A.J_DATE ");
            //strSql.AppendLine(" 	 , A.KERATYPE ");
            //strSql.AppendLine(" 	 , CASE A.KERATYPE WHEN '입고' THEN A.MAIPCHER ");
            //strSql.AppendLine(" 	     			   WHEN '출고' THEN A.J_COMPANY END AS DEALER_NM ");
            //strSql.AppendLine(" 	 , A.J_BNUM ");
            //strSql.AppendLine(" 	 , D.GUBUN1 AS GRADE_NM ");
            //strSql.AppendLine(" 	 , C.EMP_NM AS CHRG_NM ");
            //strSql.AppendLine(" 	 , A.SECONDWEIGHT AS TOT_WEIGHT ");
            //strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.FIRSTTIME, '%H : %I') ELSE DATE_FORMAT(A.SECONDTIME, '%H : %I') END AS TIME1 ");
            //strSql.AppendLine(" 	 , A.FIRSTWEIGHT AS EMPTY_WEIGHT ");
            //strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.SECONDTIME, '%H : %I') ELSE DATE_FORMAT(A.FIRSTTIME, '%H : %I') END AS TIME2 ");
            //strSql.AppendLine(" 	 , A.WEIGHT AS LANDED_WEIGHT ");
            //strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '출고' THEN A.OCHAGAM ELSE A.ICHAGAM END AS CHAGAM_WEIGHT ");
            //strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.IWEIGHT ELSE A.OWEIGHT END AS ACPT_WEIGHT ");
            //strSql.AppendLine(" 	 , A.CUSTOMWEIGHT AS COMP_WEIGHT ");
            //strSql.AppendLine(" 	 , A.LOSSWEIGHT AS LOSS_WEIGHT ");
            //strSql.AppendLine("      , CASE WHEN A.KERATYPE = '입고' THEN D.DANGA ELSE D.SELLPRC1 END AS STD_UNIT_PRC ");
            //strSql.AppendLine(" 	 , CASE WHEN A.KERATYPE = '입고' THEN A.IDANGA ELSE A.ODANGA END AS APPLY_UNIT_PRC ");
            //strSql.AppendLine("      , D.DANGA - ( CASE WHEN A.KERATYPE = '입고' THEN A.IDANGA ELSE A.ODANGA END ) AS DIFF_UNIT_PRC  ");
            //strSql.AppendLine("      , IFNULL(TRUNCATE(A.TRANSPORTKUMAK / IFNULL(F.DANJUNG, CASE WHEN A.KERATYPE = '입고' THEN A.IWEIGHT ELSE A.OWEIGHT END ), 1), 0) AS CRY_UNIT_PRC ");
            //strSql.AppendLine(" 	 , A.UNIT_PRC_CHG_YN ");
            //strSql.AppendLine("      , A.TRANSPORTKUMAK ");
            //strSql.AppendLine("      , A.ETC_DEALER_CD1 ");
            //strSql.AppendLine("      , A.ETC_REMARK1 ");
            //strSql.AppendLine("      , A.ETC_COST1 ");
            //strSql.AppendLine("      , A.ETC_DEALER_CD2 ");
            //strSql.AppendLine("      , A.ETC_REMARK2 ");
            //strSql.AppendLine("      , A.ETC_COST2 ");
            //strSql.AppendLine("      , A.J_CHECK ");
            //strSql.AppendLine("   FROM MESURING A ");
            //strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD B ");
            //strSql.AppendLine("     ON (CASE A.KERATYPE WHEN '입고' THEN A.MAIPCHERID ");
            //strSql.AppendLine("                         WHEN '출고' THEN A.J_ASSIGNID END) = B.DEALER_CD ");
            //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS C ");
            //strSql.AppendLine("     ON B.CHRG_ID = C.EMP_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN JAJAE D ");
            //strSql.AppendLine("     ON A.J_SERIAL = D.J_SERIAL ");
            //strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS E   ");
            //strSql.AppendLine("     ON A.GUMSU_SERIAL = E.EMP_ID ");
            //strSql.AppendLine("   LEFT OUTER JOIN INLIST F ");
            //strSql.AppendLine("     ON (CASE A.KERATYPE WHEN '입고' THEN A.IPCHULGO_MAIPID ");
            //strSql.AppendLine("                         WHEN '출고' THEN A.IPCHULGO_MACHULID END) = F.J_ID ");
            //strSql.AppendLine("  WHERE A.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            //strSql.AppendLine("    AND A.KERATYPE <> '직송' ");
            //strSql.AppendLine("    AND B.DEALER_CD <> 6303870006 ");//HARDCODING, 재고이동
            ////strSql.AppendLine("    AND A.UNIT_PRC_CHG_YN = '" + sUtPrcYn + "' ");
            //if (!sKeraType.Equals("ALL"))
            //    strSql.AppendLine("    AND A.KERATYPE = '" + sKeraType + "' ");
            //if(sUtPrcYn.Equals("Y"))
            //    strSql.AppendLine("    AND CASE WHEN A.KERATYPE = '입고' THEN A.IDANGA > 0 ELSE A.ODANGA > 0 END");
            //else if(sUtPrcYn.Equals("N"))
            //    strSql.AppendLine("    AND CASE WHEN A.KERATYPE = '입고' THEN A.IDANGA = 0 ELSE A.ODANGA = 0 END");
            //strSql.AppendLine("    AND D.DAEGUBUN IN ('고철A', '고철B', '슈레더') ");
            //strSql.AppendLine("    AND A.J_SERIAL NOT IN ('4049042', '5050042')"); //#0002
            //strSql.AppendLine(sQuery);
            //strSql.AppendLine("  ORDER BY A.J_DATE, A.SUN  ");
            #endregion
            strSql.AppendLine("   SET ANSI_WARNINGS OFF");
            strSql.AppendLine("   SET ARITHIGNORE ON   ");
            strSql.AppendLine("   SET ARITHABORT OFF   ");
            strSql.AppendLine("SELECT 'N' AS CHK                                                     ");
            strSql.AppendLine("     , A.JUNPYOID                                                     ");
 	        strSql.AppendLine("     , A.SUN                                                          ");
 	        strSql.AppendLine("     , A.J_DATE                                                       ");
 	        strSql.AppendLine("     , A.KERATYPE                                                     ");
 	        strSql.AppendLine("     , CASE A.KERATYPE WHEN '입고' THEN A.MAIPCHER                    ");
            strSql.AppendLine("                         WHEN '출고' THEN A.J_COMPANY END AS DEALER_NM");
            strSql.AppendLine("     , A.J_BNUM                                                       ");
 	        strSql.AppendLine("     , D.GUBUN1 AS GRADE_NM                                           ");
 	        strSql.AppendLine("     , C.EMP_NM AS CHRG_NM                                            ");
            strSql.AppendLine("     , A.SECONDWEIGHT AS TOT_WEIGHT                                   ");
 	        strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN CONVERT(VARCHAR(5),CONVERT(DATETIME, A.FIRSTTIME),8) ELSE CONVERT(VARCHAR(5),CONVERT(DATETIME, A.SECONDTIME),8) END AS TIME1");
 	        strSql.AppendLine("     , A.FIRSTWEIGHT AS EMPTY_WEIGHT                                                                                                                                  ");
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN CONVERT(VARCHAR(5),CONVERT(DATETIME, A.SECONDTIME),8) ELSE CONVERT(VARCHAR(5),CONVERT(DATETIME, A.FIRSTTIME),8) END AS TIME2");
 	        strSql.AppendLine("     , A.WEIGHT AS LANDED_WEIGHT                                                        ");
 	        strSql.AppendLine("     , CASE WHEN A.KERATYPE = '출고' THEN A.OCHAGAM ELSE A.ICHAGAM END AS CHAGAM_WEIGHT ");
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN A.IWEIGHT ELSE A.OWEIGHT END AS ACPT_WEIGHT   ");
            strSql.AppendLine("     , A.CUSTOMWEIGHT AS COMP_WEIGHT                                                    ");
 	        strSql.AppendLine("     , A.LOSSWEIGHT AS LOSS_WEIGHT                                                      ");
            //#0005
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN D.DANGA ELSE D.SELLPRC2 END AS STD_UNIT_PRC   ");
            //strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN D.DANGA ELSE D.SELLPRC1 END AS STD_UNIT_PRC   ");
            strSql.AppendLine("     , CASE WHEN A.KERATYPE = '입고' THEN A.IDANGA ELSE A.ODANGA END AS APPLY_UNIT_PRC  ");
            strSql.AppendLine("     , D.DANGA - (CASE WHEN A.KERATYPE = '입고' THEN A.IDANGA ELSE A.ODANGA END ) AS DIFF_UNIT_PRC");
            strSql.AppendLine("     , ISNULL(ROUND(A.TRANSPORTKUMAK / ISNULL(F.DANJUNG, CASE WHEN A.KERATYPE = '입고' THEN A.IWEIGHT ELSE A.OWEIGHT END), 1), 0) AS CRY_UNIT_PRC");
            strSql.AppendLine("     , A.UNIT_PRC_CHG_YN");
            strSql.AppendLine("     , A.TRANSPORTKUMAK ");
            strSql.AppendLine("     , A.ETC_DEALER_CD1 ");
            strSql.AppendLine("     , A.ETC_REMARK1    ");
            strSql.AppendLine("     , A.ETC_COST1      ");
            strSql.AppendLine("     , A.ETC_DEALER_CD2 ");
            strSql.AppendLine("     , A.ETC_REMARK2    ");
            strSql.AppendLine("     , A.ETC_COST2      ");
            strSql.AppendLine("     , A.J_CHECK        ");
            strSql.AppendLine("     , G.DEALER_NM AS ETC_DEALER_NM1");
            strSql.AppendLine("     , H.DEALER_NM AS ETC_DEALER_NM2");
            strSql.AppendLine("  FROM MESURING A       ");
            strSql.AppendLine("  LEFT OUTER JOIN ACC_DEALER_CD B");
            strSql.AppendLine("    ON(CASE A.KERATYPE WHEN '입고' THEN A.MAIPCHERID");
            strSql.AppendLine("                       WHEN '출고' THEN A.J_ASSIGNID END) = B.DEALER_CD");
            strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS C ");
            strSql.AppendLine("    ON B.CHRG_ID = C.EMP_ID      ");
            strSql.AppendLine("  LEFT OUTER JOIN JAJAE D        ");
            strSql.AppendLine("    ON A.J_SERIAL = D.J_SERIAL   ");
            strSql.AppendLine("  LEFT OUTER JOIN HR_EMP_BASIS E ");
            strSql.AppendLine("    ON A.GUMSU_SERIAL = E.EMP_ID ");
            strSql.AppendLine("  LEFT OUTER JOIN INLIST F       ");
            strSql.AppendLine("    ON(CASE A.KERATYPE WHEN '입고' THEN A.IPCHULGO_MAIPID");
            strSql.AppendLine("                        WHEN '출고' THEN A.IPCHULGO_MACHULID END) = F.J_ID");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD G          ");
            strSql.AppendLine("    ON A.ETC_DEALER_CD1 = CONVERT(VARCHAR,G.DEALER_CD)");
            strSql.AppendLine("  LEFT JOIN ACC_DEALER_CD H          ");
            strSql.AppendLine("    ON A.ETC_DEALER_CD2 = CONVERT(VARCHAR,H.DEALER_CD)");
            strSql.AppendLine(" WHERE A.J_DATE BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "'");
            strSql.AppendLine("   AND A.KERATYPE <> '직송'                          ");
            strSql.AppendLine("    AND B.DEALER_CD <> 6303870006 ");//HARDCODING, 재고이동
            //strSql.AppendLine("    AND A.UNIT_PRC_CHG_YN = '" + sUtPrcYn + "' ");
            if (!sKeraType.Equals("ALL"))
            {
                strSql.AppendLine("    AND A.KERATYPE = '" + sKeraType + "' ");
            }
            if (sUtPrcYn.Equals("Y"))
            {
                if (sKeraType.Equals("입고"))
                    strSql.AppendLine("    AND A.IDANGA > 0");
                else if (sKeraType.Equals("출고"))
                    strSql.AppendLine("    AND A.ODANGA > 0");
                else
                    strSql.AppendLine("    AND (A.IDANGA > 0 OR A.ODANGA > 0)");
            }
            else if (sUtPrcYn.Equals("N"))
            {
                if (sKeraType.Equals("입고"))
                    strSql.AppendLine("    AND A.IDANGA = 0");
                else if (sKeraType.Equals("출고"))
                    strSql.AppendLine("    AND A.ODANGA = 0");
                else
                    strSql.AppendLine("    AND (A.IDANGA = 0 AND A.ODANGA = 0 )");
            }
            strSql.AppendLine("    AND D.DAEGUBUN IN ('고철A', '고철B', '슈레더') ");
            strSql.AppendLine("    AND A.J_SERIAL NOT IN ('4049042', '5050042')"); //#0002
            strSql.AppendLine(sQuery);
            strSql.AppendLine("  ORDER BY A.J_DATE, A.SUN  ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }
        
        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void RepoTxtUnitPrice_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            string sVal = txt.EditValue?.ToString();
            if (string.IsNullOrEmpty(sVal) || sVal.Equals("0"))
            {
                GridViewRetr.SetFocusedRowCellValue("CHK", "N");
            }
            else
            {
                GridViewRetr.SetFocusedRowCellValue(GridColUnitPrice, txt.EditValue);
                GridViewRetr.SetFocusedRowCellValue("CHK", "Y");
            }
        }

        private void RepoTxtUnitPrice_Leave(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            double dVal = string.IsNullOrEmpty(txt.EditValue?.ToString()) ? 0 : Convert.ToDouble(txt.EditValue);
            double dStdUnitPrc = string.IsNullOrEmpty(GridViewRetr.GetFocusedRowCellValue(GridColSttdUnitPrc)?.ToString()) ? 0 : Convert.ToDouble(GridViewRetr.GetFocusedRowCellValue(GridColSttdUnitPrc));
            GridViewRetr.SetFocusedRowCellValue(GridColDiffUnitPrc, (dStdUnitPrc - dVal));
        }

        private void RepoTxtNumOnly_EditValueChanged(object sender, EventArgs e)
        {
            //TextEdit txt = (TextEdit)sender;
            //string sVal = txt.EditValue?.ToString();
            //if (string.IsNullOrEmpty(sVal) || sVal.Equals("0"))
            //{
            //    GridViewRetr.SetFocusedRowCellValue("CHK", "N");
            //}
            //else
            //{
            //    GridViewRetr.SetFocusedRowCellValue(GridColUnitPrice, txt.EditValue);
            //    GridViewRetr.SetFocusedRowCellValue("CHK", "Y");
            //}
        }

        private void RepoTxtNumOnly_Leave(object sender, EventArgs e)
        {
            //TextEdit txt = (TextEdit)sender;
            //double dVal = string.IsNullOrEmpty(txt.EditValue?.ToString()) ? 0 : Convert.ToDouble(txt.EditValue);
            //double dStdUnitPrc = string.IsNullOrEmpty(GridViewRetr.GetFocusedRowCellValue(GridColSttdUnitPrc)?.ToString()) ? 0 : Convert.ToDouble(GridViewRetr.GetFocusedRowCellValue(GridColSttdUnitPrc));
            //GridViewRetr.SetFocusedRowCellValue(GridColDiffUnitPrc, (dStdUnitPrc - dVal));
        }

        private void RepoTxtCarryCost_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            string sVal = txt.EditValue?.ToString();
            if (!string.IsNullOrEmpty(sVal))
            {
                GridViewRetr.SetFocusedRowCellValue("TRANSPORTKUMAK", sVal);
                GridViewRetr.SetFocusedRowCellValue("CHK", "Y");
            }
        }

        private void RepoTxtCarryCost_Leave(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            double dVal = string.IsNullOrEmpty(txt.EditValue?.ToString()) ? 0 : Convert.ToDouble(txt.EditValue);
            double dWeight = string.IsNullOrEmpty(GridViewRetr.GetFocusedRowCellValue(GridColAcptWeight)?.ToString()) ? 0 : Convert.ToDouble(GridViewRetr.GetFocusedRowCellValue(GridColAcptWeight));
            GridViewRetr.SetFocusedRowCellValue(GridColCarryCostUnitPrc, Math.Round((dVal / dWeight), 1));
        }

        private void RdgbCloseYn_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        private void CboSubject_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                ComboBoxEdit combo = (ComboBoxEdit)sender;
                string sVal = combo.EditValue?.ToString();
                if (string.IsNullOrEmpty(sVal))
                    return;

                DataTable dt = new DataTable();

                if (sVal.Equals("담당자"))
                {
                    dt = GetLookUpData("1", "Y", "Y");
                }
                else if (sVal.Equals("거래처명"))
                {
                    dt = GetLookUpData("2", "Y", "Y");
                }

                ComLib.ComGrid.SetLookUpEdit(LkupFindWord, dt, "CD", "NM", "Y");
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");
            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '공용' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.EMP_ID AS CD");
                strSql.AppendLine(" 	 , A.EMP_NM AS NM ");
                strSql.AppendLine("      , A.DEPT_CD AS SEQ ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A ");
                strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD B ");
                strSql.AppendLine("     ON A.DEPT_CD = B.DEPT_CD ");
                strSql.AppendLine("  WHERE B.DEPT_CD = '3000' ");//부서코드 변경으로 인해 수정(5000->3000)
                strSql.AppendLine("    AND A.EMPL_GB ='Y'");

                //strSql.AppendLine(" SELECT A.DEALER_CD AS CD");
                //strSql.AppendLine("      , A.DEALER_NM AS NM");
                //strSql.AppendLine("      , A.DEALER_CD AS SEQ ");
                //strSql.AppendLine("   FROM ACC_DEALER_CD A");
                //strSql.AppendLine("  WHERE A.DEALER_GB = '매입'");
                //strSql.AppendLine("     OR A.DEALER_GB = '매출'");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT CONVERT(VARCHAR,DEALER_CD) AS CD ");
                strSql.AppendLine("      , DEALER_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEALER_NM) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
                strSql.AppendLine("  WHERE EOB_YN = 'N' ");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE CD_GB = 'ITEM_INOUT_GB'");
            }

            if (sParam.Equals("Y"))
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
                strSql.AppendLine("     ORDER BY SEQ");
            }

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }
        
        private void LkupFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void IN07001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
                BtnDataClose_Click(null, null);
            else if (e.KeyCode == Keys.F5)
                BtnRetr_Click(null, null);
            else if (e.KeyCode == Keys.F8)
                BtnExcel_Click(null, null);
        }

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (e.Clicks == 1)
            {
                //string chkVal = GridViewRetr.GetFocusedRowCellValue(GridColClo).ToString();
                //if(chkVal.Equals("Y"))
                //    GridViewRetr.SetFocusedRowCellValue(GridColClo, "N");
                //else
                //    GridViewRetr.SetFocusedRowCellValue(GridColClo, "Y");

                //if (chkVal.Equals("Y") || chkVal.Equals("True"))
                //{
                //    GridViewRetr.SetFocusedRowCellValue(GridColClo, "N");
                //}
                //else if (chkVal.Equals("N") || chkVal.Equals("False"))
                //{
                //    GridViewRetr.SetFocusedRowCellValue(GridColClo, "Y");
                //    GridViewRetr.FocusedColumn = GridViewRetr.VisibleColumns[15];
                //}

                //string sJunpyoId = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();

                //StringBuilder strSql = new StringBuilder();

                //strSql.Clear();
                //strSql.AppendLine(" SELECT COUNT(1) AS CNT ");
                //strSql.AppendLine("   FROM ACTRAN A ");
                //strSql.AppendLine("   LEFT OUTER JOIN INLIST B ");
                //strSql.AppendLine("     ON A.REF1 = B.JUNPYOID ");
                //strSql.AppendLine("   LEFT OUTER JOIN MESURING C ");
                //strSql.AppendLine("     ON B.J_RID = C.JUNPYOID ");
                //strSql.AppendLine("  WHERE C.JUNPYOID = '" + sJunpyoId + "' ");

                //DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                //Cursor = Cursors.Default;
                //if (dt.Rows.Count > 0)
                //{
                //    double dCnt = Convert.ToDouble(dt.Rows[0]["CNT"]);
                //    if (dCnt > 0)
                //    {
                //        XtraMessageBox.Show("해당 내역은 전표가 존재합니다.");
                //        GridViewRetr.SetFocusedRowCellValue(GridColClo, "N");
                //        return;
                //    }
                //}
            }
        }

        #region [정렬기능(2020-06-02 정은영)]2022-08-19 필터추가요청으로 제거 #0004
        //private ColumnSortOrder GetNextSortOrder(ColumnSortOrder order)
        //{
        //    switch (order)
        //    {
        //        case ColumnSortOrder.None: return ColumnSortOrder.Descending;
        //        case ColumnSortOrder.Descending: return ColumnSortOrder.Ascending;
        //        case ColumnSortOrder.Ascending: return ColumnSortOrder.None;
        //    }

        //    return ColumnSortOrder.None;
        //}

        //private void GridViewColumnSort_MouseUp(object sender, MouseEventArgs e)
        //{
        //    BandedGridView view = (BandedGridView)sender;
        //    BandedGridHitInfo hitInfo = view.CalcHitInfo(e.Location);

        //    if (hitInfo.InBandPanel)
        //    {
        //        if (hitInfo.Band == null)
        //            return;

        //        if (hitInfo.Band.Name.Equals("gridBand9")) //거래처명
        //        {
        //            ColumnSortOrder order = view.Columns["DEALER_NM"].SortOrder;
        //            order = GetNextSortOrder(order);

        //            hitInfo.Band.Columns[0].SortOrder = order;

        //            view.FocusedRowHandle = 0;
        //            if (order == ColumnSortOrder.Descending)
        //                hitInfo.Band.Caption = "거래처명↓";
        //            else if (order == ColumnSortOrder.Ascending)
        //                hitInfo.Band.Caption = "거래처명↑";
        //            else if (order == ColumnSortOrder.None)
        //                hitInfo.Band.Caption = "거래처명";
        //        }
        //        else if (hitInfo.Band.Name.Equals("gridBand10")) //담당자
        //        {
        //            ColumnSortOrder order = view.Columns["CHRG_NM"].SortOrder;
        //            order = GetNextSortOrder(order);

        //            hitInfo.Band.Columns[0].SortOrder = order;

        //            view.FocusedRowHandle = 0;
        //            if (order == ColumnSortOrder.Descending)
        //                hitInfo.Band.Caption = "담당자↓";
        //            else if (order == ColumnSortOrder.Ascending)
        //                hitInfo.Band.Caption = "담당자↑";
        //            else if (order == ColumnSortOrder.None)
        //                hitInfo.Band.Caption = "담당자";
        //        }
        //        else if (hitInfo.Band.Name.Equals("gridBand12")) //등급
        //        {
        //            ColumnSortOrder order = view.Columns["GRADE_NM"].SortOrder;
        //            order = GetNextSortOrder(order);

        //            hitInfo.Band.Columns[0].SortOrder = order;

        //            view.FocusedRowHandle = 0;
        //            if (order == ColumnSortOrder.Descending)
        //                hitInfo.Band.Caption = "등급↓";
        //            else if (order == ColumnSortOrder.Ascending)
        //                hitInfo.Band.Caption = "등급↑";
        //            else if (order == ColumnSortOrder.None)
        //                hitInfo.Band.Caption = "등급";
        //        }
        //    }
        //}

        /*
         * 2020-11-17 현업요청 로우클릭 시 체크전환되던 것에서 체크박스를 클릭할 때에만 체크되도록 수정
         */
        private void GridViewRetr_MouseDown(object sender, MouseEventArgs e)
        {
            BandedGridView view = (BandedGridView)sender;
            BandedGridHitInfo hitInfo = view.CalcHitInfo(e.Location);

            if (hitInfo.InRowCell)
            {
                //if (hitInfo.Band == null)
                //    return;

                //if (!hitInfo.InColumn)
                //    return;

                if(hitInfo.Column == GridColClo)
                {
                    string sYn = view.GetRowCellValue(hitInfo.RowHandle, GridColClo)?.ToString();

                    sYn = string.IsNullOrEmpty(sYn) ? string.Empty : sYn;
                    if (sYn.Equals("Y"))
                        view.SetRowCellValue(hitInfo.RowHandle, GridColClo, "N");
                    else
                        view.SetRowCellValue(hitInfo.RowHandle, GridColClo, "Y");

                }
                
            }
            //전체선택
            else if (hitInfo.InBandPanel)
            {
                if (hitInfo.Band == null)
                    return;

                if(hitInfo.Band == gridBand5)
                {
                    string sYn = string.Empty;
                    for(int i = 0; i < view.RowCount; i++)
                    {
                        if(i == 0)
                        {
                            sYn = view.GetRowCellValue(i, GridColClo)?.ToString();
                        }

                        if (sYn.Equals("Y"))
                            view.SetRowCellValue(i, GridColClo, "N");
                        else
                            view.SetRowCellValue(i, GridColClo, "Y");
                    }
                }
            }
        }

        #endregion

        private void IN07001F01_TextChanged(object sender, EventArgs e)
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

        private void RepoChkCloseYn_EditValueChanged(object sender, EventArgs e)
        {
            string sJunpyoId = GridViewRetr.GetFocusedRowCellValue("JUNPYOID")?.ToString();

            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT COUNT(1) AS CNT ");
            strSql.AppendLine("   FROM ACTRAN A ");
            strSql.AppendLine("   LEFT OUTER JOIN INLIST B ");
            strSql.AppendLine("     ON A.REF1 = B.JUNPYOID ");
            strSql.AppendLine("   LEFT OUTER JOIN MESURING C ");
            strSql.AppendLine("     ON B.J_RID = C.JUNPYOID ");
            strSql.AppendLine("  WHERE C.JUNPYOID = '" + sJunpyoId + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            Cursor = Cursors.Default;
            if (dt.Rows.Count > 0)
            {
                double dCnt = Convert.ToDouble(dt.Rows[0]["CNT"]);
                if (dCnt > 0)
                {
                    XtraMessageBox.Show("해당 내역은 전표가 존재합니다.");
                    CheckEdit chk = (CheckEdit)sender;
                    chk.EditValue = "N";
                    return;
                }
            }
        }

        private void RdgbKeraType_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        private void GridViewRetr_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if(e.Column == GridColUnitPrice || e.Column == GridColCarryCost)
            {
                double dVal = string.IsNullOrEmpty(e.CellValue?.ToString()) ? 0 : Convert.ToDouble(e.CellValue?.ToString());
                if (dVal > 0)
                    e.Appearance.BackColor = Color.PaleGreen;
                else if (dVal == 0)
                    e.Appearance.BackColor = SystemColors.Info;
            }
        }

        private void RepoTxtETC_Num0_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit txt = (TextEdit)sender;
            GridViewRetr.SetFocusedRowCellValue(GridViewRetr.FocusedColumn, txt.EditValue);
        }

        //#0004
        private void BtnAllDan_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)GridRetr.DataSource;
            List<string> list = new List<string>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string sVal = dt.Rows[i]["CHK"]?.ToString();
                if (!string.IsNullOrEmpty(sVal))
                {
                    if (sVal.Equals("Y"))
                    {
                        string sJunpyoId = dt.Rows[i]["JUNPYOID"]?.ToString();
                        list.Add(sJunpyoId);
                    }
                }
            }

            string sUnitPrc = string.Empty;
            string sCarryCost = string.Empty;

            //체크 된 것이 없을 때 Break
            if (list.Count == 0)
            {
                XtraMessageBox.Show("적용하려는 행을 체크하세요.");
                return;
            }
            else
            {
                IN07001F02 frm = new IN07001F02();

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    sUnitPrc = frm.sUnitPrc;
                    sCarryCost = frm.sCarryCost;

                    double dUnitPrc = 0;
                    double dCarryCost = 0;

                    double.TryParse(sUnitPrc, out dUnitPrc);
                    double.TryParse(sCarryCost, out dCarryCost);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sJunpyoId = dt.Rows[i]["JUNPYOID"]?.ToString();

                        for (int j = 0; j < list.Count; j++)
                        {
                            string sJunpyoId2 = list[j];

                            if (sJunpyoId.Equals(sJunpyoId2))
                            {
                                dt.Rows[i]["APPLY_UNIT_PRC"] = dUnitPrc;
                                dt.Rows[i]["TRANSPORTKUMAK"] = dCarryCost;
                            }
                        }
                    }

                    BtnDataClose.PerformClick();
                }
                else
                {
                    return;
                }
            }
        }

        private void GridViewRetr_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            if (e.Column.FieldName != "CHK")
            {
                if (e.Column.OptionsColumn.AllowEdit)
                {
                    GridViewRetr.SetRowCellValue(e.RowHandle, "CHK", "Y");
                }
            }
        }

        private void RepoGridLkupDealer1_EditValueChanged(object sender, EventArgs e)
        {
            LookUpEditBase lookUpEdit = (LookUpEditBase)sender;

            string sVal = lookUpEdit.Text?.ToString();
            string sCol = GridViewRetr.FocusedColumn.FieldName;

            if(!string.IsNullOrEmpty(sCol) && sCol.Equals("ETC_DEALER_CD1"))
            {
                if (string.IsNullOrEmpty(sVal))
                {
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColEtcDealerNm1, "");
                }
                else
                {
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColEtcDealerNm1, sVal);
                }
            }
            else if(!string.IsNullOrEmpty(sCol) && sCol.Equals("ETC_DEALER_CD2"))
            {
                if (string.IsNullOrEmpty(sVal))
                {
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColEtcDealerNm2, "");
                }
                else
                {
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColEtcDealerNm2, sVal);
                }
                
            }
        }

        private void BtnEtcSave_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)GridRetr.DataSource;
            List<string> list = new List<string>();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string sVal = dt.Rows[i]["CHK"]?.ToString();
                if (!string.IsNullOrEmpty(sVal))
                {
                    if (sVal.Equals("Y"))
                    {
                        string sJunpyoId = dt.Rows[i]["JUNPYOID"]?.ToString();
                        list.Add(sJunpyoId);
                    }
                }
            }

            string sCVCOD1 = string.Empty;
            string sRK1 = string.Empty;
            string sCOST1 = string.Empty;
            string sCVCOD2 = string.Empty;
            string sRK2 = string.Empty;
            string sCOST2 = string.Empty;

            //체크 된 것이 없을 때 Break
            if (list.Count == 0)
            {
                XtraMessageBox.Show("적용하려는 행을 체크하세요.");
                return;
            }
            else
            {
                IN07001F03 frm = new IN07001F03();

                frm.Owner = this;

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    sCVCOD1 = frm.sCVCOD1;
                    sRK1 = frm.sRK1;
                    sCOST1 = frm.sCOST1;
                    sCVCOD2 = frm.sCVCOD2;
                    sRK2 = frm.sRK2;
                    sCOST2 = frm.sCOST2;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sJunpyoId = dt.Rows[i]["JUNPYOID"]?.ToString();

                        for (int j = 0; j < list.Count; j++)
                        {
                            string sJunpyoId2 = list[j];

                            if (sJunpyoId.Equals(sJunpyoId2))
                            {
                                if (!string.IsNullOrEmpty(sCVCOD1))
                                    dt.Rows[i]["ETC_DEALER_CD1"] = sCVCOD1;
                                if (!string.IsNullOrEmpty(sRK1))
                                    dt.Rows[i]["ETC_COST1"] = sCOST1;
                                if (!string.IsNullOrEmpty(sCOST1))
                                    dt.Rows[i]["ETC_REMARK1"] = sRK1;
                                if (!string.IsNullOrEmpty(sCVCOD2))
                                    dt.Rows[i]["ETC_DEALER_CD2"] = sCVCOD2;
                                if (!string.IsNullOrEmpty(sRK2))
                                    dt.Rows[i]["ETC_COST2"] = sCOST2;
                                if (!string.IsNullOrEmpty(sCOST2))
                                    dt.Rows[i]["ETC_REMARK2"] = sRK2;
                            }
                        }
                    }

                    BtnDataClose.PerformClick();
                }
                else
                {
                    return;
                }
            }
        }
    }
}