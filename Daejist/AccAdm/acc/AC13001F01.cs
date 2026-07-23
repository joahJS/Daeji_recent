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
using System.Data.SqlClient;
/*
* 작성일자 : 모름
* 작성자 : 고혜성
* ---------------------HISTORY-----------------------
* 
* 수정일자 : 2021-02-25 ~ 2021-02-26
* 수정자   : 고혜성
* 수정내용 : (현업요청)
*            1. 그리드 폰트 설정
*            2. 레이아웃 전체 저장 설정
*/
namespace AccAdm
{
    public partial class AC13001F01 : DevExpress.XtraEditors.XtraForm
    {
        public AC13001F01()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AC13001F01_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            ComnEtcFunc.SetDateFromToValue(DateEditFrom, DateEditTo);
            ComnGridFunc.SetInitGridRowColor(GridViewRetr);
            RdgbSlipYn.SelectedIndex = 2;

            DataTable dtUser = GetLookUpData("1", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupUserCd, dtUser, GridRetr, GridColMgUser, "CD", "NM", "Y");

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewRetr };
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            ComnEtcFunc.YmdFromToValuesCheck(DateEditFrom, DateEditTo);
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
            string sSlipYn = RdgbSlipYn.EditValue?.ToString();

            GridRetr.DataSource = GetEquipCostInfo(sYmdFrom, sYmdTo, sSlipYn);
        }
        
        private void BtnAddSlip_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            int[] iArrSelectedRows = GridViewRetr.GetSelectedRows();
            if(iArrSelectedRows.Length == 0)
            {
                XtraMessageBox.Show("전표발행하려는 데이터에 체크하세요.");
                return;
            }
            else if (iArrSelectedRows.Length == 1)
            {
                CheckSlipInfo();
            }
            else if (iArrSelectedRows.Length > 1)
            {
               CreateSlip(iArrSelectedRows);
            }
        }

        private void BtnCancelSlip_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            int[] iArrSelectedRows = GridViewRetr.GetSelectedRows();
            if (iArrSelectedRows.Length == 0)
            {
                XtraMessageBox.Show("전표취소하려는 데이터에 체크하세요.");
                return;
            }
            else if (iArrSelectedRows.Length >= 1)
            {
                DataRow row = GridViewRetr.GetDataRow(iArrSelectedRows[0]);
                string sMgNo = row["MG_NO"]?.ToString();

                if (XtraMessageBox.Show("설비코드 : " + sMgNo + "외 " +iArrSelectedRows.Length  + "건에 대하여 전표삭제를 진행하겠습니까?", "전표삭제여부", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    DeleteSlipInfo(iArrSelectedRows);
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void DeleteSlipInfo(int[] selectedRows)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                string sUsrCd = FmMainToolBar2.UserID;
                
                for (int i = 0; i < selectedRows.Length; i++)
                {
                    DataRow row = GridViewRetr.GetDataRow(selectedRows[i]);
                    string sMgNo = row["MG_NO"]?.ToString();
                    string sMgHisSeq = row["MG_HIS_SEQ"]?.ToString();
                    string sMakeNo = row["MAKENO"]?.ToString();
                    string sMakeNoLn = row["MAKENO_LN"]?.ToString();

                    strSql.Clear();
                    strSql.AppendLine(" DELETE FROM ACTRAN ");
                    strSql.AppendLine("       WHERE REF1 = '" + sMgNo + "' ");
                    strSql.AppendLine("         AND REF2 = '" + sMgHisSeq + "' ");
                    strSql.AppendLine("         AND REF3 = 'MAKE_EXPENSE' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    strSql.Clear();
                    strSql.AppendLine(" UPDATE EQUIP_CD_HISTORY ");
                    strSql.AppendLine("    SET SLIP_YN = 'N' ");
                    strSql.AppendLine("  WHERE MG_NO = '" + sMgNo + "' ");
                    strSql.AppendLine("    AND MG_HIS_SEQ = '" + sMgHisSeq + "' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");
                BtnRetr_Click(null, null);
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }


        #region[Query]

        private DataTable GetEquipCostInfo(string sYmdFrom, string sYmdTo, string sSlipYn)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT DISTINCT A.MG_NO ");
            strSql.AppendLine(" 	 , C.EQUIP_NM ");
            strSql.AppendLine(" 	 , A.MG_HIS_SEQ ");
            strSql.AppendLine(" 	 , A.OCCUR_DT ");
            strSql.AppendLine(" 	 , A.MG_DESC ");
            strSql.AppendLine(" 	 , A.MG_COST ");
            strSql.AppendLine(" 	 , A.MG_USER ");
            strSql.AppendLine(" 	 , A.MAKENO ");
            strSql.AppendLine(" 	 , A.MAKENO_LN ");
            strSql.AppendLine(" 	 , A.LN_ESEQ ");
            strSql.AppendLine(" 	 , A.SLIP_YN ");
            strSql.AppendLine(" 	 , B.REF1 ");
            strSql.AppendLine(" 	 , B.REF2 ");
            strSql.AppendLine(" 	 , B.REF3 ");
            strSql.AppendLine("   FROM EQUIP_CD_HISTORY A ");
            strSql.AppendLine("   LEFT OUTER JOIN ACTRAN B ");
            strSql.AppendLine("     ON A.MG_NO = B.REF1 ");
            strSql.AppendLine("    AND A.MG_HIS_SEQ = B.REF2 ");
            strSql.AppendLine("   LEFT OUTER JOIN EQUIP_CD C  ");
            strSql.AppendLine("     ON A.MG_NO = C.MG_NO ");
            strSql.AppendLine("  WHERE A.OCCUR_DT BETWEEN '" + sYmdFrom + "' AND '" + sYmdTo + "' ");
            if(!sSlipYn.Equals("ALL"))
                strSql.AppendLine("    AND A.SLIP_YN = '" + sSlipYn + "' ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        #endregion[Query]

        #region[GetLookupData]

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
            else if (sNullYn.Equals("N"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine(" UNION ALL");
            }
            else
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1")) //전표구분
            {
                strSql.AppendLine(" SELECT A.INSANO AS CD ");
                strSql.AppendLine("      , A.USRNM AS NM ");
                strSql.AppendLine("      , A.USRCD AS SEQ ");
                strSql.AppendLine("   FROM ZUSRLST A ");
            }

            if (sParam.Equals("Y"))
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
                strSql.AppendLine("     ORDER BY SEQ");
            }
            else
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
            }

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        #endregion[GetLookupData]

        private void AC13001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.F1)
            {
                BtnAddSlip_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnCancelSlip_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.Escape)
            {
            }
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                CheckSlipInfo();
            }
            else if (e.Clicks == 1)
            {
                if (GridViewRetr.IsRowSelected(e.RowHandle))
                    GridViewRetr.UnselectRow(e.RowHandle);
                else
                    GridViewRetr.SelectRow(e.RowHandle);
            }

        }

        private void CheckSlipInfo()
        {
            string sMgNo = GridViewRetr.GetFocusedRowCellValue("MG_NO")?.ToString();
            string sSeq = GridViewRetr.GetFocusedRowCellValue("MG_HIS_SEQ")?.ToString();
            string sMgHisSeq = GridViewRetr.GetFocusedRowCellValue("MG_HIS_SEQ")?.ToString();
            string sMakeNo = GridViewRetr.GetFocusedRowCellValue("MAKENO")?.ToString();
            string sMakeNoLn = GridViewRetr.GetFocusedRowCellValue("MAKENO_LN")?.ToString();
            string sLnEseq = GridViewRetr.GetFocusedRowCellValue("LN_ESEQ")?.ToString();
            string sMgCost = GridViewRetr.GetFocusedRowCellValue("MG_COST")?.ToString();
            string sOccurDt = GridViewRetr.GetFocusedRowCellValue("OCCUR_DT")?.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT CASE COUNT(1) WHEN 0 THEN 'Y' ELSE 'N' END AS ADD_YN ");
            strSql.AppendLine("   FROM ACTRAN A ");
            strSql.AppendLine("  WHERE A.REF1 = '" + sMgNo + "' ");
            strSql.AppendLine("    AND A.REF2 = '" + sSeq + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            string sResult = dt.Rows[0]["ADD_YN"]?.ToString();
            if (sResult.Equals("Y"))
            {
                AC02001F02 frm = new AC02001F02();
                frm.AddModifyGb = "ADD";
                frm.PARENT_FORM = this;
                frm.ExternalParam = new Dictionary<string, string>();
                frm.ExternalParam.Add("MG_NO", sMgNo);
                frm.ExternalParam.Add("MG_HIS_SEQ", sSeq);
                frm.ExternalParam.Add("MAKENO", sMakeNo);
                frm.ExternalParam.Add("MAKENO_LN", sMakeNoLn);
                frm.ExternalParam.Add("LN_ESEQ", sLnEseq);
                frm.ExternalParam.Add("MG_COST", sMgCost);
                frm.ExternalParam.Add("OCCUR_DT", sOccurDt);
                frm.DataRowSendEvent += new AC02001F02.SendDataHandler(GetDataRow);
                frm.Show();
                //if (frm.ShowDialog() == DialogResult.OK)
                //{
                //    BtnRetr_Click(null, null);
                //}
            }
            else
            {
                strSql.Clear();

                #region mariaDB
                //strSql.AppendLine(" SELECT TDATE ");
                //strSql.AppendLine(" 	 , ATGUB ");
                //strSql.AppendLine(" 	 , SEQNO ");
                //strSql.AppendLine("   FROM ACTRAN A  ");
                //strSql.AppendLine("  WHERE A.REF1 = '" + sMgNo + "' ");
                //strSql.AppendLine("    AND A.REF2 = '" + sSeq + "' ");
                //strSql.AppendLine("   LIMIT 1 ");
                #endregion

                strSql.AppendLine(" SELECT TOP 1 TDATE ");
                strSql.AppendLine(" 	 , ATGUB ");
                strSql.AppendLine(" 	 , SEQNO ");
                strSql.AppendLine("   FROM ACTRAN A  ");
                strSql.AppendLine("  WHERE A.REF1 = '" + sMgNo + "' ");
                strSql.AppendLine("    AND A.REF2 = '" + sSeq + "' ");

                DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                AC02001F02 frm = new AC02001F02();
                frm.PARENT_FORM = this;
                frm.ExternalParam = new Dictionary<string, string>();
                frm.DrSlipInfo = dt2.Rows[0];
                frm.AddModifyGb = "MODIFY";
                frm.ExternalParam.Add("MG_NO", sMgNo);
                frm.ExternalParam.Add("MG_HIS_SEQ", sSeq);
                frm.ExternalParam.Add("MAKENO", sMakeNo);
                frm.ExternalParam.Add("MAKENO_LN", sMakeNoLn);
                frm.ExternalParam.Add("LN_ESEQ", sLnEseq);
                frm.ExternalParam.Add("MG_COST", sMgCost);
                frm.DataRowSendEvent += new AC02001F02.SendDataHandler(GetDataRow);
                frm.Show();

                //if (frm.ShowDialog() == DialogResult.OK)
                //{
                //    BtnRetr_Click(null, null);
                //}
            }
        }

        public void GetDataRow(Dictionary<string, string> sGb)
        {
            BtnRetr.PerformClick();
        }

        public string ACCOD, ACNAM;
        private void CreateSlip(int[] iArrRowHandle)
        {
            string sAcCod = string.Empty;
            string sAcNam = string.Empty;

            AC13001F02 frm = new AC13001F02();
            frm.AC13001F01 = this;

            if(frm.ShowDialog() == DialogResult.OK)
            {
                sAcCod = ACCOD;
                sAcNam = ACNAM;
            }
            else
            {
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                string sUsrCd = FmMainToolBar2.UserID;

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT B.EMP_NM ");
                strSql.AppendLine("      , B.EMP_ID ");
                strSql.AppendLine(" 	 , D.DEPT_CD ");
                strSql.AppendLine("   FROM ZUSRLST A ");
                strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
                strSql.AppendLine("     ON B.EMP_ID = A.INSANO ");
                strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD D ");
                strSql.AppendLine("     ON D.DEPT_CD = B.REAL_DUTY_DEPT ");
                strSql.AppendLine("  WHERE A.USRCD = '" + sUsrCd + "' ");

                DataTable dtUser = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                string sEmpId = dtUser.Rows[0]["EMP_ID"]?.ToString();
                string sDeptCd = dtUser.Rows[0]["DEPT_CD"]?.ToString();
                
                string sSeqNo = string.Empty;
                double dSeqNo = 1;
                for (int i = 0; i <iArrRowHandle.Length; i++)
                {
                    DataRow row = GridViewRetr.GetDataRow(iArrRowHandle[i]);
                    string sMgNo = row["MG_NO"]?.ToString();
                    string sMgHisSeq = row["MG_HIS_SEQ"]?.ToString();
                    string sMakeNo = row["MAKENO"]?.ToString();
                    string sMakeNoLn = row["MAKENO_LN"]?.ToString();
                    string sOccurDt = row["OCCUR_DT"]?.ToString();
                    string sAtext = row["MG_DESC"]?.ToString();
                    string sRK = string.Empty;
                    if (!string.IsNullOrEmpty(sAtext))
                    {
                        sRK = sAtext.Split(',')[1];
                        sAtext = sAtext.Split(',')[0];
                    }
                    
                    if (string.IsNullOrEmpty(sOccurDt))
                        sOccurDt = DateTime.Now.ToString();

                    //ACTRAN_SEQNO 채번
                    if (i == 0)
                    {
                        strSql.Clear();
                        strSql.AppendLine(" SELECT CASE WHEN MAX(SEQNO) IS NULL THEN '0001' ");
                        strSql.AppendLine("             WHEN MAX(SEQNO) IS NOT NULL THEN CAST(REPLICATE(0, 4 - LEN(MAX(SEQNO))) + (MAX(SEQNO)+1) AS VARCHAR) END ");
                        strSql.AppendLine("   FROM ACTRAN A ");
                        strSql.AppendLine("  WHERE TDATE = REPLACE(CAST('"+ sOccurDt + "' AS CHAR(10)), '-', '') ");
                        strSql.AppendLine("    AND ATGUB NOT IN('3', '4')  ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        object obj = cmd.ExecuteScalar();
                        sSeqNo = obj == null ? "0001" : Convert.ToString(obj);
                        dSeqNo = Convert.ToDouble(sSeqNo);
                        sSeqNo = sSeqNo.PadLeft(4, '0');
                    }
                    else
                    {
                        sSeqNo = dSeqNo.ToString().PadLeft(4, '0');
                    }

                    dSeqNo++;

                    strSql.Clear();
                    strSql.AppendLine(" INSERT INTO ACTRAN  ");
                    strSql.AppendLine("           ( TDATE, ATGUB, SEQNO, LINNO ");
                    strSql.AppendLine("           , ACCOD, ACNAM, CVCOD, CVNAM ");
                    strSql.AppendLine("           , ATEXT, ADAMT, ACAMT, ADPCD ");
                    strSql.AppendLine("           , REF1 , REF2, REF3, RK ");
                    strSql.AppendLine("           , CUSER, CDATE ) ");
                    strSql.AppendLine(" SELECT REPLACE(CAST('" + sOccurDt + "' AS CHAR(10)), '-', '') AS TDATE ");
                    strSql.AppendLine(" 	 , '1' --HARDCODING(출금, CD_GB=AC02001_01) ");
                    strSql.AppendLine(" 	 , '" + sSeqNo + "' ");
                    strSql.AppendLine(" 	 , 1 ");
                    strSql.AppendLine(" 	 , '" + sAcCod + "' ");
                    strSql.AppendLine(" 	 , '" + sAcNam + "' ");
                    strSql.AppendLine(" 	 , C.DEALER_CD AS CVCOD ");
                    strSql.AppendLine(" 	 , MAX(B.ECVNAM) AS CVNAM  ");
                    strSql.AppendLine("      , '"+ sAtext + "' AS ATEXT ");
                    strSql.AppendLine("      , SUM(B.EAMT) AS ADAMT ");
                    strSql.AppendLine("      , 0 ");
                    strSql.AppendLine("      , '"+ sDeptCd + "' AS ADPCD ");
                    strSql.AppendLine("      , '" + sMgNo + "' ");
                    strSql.AppendLine("      , '" + sMgHisSeq + "' ");
                    strSql.AppendLine("      , 'MAKE_EXPENSE' ");
                    strSql.AppendLine("      , '"+ sRK + "' AS REMARK ");
                    strSql.AppendLine("      , '" + sEmpId + "' ");
                    strSql.AppendLine("      , CONVERT(VARCHAR(20), GETDATE(),20)  ");
                    strSql.AppendLine("   FROM EQUIP_CD_HISTORY A ");
                    strSql.AppendLine("   LEFT OUTER JOIN MAKE_EXPENSE B ");
                    strSql.AppendLine("     ON A.MAKENO = B.MAKENO ");
                    strSql.AppendLine("    AND A.MAKENO_LN = B.MAKENO_LN ");
                    strSql.AppendLine("   LEFT OUTER JOIN ACC_DEALER_CD C ");
                    strSql.AppendLine("     ON B.ECVNAM = C.DEALER_NM ");
                    strSql.AppendLine("  WHERE B.MAKENO = '" + sMakeNo + "' ");
                    strSql.AppendLine("    AND B.MAKENO_LN = '" + sMakeNoLn + "' ");
                    strSql.AppendLine("  GROUP BY B.MAKENO, B.MAKENO_LN, C.DEALER_CD ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    strSql.Clear();
                    strSql.AppendLine(" UPDATE EQUIP_CD_HISTORY ");
                    strSql.AppendLine("    SET SLIP_YN = 'Y' ");
                    strSql.AppendLine("  WHERE MG_NO = '" + sMgNo + "' ");
                    strSql.AppendLine("    AND MG_HIS_SEQ = '" + sMgHisSeq + "' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }
                
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");
                BtnRetr_Click(null, null);
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (e.FocusedRowHandle < 0)
                return;

            //string sSlipYn = GridViewRetr.GetFocusedRowCellValue("SLIP_YN")?.ToString();
            //if (sSlipYn.Equals("Y"))
            //    BtnCancelSlip.Enabled = false;
            //else
            //    BtnCancelSlip.Enabled = true;
        }

        //선택시
        private void GridViewRetr_SelectionChanged(object sender, DevExpress.Data.SelectionChangedEventArgs e)
        {
            if (e.ControllerRow < 0)
                return;

            if (GridViewRetr.IsRowSelected(e.ControllerRow) == false)
                return;

            string sMakeNo = GridViewRetr.GetDataRow(e.ControllerRow)["MAKENO"]?.ToString();
            string sMgNo= GridViewRetr.GetDataRow(e.ControllerRow)["MG_NO"]?.ToString();
            string sMgHisSeq = GridViewRetr.GetDataRow(e.ControllerRow)["MG_HIS_SEQ"]?.ToString();

            if (string.IsNullOrEmpty(sMakeNo)){
                XtraMessageBox.Show("설비코드 : " + sMgNo + "\r\n행번 : " + sMgHisSeq + "\r\n해당 정보는 생산일보데이터가 아닙니다. 개별등록하세요.");
                GridViewRetr.UnselectRow(e.ControllerRow);
                return;
            }
            //GridViewRetr.UnselectRow(e.ControllerRow);
        }

        private void AC13001F01_TextChanged(object sender, EventArgs e)
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

        private void DateEditTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        //라디오버튼 변경 시 값에 따라 버튼 보이기
        private void RdgbSlipYn_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sVal = RdgbSlipYn.EditValue?.ToString();
            if (!string.IsNullOrEmpty(sVal))
            {
                if (sVal.Equals("ALL"))
                {
                    LayBtnSlipAdd.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    LayBtnSlipCancle.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                }
                else if (sVal.Equals("Y"))
                {
                    LayBtnSlipAdd.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                    LayBtnSlipCancle.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                }
                else if (sVal.Equals("N"))
                {
                    LayBtnSlipAdd.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                    LayBtnSlipCancle.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                }
            }

            BtnRetr_Click(null, null);
        }
    }
}