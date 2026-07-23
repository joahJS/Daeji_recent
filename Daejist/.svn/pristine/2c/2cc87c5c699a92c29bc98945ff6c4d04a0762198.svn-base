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
using System.Diagnostics;
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
    public partial class ProdPlanMgt : DevExpress.XtraEditors.XtraForm
    {
        public static string s_Dept;
        public ProdPlanMgt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void ProdPlanMgt_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.SetDateFromToValue(DateEditFrom, DateEditTo);

            

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "ProdPlanMgt");

            DataTable dtEmpId = GetLookUpData("1", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupEmpId, dtEmpId, GridRetr, GridColProdEmp, "CD", "NM", "");

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

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT EMP_ID AS CD ");
                strSql.AppendLine("      , EMP_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY EMP_NM) AS SEQ ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A");
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

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            string sYmFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);

            if(string.IsNullOrEmpty(sYmFrom))
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show("조회일자를 설정해주세요.");
                return;
            }

            string sUsrCd = rowUserInfo["USRCD"]?.ToString();

            if (string.IsNullOrEmpty(sUsrCd))
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show("접속정보가 없습니다.");
                return;
            }
            
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT B.INSANO ");
            strSql.AppendLine("   FROM ZPGMAUT A ");
            strSql.AppendLine("   LEFT OUTER JOIN ZUSRLST B  ");
            strSql.AppendLine("     ON B.USRCD = A.USRCD ");
            strSql.AppendLine("  WHERE A.USRCD = '" + sUsrCd + "' ");

            DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            string sEmpCd = dtChk.Rows[0]["INSANO"]?.ToString();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.REAL_DUTY_DEPT ");
            strSql.AppendLine("   FROM HR_EMP_BASIS A ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpCd + "' ");
            strSql.AppendLine("    AND EMPL_GB = 'Y' ");
            DataTable dtDept = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            string sDept = dtDept.Rows[0]["REAL_DUTY_DEPT"]?.ToString();
            s_Dept = sDept;

            if (string.IsNullOrEmpty(sEmpCd))
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show("해당 사용자에 대한 인사번호(EMP_CD)가 존재하지 않습니다.");
                return;
            }

            strSql.Clear();
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.MDATE ");
            strSql.AppendLine("      , A.MUSER_ID");
            if (sDept == "4100") { strSql.AppendLine("      , B.SIGN1a AS TEAM_LEADER_APRV "); }
            //if (sDept == "4150") { strSql.AppendLine("      , B.SIGN1a AS TEAM_LEADER_APRV "); }
            if (sDept == "4200") { strSql.AppendLine("      , B.SIGN2 AS TEAM_LEADER_APRV "); }
            if (sDept == "4300") { strSql.AppendLine("      , B.SIGN3 AS TEAM_LEADER_APRV "); }
            strSql.AppendLine("      , B.SIGN4 AS DEPT_MANAGER_APRV ");
            strSql.AppendLine("      , B.SIGN5 AS REP_APRV ");
            strSql.AppendLine("   FROM MAKE_M A ");
            strSql.AppendLine("   LEFT OUTER JOIN MAKE_S B ");
            strSql.AppendLine("     ON B.MDATE = A.MDATE");
            strSql.AppendLine("  WHERE A.MDATE >= '" + sYmFrom + "' ");
            strSql.AppendLine("    AND A.MDATE <= '" + sYmTo + "' ");
            strSql.AppendLine("    AND A.MUSER_ID = '" + sEmpCd + "' ");
            strSql.AppendLine("  ORDER BY MDATE ");
            //strSql.AppendLine("    AND B.GUBUN = '1' ");
            strSql.AppendLine(" ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;
            Cursor = Cursors.Default;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void ProdPlanMgt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnProdRpt_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                
            }
            else if (e.KeyCode == Keys.F4)
            {
                BtnProdDel_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnExcel_Click(null, null);
            }
        }

        private void GridViewRetr_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            string sValue = Convert.ToString(e.Value);
            if (e.Column.FieldName.Equals("MDATE"))
            {
                if (sValue.Length == 8)
                {
                    e.DisplayText = sValue.Substring(0, 4) + "-" + sValue.Substring(4, 2) + "-" + sValue.Substring(6, 2);
                }
            }
        }
        
        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                StringBuilder strSql = new StringBuilder();
                string sUserCode = rowUserInfo["USRCD"]?.ToString();
                string sWorkYmd = GridViewRetr.GetFocusedRowCellValue("MDATE")?.ToString();
                if (!string.IsNullOrEmpty(sUserCode))
                {
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT B.EMP_NM ");
                    strSql.AppendLine("      , B.EMP_ID ");
                    strSql.AppendLine(" 	 , C.COM_NM ");
                    strSql.AppendLine(" 	 , D.DEPT_CD ");
                    strSql.AppendLine(" 	 , A.USRCD ");
                    strSql.AppendLine("   FROM ZUSRLST A ");
                    strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
                    strSql.AppendLine("     ON B.EMP_ID = A.INSANO ");
                    strSql.AppendLine("   LEFT OUTER JOIN COM_BASE_CD C ");
                    strSql.AppendLine("     ON C.COM_CD = B.GRADE_CD ");
                    strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD D ");
                    strSql.AppendLine("     ON D.DEPT_CD = B.REAL_DUTY_DEPT ");
                    strSql.AppendLine("  WHERE A.USRCD = '" + sUserCode + "' ");

                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    string sEmpCd = dt.Rows[0]["EMP_ID"]?.ToString();

                    if (dt != null)
                    {
                        if (dt.Rows.Count > 0)
                        {
                            ProdMgtReport frm = new ProdMgtReport();
                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" SELECT A.REAL_DUTY_DEPT ");
                            strSql.AppendLine("   FROM HR_EMP_BASIS A ");
                            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpCd + "' ");
                            strSql.AppendLine("    AND EMPL_GB = 'Y' ");
                            DataTable dtDept = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                            string sDept = dtDept.Rows[0]["REAL_DUTY_DEPT"]?.ToString();
                            

                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" SELECT A.MDATE ");
                            if (sDept == "4100") { strSql.AppendLine("      , A.SIGN1a  "); }
                            //if (sDept == "4150") { strSql.AppendLine("      , A.SIGN1a  "); }
                            if (sDept == "4200") { strSql.AppendLine("      , A.SIGN2  "); }
                            if (sDept == "4300") { strSql.AppendLine("      , A.SIGN3  "); }
                            strSql.AppendLine("   FROM MAKE_S A ");
                            strSql.AppendLine("  WHERE MDATE = '" + sWorkYmd + "' ");

                            DataTable MakeSChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                            if (sDept == "4100")
                            {
                                if (MakeSChk.Rows.Count > 0 && MakeSChk.Rows[0]["SIGN1a"].ToString().Equals("Y"))
                                {
                                    XtraMessageBox.Show("선택하신 작업정보건은 결제승인완료 상태이므로 조회만 가능합니다.");
                                    frm.MdiYN = false;
                                }
                                else
                                {
                                    frm.MdiYN = true;
                                }
                            }
                            //if (sDept == "4150")
                            //{
                            //    if (MakeSChk.Rows.Count > 0 && MakeSChk.Rows[0]["SIGN1a"].ToString().Equals("Y"))
                            //    {
                            //        XtraMessageBox.Show("선택하신 작업정보건은 결제승인완료 상태이므로 조회만 가능합니다.");
                            //        frm.MdiYN = false;
                            //    }
                            //    else
                            //    {
                            //        frm.MdiYN = true;
                            //    }
                            //}
                            if (sDept == "4200")
                            {
                                if (MakeSChk.Rows.Count > 0 && MakeSChk.Rows[0]["SIGN2"].ToString().Equals("Y"))
                                {
                                    XtraMessageBox.Show("선택하신 작업정보건은 결제승인완료 상태이므로 조회만 가능합니다.");
                                    frm.MdiYN = false;
                                }
                                else
                                {
                                    frm.MdiYN = true;
                                }
                            }
                            if (sDept == "4300")
                            {
                                if (MakeSChk.Rows.Count > 0 && MakeSChk.Rows[0]["SIGN3"].ToString().Equals("Y"))
                                {
                                    XtraMessageBox.Show("선택하신 작업정보건은 결제승인완료 상태이므로 조회만 가능합니다.");
                                    frm.MdiYN = false;
                                }
                                else
                                {
                                    frm.MdiYN = true;
                                }
                            }

                            DataRow drInfo = dt.Rows[0];
                            string sWorkDate = GridViewRetr.GetFocusedRowCellValue("MDATE")?.ToString().Replace("-", "").Substring(0, 8);
                            string sMakeNO = GridViewRetr.GetFocusedRowCellValue("MAKENO")?.ToString();
                            
                            frm.drUserInfo = drInfo;
                            frm.dMakeNo = Convert.ToDouble(sMakeNO);
                            frm.sMakeNo = sMakeNO;
                            frm.sYmd = sWorkDate;
                            frm.sReadOnly = true;
                            frm.MdiParent = this.MdiParent;
                            frm.ADD_MODIFY_GB = "MOD";
                            frm.Show();
                        }
                        else
                        {
                            XtraMessageBox.Show("로그인 사용자 정보가 없습니다.");
                            return;
                        }
                    }
                }
            }
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }


        public string RESULT_GB;
        public double RESULT_MAKENO;
        public string RESULT_YMD;
        private void BtnProdRpt_Click(object sender, EventArgs e)
        {
            StringBuilder strSql = new StringBuilder();
            string sUserCode = rowUserInfo["USRCD"]?.ToString();

            if (!string.IsNullOrEmpty(sUserCode))
            {
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT B.EMP_NM ");
                strSql.AppendLine("      , B.EMP_ID ");
                strSql.AppendLine(" 	 , C.COM_NM ");
                strSql.AppendLine(" 	 , D.DEPT_CD ");
                strSql.AppendLine(" 	 , A.USRCD ");
                strSql.AppendLine("   FROM ZUSRLST A ");
                strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
                strSql.AppendLine("     ON B.EMP_ID = A.INSANO ");
                strSql.AppendLine("   LEFT OUTER JOIN COM_BASE_CD C ");
                strSql.AppendLine("     ON C.COM_CD = B.GRADE_CD ");
                strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD D ");
                strSql.AppendLine("     ON D.DEPT_CD = B.REAL_DUTY_DEPT ");
                strSql.AppendLine("  WHERE A.USRCD = '" + sUserCode + "' ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        string sEmpId = dt.Rows[0]["EMP_ID"]?.ToString();

                        //strSql.Clear();
                        //strSql.AppendLine(" SELECT CASE WHEN COUNT(1) > 0 THEN 'Y' ELSE 'N' END AS EXIST_YN ");
                        //strSql.AppendLine("   FROM MAKE_M  ");
                        //strSql.AppendLine("  WHERE MDATE = REPLACE(CURDATE(), '-', '') ");
                        //strSql.AppendLine("    AND MUSER_ID = '" + sEmpId + "' ");

                        //DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                        //string sYn = dt2.Rows[0]["EXIST_YN"]?.ToString();
                        //if (sYn.Equals("Y"))
                        //{
                        //    XtraMessageBox.Show("금일 생산일보 내역이 존재합니다.");
                        //    return;
                        //}

                        if (string.IsNullOrEmpty(sEmpId))
                        {
                            XtraMessageBox.Show("등록된 인사정보가 없습니다.\r\n인사등록후 다시 시도해주세요.");
                            return;
                        }


                        PD01001F01 frm_1 = new PD01001F01();
                        frm_1.USRCD = sEmpId;
                        frm_1.PARENT_FORM = this;
                        if (frm_1.ShowDialog() == DialogResult.OK)
                        {
                            //PD01001F01에 주석처리 참조
                            if (RESULT_GB.Equals("A")) //미결재(수정가능)
                            {
                                Cursor = Cursors.WaitCursor;
                                ProdMgtReport frm = new ProdMgtReport();
                                frm.drUserInfo = dt.Rows[0];
                                frm.dMakeNo = RESULT_MAKENO;
                                frm.sYmd = RESULT_YMD;
                                frm.MdiParent = this.MdiParent;
                                frm.MdiYN = true;
                                frm.sReadOnly = true;
                                frm.ADD_MODIFY_GB = "MOD";
                                frm.Show();

                                Cursor = Cursors.Default;
                            }
                            else if (RESULT_GB.Equals("B")) //결재승인
                            {
                                Cursor = Cursors.WaitCursor;
                                ProdMgtReport frm = new ProdMgtReport();
                                frm.drUserInfo = dt.Rows[0];
                                frm.dMakeNo = RESULT_MAKENO;
                                frm.sYmd = RESULT_YMD;
                                frm.MdiParent = this.MdiParent;
                                frm.MdiYN = false;
                                frm.sReadOnly = true;
                                frm.ADD_MODIFY_GB = "MOD";
                                frm.Show();
                                Cursor = Cursors.Default;
                            }
                            else if (RESULT_GB.Equals("C")) //생산정보 미존재 시 채번 후 등록
                            {
                                Cursor = Cursors.WaitCursor;
                                //채번
                                strSql.Clear();
                                
                                strSql.AppendLine("SELECT CASE WHEN MAX(MAKENO) IS NOT NULL THEN MAX(MAKENO) + 1 ELSE CAST(CONCAT(SUBSTRING(CONVERT(CHAR(8),GETDATE(),112),1,6),'001')AS NUMERIC) END AS MAKENO");
                                strSql.AppendLine("  FROM MAKE_M A");
                                strSql.AppendLine(" WHERE SUBSTRING(CONVERT(VARCHAR,CONVERT(NUMERIC,MAKENO)), 1, 6) = SUBSTRING(CONVERT(CHAR(8), GETDATE(), 112), 1, 6)");

                                DataTable dt2  = null;
                                dt2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                                double Result_MakeNo = Convert.ToDouble(dt2.Rows[0]["MAKENO"]);

                                //채번된 번호로 Insert 진행
                                if (!InsertMakeM(Result_MakeNo, FmMainToolBar2.UserID, RESULT_YMD))
                                {
                                    Cursor = Cursors.Default;
                                    return;
                                }
                                

                                ProdMgtReport frm = new ProdMgtReport();
                                frm.drUserInfo = dt.Rows[0];
                                frm.dMakeNo = Result_MakeNo;
                                frm.sYmd = RESULT_YMD;
                                frm.MdiParent = this.MdiParent;
                                frm.MdiYN = true;
                                frm.sReadOnly = false;
                                frm.ADD_MODIFY_GB = "ADD";
                                frm.Show();
                                Cursor = Cursors.Default;
                            }

                            BtnRetr.PerformClick();
                        }
                    }
                    else
                    {
                        XtraMessageBox.Show("로그인 사용자 정보가 없습니다.");
                        return;
                    }
                }
            }
            
        }

        private bool InsertMakeM(double dMakeNo, string sUsrCd, string sYmd)
        {
            StringBuilder strSql = new StringBuilder();

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;
                
                strSql.Clear();
                strSql.AppendLine(" INSERT INTO MAKE_M ");
                strSql.AppendLine("           ( MAKENO ");
                strSql.AppendLine("           , MDATE ");
                strSql.AppendLine("           , MUSER_ID ");
                strSql.AppendLine("           , MUSER ");
                strSql.AppendLine("           , ENT_DT ");
                strSql.AppendLine("           , ENT_ID) ");
                strSql.AppendLine("     VALUES(  ");
                strSql.AppendLine("             @MAKENO ");
                strSql.AppendLine("           , '" + sYmd + "' ");
                strSql.AppendLine("           , ( SELECT A.EMP_ID");
                strSql.AppendLine("                 FROM HR_EMP_BASIS A");
                strSql.AppendLine("                 LEFT OUTER JOIN ZUSRLST B");
                strSql.AppendLine("                   ON A.EMP_ID = B.INSANO ");
                strSql.AppendLine("                WHERE B.USRCD = @USRCD ) ");
                strSql.AppendLine("           , ( SELECT A.EMP_NM");
                strSql.AppendLine("                 FROM HR_EMP_BASIS A");
                strSql.AppendLine("                 LEFT OUTER JOIN ZUSRLST B");
                strSql.AppendLine("                   ON A.EMP_ID = B.INSANO ");
                strSql.AppendLine("                WHERE B.USRCD = @USRCD ) ");
                strSql.AppendLine("           , CONVERT([varchar](20),getdate(),(21)) ");
                strSql.AppendLine("           , @USRCD) ");
                
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.Parameters.AddWithValue("@MAKENO", dMakeNo);
                cmd.Parameters.AddWithValue("@USRCD", FmMainToolBar2.UserID);
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                //XtraMessageBox.Show("저장을 완료했습니다.");

                return true;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);

                return false;
            }
            
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
                string sFileNM = "생산일보 리스트";
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

        private void BtnProdDel_Click(object sender, EventArgs e)
        {
            string sWorkDate = GridViewRetr.GetFocusedRowCellValue("MDATE")?.ToString().Replace("-", "").Substring(0, 8);
            string sMakeNO = GridViewRetr.GetFocusedRowCellValue("MAKENO")?.ToString();
            string sWorkDate2 = DateTime.ParseExact(sWorkDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture).ToString().Substring(0, 10);

            if (MessageBox.Show(sWorkDate2 + "의 생산일보를 삭제하시겠습니까?","",MessageBoxButtons.YesNo) == DialogResult.No)
            {
                return;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT CASE WHEN COUNT(*) > 0 THEN 'Y' ELSE 'N' END AS YN ");
            strSql.AppendLine("   FROM MAKE_S  ");
            strSql.AppendLine("  WHERE MDATE = '" + sWorkDate + "' --개수 체크 ");
            strSql.AppendLine(" UNION ALL ");
            if (s_Dept == "4100") { strSql.AppendLine(" SELECT CASE SIGN1a WHEN 'Y' THEN 'Y' ELSE 'N' END AS YN "); }
            //if (s_Dept == "4150") { strSql.AppendLine(" SELECT CASE SIGN1a WHEN 'Y' THEN 'Y' ELSE 'N' END AS YN "); }
            if (s_Dept == "4200") { strSql.AppendLine(" SELECT CASE SIGN2 WHEN 'Y' THEN 'Y' ELSE 'N' END AS YN "); }
            if (s_Dept == "4300") { strSql.AppendLine(" SELECT CASE SIGN3 WHEN 'Y' THEN 'Y' ELSE 'N' END AS YN "); }
            strSql.AppendLine("   FROM MAKE_S ");
            strSql.AppendLine("  WHERE MDATE = '" + sWorkDate + "' --결재부분 체크 ");
            
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //MAKE_S 테이블에 해당 데이터 존재체크
            if (dt.Rows[0]["YN"].ToString().Equals("Y"))
            {
                //검색일자에 결재건 체크
                if (dt.Rows[1]["YN"].ToString().Equals("Y"))
                {
                    XtraMessageBox.Show("해당 데이터는 결재완료 건이므로 삭제할 수 없습니다.");
                    return;
                }
            }

            strSql.Clear();
            #region MARIADB
            //strSql.AppendLine(" SELECT COUNT(IF(A.SLIP_YN='Y', A.SLIP_YN, NULL)) AS CNT ");
            //strSql.AppendLine("   FROM EQUIP_CD_HISTORY A ");
            //strSql.AppendLine("   LEFT JOIN MAKE_EXPENSE B ");
            //strSql.AppendLine("     ON A.`MAKENO` = B.MAKENO ");
            //strSql.AppendLine("    AND A.MAKENO_LN = B.MAKENO_LN ");
            //strSql.AppendLine("    AND A.LN_ESEQ = B.LN_ESEQ ");
            //strSql.AppendLine("  WHERE B.MAKENO = " + sMakeNO + " ");
            #endregion

            strSql.AppendLine(" SELECT COUNT(CASE WHEN A.SLIP_YN='Y' THEN 1 END) AS CNT ");
            strSql.AppendLine("   FROM EQUIP_CD_HISTORY A ");
            strSql.AppendLine("   LEFT JOIN MAKE_EXPENSE B ");
            strSql.AppendLine("     ON A.MAKENO = B.MAKENO ");
            strSql.AppendLine("    AND A.MAKENO_LN = B.MAKENO_LN ");
            strSql.AppendLine("    AND A.LN_ESEQ = B.LN_ESEQ ");
            strSql.AppendLine("  WHERE B.MAKENO = " + sMakeNO + " ");

            dt = null;
            dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            int iCnt = Convert.ToInt32(dt.Rows[0]["CNT"]);
            if(iCnt > 0)
            {
                XtraMessageBox.Show(string.Format("해당 생산에서 등록한 비용에 대하여 전표가 발행되어 삭제할 수 없습니다.\r\n참조 : 생산번호 -> {0}", sMakeNO));
                return;
            }

            string sUserCode = rowUserInfo["USRCD"]?.ToString();

            if (!string.IsNullOrEmpty(sUserCode))
            {
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT B.EMP_NM ");
                strSql.AppendLine("      , B.EMP_ID ");
                strSql.AppendLine(" 	 , C.COM_NM ");
                strSql.AppendLine(" 	 , D.DEPT_CD ");
                strSql.AppendLine(" 	 , A.USRCD ");
                strSql.AppendLine("   FROM ZUSRLST A ");
                strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
                strSql.AppendLine("     ON B.EMP_ID = A.INSANO ");
                strSql.AppendLine("   LEFT OUTER JOIN COM_BASE_CD C ");
                strSql.AppendLine("     ON C.COM_CD = B.GRADE_CD ");
                strSql.AppendLine("   LEFT OUTER JOIN ACC_DEPT_CD D ");
                strSql.AppendLine("     ON D.DEPT_CD = B.REAL_DUTY_DEPT ");
                strSql.AppendLine("  WHERE A.USRCD = '" + sUserCode + "' ");

                DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt2 != null)
                {
                    if (dt2.Rows.Count > 0)
                    {
                        try
                        {
                            Cursor = Cursors.WaitCursor;

                            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                            SqlCommand cmd = DBConn.dbCon.CreateCommand();
                            cmd.Transaction = DBConn.dbTran;

                            string sEmpId = dt2.Rows[0]["EMP_ID"].ToString();

                            Console.WriteLine("MakeNO: " + sMakeNO + "// WorkDate: " + sWorkDate + "// EMPID: " + sEmpId);

                            //make_m 삭제
                            strSql.Clear();
                            strSql.AppendLine(" DELETE ");
                            strSql.AppendLine("   FROM MAKE_M");
                            strSql.AppendLine("  WHERE MAKENO = " + sMakeNO);
                            strSql.AppendLine("    AND MDATE = '" + sWorkDate + "'");
                            strSql.AppendLine("    AND MUSER_ID = '" + sEmpId + "'");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();

                            ////make_s 삭제
                            //strSql.Clear();
                            //strSql.AppendLine(" DELETE ");
                            //strSql.AppendLine("   FROM MAKE_S");
                            //strSql.AppendLine("  WHERE MDATE = '" + sWorkDate + "'");
                            //strSql.AppendLine("    AND GUBUN = '1'");

                            //cmd.CommandType = CommandType.Text;
                            //cmd.CommandText = strSql.ToString();
                            //cmd.ExecuteNonQuery();

                            //make_1 삭제
                            strSql.Clear();
                            strSql.AppendLine(" DELETE ");
                            strSql.AppendLine("   FROM MAKE_1");
                            strSql.AppendLine("  WHERE MAKENO = " + sMakeNO);

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();

                            //make_2 삭제
                            strSql.Clear();
                            strSql.AppendLine(" DELETE ");
                            strSql.AppendLine("   FROM MAKE_2");
                            strSql.AppendLine("  WHERE MAKENO = " + sMakeNO);

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();

                            //make_3 삭제
                            strSql.Clear();
                            strSql.AppendLine(" DELETE ");
                            strSql.AppendLine("   FROM MAKE_3");
                            strSql.AppendLine("  WHERE MAKENO = " + sMakeNO);

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();

                            //make_4 삭제
                            strSql.Clear();
                            strSql.AppendLine(" DELETE ");
                            strSql.AppendLine("   FROM MAKE_4");
                            strSql.AppendLine("  WHERE MAKENO = " + sMakeNO);

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();

                            //make_5 삭제
                            strSql.Clear();
                            strSql.AppendLine(" DELETE ");
                            strSql.AppendLine("   FROM MAKE_5");
                            strSql.AppendLine("  WHERE MAKENO = " + sMakeNO);

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();

                            //make_6 삭제
                            strSql.Clear();
                            strSql.AppendLine(" DELETE ");
                            strSql.AppendLine("   FROM MAKE_6");
                            strSql.AppendLine("  WHERE MAKENO = " + sMakeNO);

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();

                            //make_7 삭제
                            strSql.Clear();
                            strSql.AppendLine(" DELETE ");
                            strSql.AppendLine("   FROM MAKE_7");
                            strSql.AppendLine("  WHERE MAKENO = " + sMakeNO);

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();

                            //make_expense 삭제
                            strSql.Clear();
                            strSql.AppendLine(" DELETE ");
                            strSql.AppendLine("   FROM MAKE_EXPENSE");
                            strSql.AppendLine("  WHERE MAKENO = " + sMakeNO);

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();

                            //equip_cd_history 삭제
                            strSql.Clear();
                            strSql.AppendLine(" DELETE ");
                            strSql.AppendLine("   FROM EQUIP_CD_HISTORY");
                            strSql.AppendLine("  WHERE MAKENO = " + sMakeNO);
                            strSql.AppendLine("    AND MG_USER = '"+ sEmpId + "'");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();


                            DBConn.dbTran.Commit();
                            DBConn.dbTran = null;
                            MessageBox.Show("삭제를 완료했습니다.");

                            this.DialogResult = DialogResult.OK;

                            BtnRetr_Click(null, null);

                            Cursor = Cursors.Default;
                        }
                        catch (Exception ex)
                        {
                            Cursor = Cursors.Default;

                            DBConn.dbTran.Rollback();
                            DBConn.dbTran = null;
                            MessageBox.Show(ex.Message);
                        }
                    }
                    else
                    {
                        XtraMessageBox.Show("로그인 사용자 정보가 없습니다.");
                        return;
                    }
                }
            }
        }

        private void ProdPlanMgt_TextChanged(object sender, EventArgs e)
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

        private void GridRetr_Click(object sender, EventArgs e)
        {

        }
    }
}