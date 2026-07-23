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
using DevExpress.XtraEditors.Repository;
using System.Net;
using DevExpress.XtraGrid.Views.Grid;
using System.Data.SqlClient;
using System.IO;
using System.Security.Cryptography;
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
*            
* 수정일자 : 2021-08-06
* 수정자   : 정은영
* ID       : #00001
* 수정내용 : (현업요청)
*           1. 자동채번 부서코드별로 최대값으로 채번
*           
* 수정일자 : 2021-09-08
* 수정자 : 정은영
* ID : #00002
* 수정내용: (현업요청)
*         1. 기본정보 입력 체크
*         2. 검색조건 사원명으로 변경
*/
namespace AccAdm
{
    public partial class AccPersonnelRecordsDev : DevExpress.XtraEditors.XtraForm
    {
        public AccPersonnelRecordsDev()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AccPersonnelRecordsDev_Load(object sender, EventArgs e)
        {
            //
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            //
            arrGrdView = new GridView[] { GridViewRetr, GridViewFamily, GridViewFamily, GridViewCert, GridViewEdu, GridViewCareer };
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

            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            ComnEtcFunc.gp_SetColorFocused(layoutControl8);
            ComnEtcFunc.gp_SetColorFocused(layoutControl9);
            ComnEtcFunc.gp_SetColorFocused(layoutControl10);

            DataTable dtDeptCD = GetLookUpData("1", "3", "", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupEditDeptCd, dtDeptCD, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupDeptCd, dtDeptCD, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupRealDutyDept, dtDeptCD, "CD", "NM", "Y");
            DataTable dtRetireCd = GetLookUpData("2", "", "", "");
            ComLib.ComGrid.SetLookUpEdit(LkupRetireResnCd, dtRetireCd, "CD", "NM", "Y");
            DataTable dtGradeCd = GetLookUpData("3", "", "", "");
            ComLib.ComGrid.SetLookUpEdit(LkupGradeCd, dtGradeCd, "CD", "NM", "Y");
            DataTable dtDeptCDGrid = GetLookUpData("1", "", "", "");
            RepositoryItemGridLookUpEdit DeptLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(DeptLkup, dtDeptCD, GridRetr, GridColDeptCd, "CD", "NM", "");
            DataTable dtBloodType = GetLookUpData("8", "3", "", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupBloodType, dtBloodType, "CD", "NM", "Y");

            DataTable dtFamilyRelationGb = GetLookUpData("4", "3", "", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoFamilyGLkupFamlRelatCd, dtFamilyRelationGb, GridFamily, GridColFamilyFamlRelationCd, "CD", "NM", "");

            DataTable dtFamilyAcademicGb = GetLookUpData("5", "3", "", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoFamilyGLkupAcademicGb, dtFamilyAcademicGb, GridFamily, GridColFamilyAcademicGb, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoEduGLkupAcademicGb, dtFamilyAcademicGb, GridEdu, GridColEduAcademicGb, "CD", "NM", "");

            DataTable dtReligion = GetLookUpData("6", "3", "", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupReligion, dtReligion, "CD", "NM", "Y");

            DataTable dtGrduGb = GetLookUpData("7", "3", "", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoEduGLkupGrduGb, dtGrduGb, GridEdu, GridColEduGrduGb, "CD", "NM", "");

            DataTable dtBankCd = GetLookUpData("9", "3", "", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupDealBankCd, dtBankCd, "CD", "NM", "Y");

            BtnRetr_Click(null, null);
        }

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam, string sOther)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("1"))
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else if (sNullYn.Equals("2"))
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine(" UNION ALL");
            }
            else if (sNullYn.Equals("3"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.DEPT_CD AS CD");
                strSql.AppendLine("      , A.DEPT_NM AS NM");
                strSql.AppendLine("      , A.DEPT_CD AS SEQ");
                strSql.AppendLine("   FROM ACC_DEPT_CD A");
                strSql.AppendLine("  WHERE USE_YN = 'Y' ");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'RETIRE_RESN_CD'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'GRADE_CD'");
            }
            else if(sGb.Equals("4"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'FAML_RELATION_CD'");
            }
            else if (sGb.Equals("5"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'ACADEMIC_GB'");
            }
            else if(sGb.Equals("6"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE CD_GB = 'RELIGION_CD'");
            }
            else if (sGb.Equals("7"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE CD_GB = 'GRDU_GB' ");
            }
            else if (sGb.Equals("8"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE CD_GB = 'BLOOD_TYPE' ");
            }
            else if (sGb.Equals("9"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE CD_GB = 'BANK_CD' ");
            }

            if (sOther.Equals("Y"))
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
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            GetGridRetr();
            GetBasisInfo();

            if (XtTControl.SelectedTabPage.Name.Equals("XtTPerson"))
            {
                GetPersonalInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPCarrer"))
            {
                GetCareerInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPCert"))
            {
                GetCertificationInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPFamily"))
            {
                GetFamilyInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPEdctn"))
            {
                GetEducationInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPPay"))
            {
                GetPaymentInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPMltrSvc"))
            {
                GetMilitaryInfo();
            }
        }

        private void GetGridRetr()
        {
            Cursor = Cursors.WaitCursor;

            string sSaNm = TxtSaNm.EditValue?.ToString();
            //string sDeptCd = LkupEditDeptCd.EditValue?.ToString();
            string sEmplGb = RdgbRetrGb.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine("SELECT EMP_ID ");
            strSql.AppendLine("     , EMP_NM  ");
            strSql.AppendLine("     , ENG_NM ");
            strSql.AppendLine("     , CHN_NM ");
            strSql.AppendLine("     , ENTRANCE_YMD ");
            strSql.AppendLine("     , RETIRE_YMD ");
            strSql.AppendLine("     , DEPT_CD ");
            strSql.AppendLine("     , REAL_DUTY_DEPT ");
            strSql.AppendLine("     , MANAGER_ID  ");
            strSql.AppendLine("     , JOBKIND_CD  ");
            strSql.AppendLine("     , JOBPOSITION_CD  ");
            strSql.AppendLine("     , GRADE_CD   ");
            strSql.AppendLine("     , JOBDUTY_CD ");
            strSql.AppendLine("     , PAYSTEP_CD ");
            strSql.AppendLine("     , GENDER_CD  ");
            strSql.AppendLine("     , IDT_NO  ");
            strSql.AppendLine("     , CNTR_GB ");
            strSql.AppendLine("     , RETIRE_RESN  ");
            strSql.AppendLine("     , RETIRE_RESN_CD ");
            strSql.AppendLine("     , ENTRANCE_GB  ");
            strSql.AppendLine("     , EMPL_GB ");
            strSql.AppendLine("     , CHG_BFR_ID1 ");
            strSql.AppendLine("     , CHG_BFR_ID2 ");
            strSql.AppendLine("     , FOREIGNER_YN ");
            strSql.AppendLine("     , ENT_DT");
            strSql.AppendLine("     , ENT_ID");
            strSql.AppendLine("     , MFY_DT");
            strSql.AppendLine("     , MFY_ID");
            strSql.AppendLine("     , GDAYYN");
            strSql.AppendLine("  FROM HR_EMP_BASIS");
            strSql.AppendLine(" WHERE 1=1 ");
            //if (!string.IsNullOrEmpty(sDeptCd))
            //{
            //    strSql.AppendLine("  AND DEPT_CD = '" + sDeptCd + "' ");
            //}
            //#00002
            if (!string.IsNullOrEmpty(sSaNm))
            {
                strSql.AppendLine("  AND EMP_NM LIKE '%" + sSaNm + "%' ");
            }
            if(sEmplGb != "A" || string.IsNullOrEmpty(sEmplGb)) strSql.AppendLine("  AND EMPL_GB = '" + sEmplGb + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            //#00002
            //if (dt.Rows.Count > 0)
            //{
            //    GridRetr.DataSource = dt;
            //}
            GridRetr.DataSource = dt;
            Cursor = Cursors.Default;
        }

        #region[인사 기본자료 Select, Save]

        private void GetBasisInfo()
        {
            Cursor = Cursors.WaitCursor;

            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.EMP_ID ");
            strSql.AppendLine("      , A.EMP_NM ");
            strSql.AppendLine("      , A.ENG_NM ");
            strSql.AppendLine("      , A.CHN_NM ");
            strSql.AppendLine("      , A.ENTRANCE_YMD ");
            strSql.AppendLine("      , A.RETIRE_YMD ");
            strSql.AppendLine("      , A.DEPT_CD ");
            strSql.AppendLine("      , A.REAL_DUTY_DEPT ");
            strSql.AppendLine("      , A.MANAGER_ID ");
            strSql.AppendLine("      , A.JOBKIND_CD ");
            strSql.AppendLine("      , A.JOBPOSITION_CD ");
            strSql.AppendLine("      , A.GRADE_CD ");
            strSql.AppendLine("      , A.JOBDUTY_CD ");
            strSql.AppendLine("      , A.PAYSTEP_CD ");
            strSql.AppendLine("      , A.GENDER_CD ");
            strSql.AppendLine("      , A.IDT_NO ");
            strSql.AppendLine("      , A.CNTR_GB ");
            strSql.AppendLine("      , A.RETIRE_RESN ");
            strSql.AppendLine("      , A.RETIRE_RESN_CD ");
            strSql.AppendLine("      , A.ENTRANCE_GB ");
            strSql.AppendLine("      , A.EMPL_GB ");
            strSql.AppendLine("      , A.CHG_BFR_ID1 ");
            strSql.AppendLine("      , A.CHG_BFR_ID2 ");
            strSql.AppendLine("      , A.FOREIGNER_YN ");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , A.ENT_ID ");
            strSql.AppendLine("      , A.MFY_DT ");
            strSql.AppendLine("      , A.MFY_ID ");
            strSql.AppendLine("      , A.SINIMG");
            strSql.AppendLine("      , A.GDAYYN");
            strSql.AppendLine("      , A.JMUNYN");
            strSql.AppendLine("   FROM HR_EMP_BASIS A ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if(dt.Rows.Count != 0)
            {
                TxtEMPId.EditValue = dt.Rows[0]["EMP_ID"].ToString();
                TxtEmpNm.EditValue = dt.Rows[0]["EMP_NM"].ToString();
                TxtEngNm.EditValue = dt.Rows[0]["ENG_NM"].ToString();
                TxtChnNm.EditValue = dt.Rows[0]["CHN_NM"].ToString();

                if(dt.Rows[0]["ENTRANCE_YMD"].ToString().Length == 8)
                {
                    string sTemp = dt.Rows[0]["ENTRANCE_YMD"].ToString();
                    string sResult = sTemp.Substring(0, 4) + "-" + sTemp.Substring(4, 2) + "-" + sTemp.Substring(6, 2);
                    DtpEntranceYmd.EditValue = sResult;
                }

                if (dt.Rows[0]["RETIRE_YMD"].ToString().Length == 8)
                {
                    string sTemp = dt.Rows[0]["RETIRE_YMD"].ToString();
                    string sResult = sTemp.Substring(0, 4) + "-" + sTemp.Substring(4, 2) + "-" + sTemp.Substring(6, 2);
                    DtpRetireYmd.EditValue = sResult;
                }

                LkupDeptCd.EditValue = dt.Rows[0]["DEPT_CD"].ToString(); 
                LkupRealDutyDept.EditValue = dt.Rows[0]["REAL_DUTY_DEPT"].ToString(); 
                TxtJobkindCd.EditValue = dt.Rows[0]["JOBKIND_CD"].ToString(); 
                LkupGradeCd.EditValue = dt.Rows[0]["GRADE_CD"].ToString(); 
                TxtPayStepCd.EditValue = dt.Rows[0]["PAYSTEP_CD"].ToString(); 
                RdgbGenderCd.EditValue = dt.Rows[0]["GENDER_CD"].ToString(); 
                string sIdtNo = dt.Rows[0]["IDT_NO"].ToString();
                if (!string.IsNullOrEmpty(sIdtNo))
                {
                    sIdtNo= ComnEtcFunc.Decrypt(sIdtNo, ComnEtcFunc._SECRET_KEY2);
                }
                TxtIdtNo.EditValue = sIdtNo; 
                LkupRetireResnCd.EditValue = dt.Rows[0]["RETIRE_RESN_CD"].ToString(); 
                TxtRetireResn.EditValue = dt.Rows[0]["RETIRE_RESN"].ToString(); 
                RdgbEmplGb.EditValue = dt.Rows[0]["EMPL_GB"].ToString(); 
                RdgbForeignerYn.EditValue = dt.Rows[0]["FOREIGNER_YN"].ToString();
                RadiGDAYYN.EditValue = dt.Rows[0]["GDAYYN"]?.ToString();
                RadiJmonYn.EditValue = dt.Rows[0]["JMUNYN"];

                if (!DBNull.Value.Equals(dt.Rows[0]["SINIMG"]))
                {
                    picImg.EditValue = (byte[])dt.Rows[0]["SINIMG"];
                }
                else
                {
                    picImg.EditValue = null;
                }
            }
            else
            {
                //TxtEMPId.EditValue = GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();
                TxtEmpNm.EditValue = null;
                TxtEngNm.EditValue = null;
                TxtChnNm.EditValue = null;
                DtpEntranceYmd.EditValue = null;
                DtpRetireYmd.EditValue = null;
                LkupDeptCd.EditValue = null;
                LkupRealDutyDept.EditValue = null;
                TxtJobkindCd.EditValue = null;
                LkupGradeCd.EditValue = null;
                TxtPayStepCd.EditValue = null;
                RdgbGenderCd.EditValue = null;
                TxtIdtNo.EditValue = null;
                LkupRetireResnCd.EditValue = null;
                TxtRetireResn.EditValue = null;
                RdgbEmplGb.EditValue = null;
                RdgbForeignerYn.EditValue =null;
                RadiGDAYYN.EditValue = null;
                RadiJmonYn.EditValue = null;
            }

            Cursor = Cursors.Default;
        }

        private bool SaveBasisInfo(SqlCommand cmd)
        {
            string sEmpId = TxtEMPId.EditValue.ToString();
            string sEmpNm = TxtEmpNm.EditValue == null ? string.Empty : TxtEmpNm.EditValue.ToString();
            string sEngNm = TxtEngNm.EditValue == null ? string.Empty : TxtEngNm.EditValue.ToString();
            string sChnNm = TxtChnNm.EditValue == null ? string.Empty : TxtChnNm.EditValue.ToString();

            string sEntranceYmd = string.Empty;
            if (DtpEntranceYmd.EditValue != null)
            {
                sEntranceYmd = DtpEntranceYmd.EditValue.ToString().Replace("-", "").Substring(0, 8);
            }

            string sRetireYmd = string.Empty;
            if (DtpRetireYmd.EditValue != null)
            {
                sRetireYmd = DtpRetireYmd.EditValue.ToString().Replace("-", "").Substring(0, 8);
            }

            string sDeptCd = LkupDeptCd.EditValue == null ? string.Empty : LkupDeptCd.EditValue.ToString();
            string sRealDutyDept = LkupRealDutyDept.EditValue == null ? string.Empty : LkupRealDutyDept.EditValue.ToString();
            string sJobKindCd = TxtJobkindCd.EditValue == null ? string.Empty : TxtJobkindCd.EditValue.ToString();
            string sGradeCd = LkupGradeCd.EditValue == null ? string.Empty : LkupGradeCd.EditValue.ToString();
            string sPayStepCd = TxtPayStepCd.EditValue == null ? string.Empty : TxtPayStepCd.EditValue.ToString();
            string sGenderCd = RdgbGenderCd.EditValue == null ? string.Empty : RdgbGenderCd.EditValue.ToString();
            string sIdtNo = TxtIdtNo.EditValue == null ? string.Empty : TxtIdtNo.EditValue.ToString();
            if (!string.IsNullOrEmpty(sIdtNo))
            {
                sIdtNo = ComnEtcFunc.Encrypt(sIdtNo, ComnEtcFunc._SECRET_KEY2);
            }
            
            string sRetireResn = TxtRetireResn.EditValue == null ? string.Empty : TxtRetireResn.EditValue.ToString();
            string sRetireResnCd = LkupRetireResnCd.EditValue == null ? string.Empty : LkupRetireResnCd.EditValue.ToString();
            string sEmplGb = RdgbEmplGb.EditValue == null ? string.Empty : RdgbEmplGb.EditValue.ToString();
            string sForeignerYn = RdgbForeignerYn.EditValue == null ? string.Empty : RdgbForeignerYn.EditValue.ToString();
            string sGDAYYN = RadiGDAYYN.EditValue?.ToString();
            string sJMUNYN = RadiJmonYn.EditValue?.ToString();
            string sEntId = FmMainToolBar2.UserID;
            string sMfyId = FmMainToolBar2.UserID;

            string sEntranceGb = string.Empty;
            string sManagerId = string.Empty;
            string sJobPositionCd = string.Empty;
            string sJobDutyCd = string.Empty;
            string sCntrGb = string.Empty;
            string sChgBfrId1 = string.Empty;
            string sChgBfrId2 = string.Empty;

            /*
             * #00002
             */
            if (string.IsNullOrEmpty(sEmpId))
            {
                XtraMessageBox.Show("사번을 입력해주세요.");

                TxtEMPId.Focus();
                return false;
            }
            if (string.IsNullOrEmpty(sEmpNm))
            {
                XtraMessageBox.Show("성명을 입력해주세요.");

                TxtEmpNm.Focus();
                return false;
            }
            if (string.IsNullOrEmpty(sDeptCd))
            {
                XtraMessageBox.Show("부서코드를 입력해주세요.");

                LkupDeptCd.Focus();
                return false;
            }
            if (string.IsNullOrEmpty(sEntranceYmd))
            {
                XtraMessageBox.Show("입사일자를 입력해주세요.");

                DtpEntranceYmd.Focus();
                return false;
            }
            if (string.IsNullOrEmpty(sGenderCd))
            {
                XtraMessageBox.Show("성별을 입력해주세요.");

                RdgbGenderCd.Focus();
                return false;
            }
            if (string.IsNullOrEmpty(sEmplGb))
            {
                XtraMessageBox.Show("재직구분을 입력해주세요.");

                RdgbEmplGb.Focus();
                return false;
            }


            StringBuilder strSql = new StringBuilder();

            /*
             * #00001 : 채번 중복확인
             */
            
            //신규인 경우 빈칸
            string sEmpidGrid = GridViewRetr.GetFocusedRowCellValue("EMP_ID")?.ToString();

            if (string.IsNullOrEmpty(sEmpidGrid))
            {
                strSql.Clear();
                strSql.AppendLine("SELECT COUNT(1) AS CNT");
                strSql.AppendLine("  FROM HR_EMP_BASIS A");
                strSql.AppendLine(" WHERE EMP_ID = '" + sEmpId + "'");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt != null && double.Parse(dt.Rows[0]["CNT"].ToString()) > 0)
                {
                    XtraMessageBox.Show("중복된 사원번호는 사용할 수 없습니다.");
                    TxtEMPId.Focus();
                    TxtEMPId.SelectAll();
                    return false;
                }
            }

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine("IF EXISTS(SELECT * FROM HR_EMP_BASIS WHERE EMP_ID = '" + sEmpId + "')");
            strSql.AppendLine("   BEGIN");
            strSql.AppendLine("         UPDATE HR_EMP_BASIS");
            strSql.AppendLine("            SET EMP_NM = '" + sEmpNm + "' ");
			strSql.AppendLine("              , ENG_NM = '" + sEngNm + "' ");
			strSql.AppendLine("              , CHN_NM = '" + sChnNm + "' ");
			strSql.AppendLine("              , ENTRANCE_YMD = '" + sEntranceYmd + "' ");
			strSql.AppendLine("              , RETIRE_YMD = '" + sRetireYmd + "' ");
			strSql.AppendLine("              , DEPT_CD = '" + sDeptCd + "' ");
			strSql.AppendLine("              , REAL_DUTY_DEPT = '" + sRealDutyDept + "' ");
			strSql.AppendLine("              , MANAGER_ID = '" + sManagerId + "' ");
			strSql.AppendLine("              , JOBKIND_CD = '" + sJobKindCd + "' ");
			strSql.AppendLine("              , JOBPOSITION_CD = '" + sJobPositionCd + "' ");
			strSql.AppendLine("              , GRADE_CD = '" + sGradeCd + "' ");
			strSql.AppendLine("              , JOBDUTY_CD = '" + sJobDutyCd + "' ");
			strSql.AppendLine("              , PAYSTEP_CD = '" + sPayStepCd + "' ");
			strSql.AppendLine("              , GENDER_CD = '" + sGenderCd + "' ");
			strSql.AppendLine("              , IDT_NO = '" + sIdtNo + "' ");
			strSql.AppendLine("              , CNTR_GB = '" + sCntrGb + "' ");
			strSql.AppendLine("              , RETIRE_RESN = '" + sRetireResn + "' ");
			strSql.AppendLine("              , RETIRE_RESN_CD = '" + sRetireResnCd + "' ");
			strSql.AppendLine("              , ENTRANCE_GB = '" + sEntranceGb + "' ");
			strSql.AppendLine("              , EMPL_GB = '" + sEmplGb + "' ");
			strSql.AppendLine("              , CHG_BFR_ID1 = '" + sChgBfrId1 + "' ");
			strSql.AppendLine("              , CHG_BFR_ID2 = '" + sChgBfrId2 + "' ");
			strSql.AppendLine("              , FOREIGNER_YN = '" + sForeignerYn + "' ");
            strSql.AppendLine("              , GDAYYN = '"+ sGDAYYN + "'");
            strSql.AppendLine("              , JMUNYN = '" + sJMUNYN + "'");
			strSql.AppendLine("              , MFY_DT = convert(varchar(19),getdate(),20) ");
			strSql.AppendLine("              , MFY_ID = '" + sEntId + "' ");
            if (_SIGNIMG != null)
            {
                strSql.AppendLine("          , SINIMG = @SINIMG");
            }
            strSql.AppendLine("          WHERE EMP_ID = '" + sEmpId + "'");
            strSql.AppendLine("     END");
            strSql.AppendLine("ELSE    ");
            strSql.AppendLine("   BEGIN");
            strSql.AppendLine("         INSERT INTO HR_EMP_BASIS ");
            strSql.AppendLine("              (EMP_ID ");
            strSql.AppendLine("              , EMP_NM ");
            strSql.AppendLine("              , ENG_NM ");
            strSql.AppendLine("              , CHN_NM ");
            strSql.AppendLine("              , ENTRANCE_YMD ");
            strSql.AppendLine("              , RETIRE_YMD ");
            strSql.AppendLine("              , DEPT_CD ");
            strSql.AppendLine("              , REAL_DUTY_DEPT ");
            strSql.AppendLine("              , MANAGER_ID ");
            strSql.AppendLine("              , JOBKIND_CD ");
            strSql.AppendLine("              , JOBPOSITION_CD ");
            strSql.AppendLine("              , GRADE_CD ");
            strSql.AppendLine("              , JOBDUTY_CD ");
            strSql.AppendLine("              , PAYSTEP_CD ");
            strSql.AppendLine("              , GENDER_CD ");
            strSql.AppendLine("              , IDT_NO ");
            strSql.AppendLine("              , CNTR_GB ");
            strSql.AppendLine("              , RETIRE_RESN ");
            strSql.AppendLine("              , RETIRE_RESN_CD ");
            strSql.AppendLine("              , ENTRANCE_GB ");
            strSql.AppendLine("              , EMPL_GB ");
            strSql.AppendLine("              , CHG_BFR_ID1 ");
            strSql.AppendLine("              , CHG_BFR_ID2 ");
            strSql.AppendLine("              , FOREIGNER_YN ");
            strSql.AppendLine("              , GDAYYN");
            strSql.AppendLine("              , JMUNYN");
            strSql.AppendLine("              , ENT_DT ");
            strSql.AppendLine("              , ENT_ID ");
            if (_SIGNIMG != null)
            {
                strSql.AppendLine(", SINIMG");
            }
            strSql.AppendLine("              ) ");
            strSql.AppendLine("         VALUES ");
            strSql.AppendLine("              ('" + sEmpId + "' ");
            strSql.AppendLine("              , '" + sEmpNm + "' ");
            strSql.AppendLine("              , '" + sEngNm + "' ");
            strSql.AppendLine("              , '" + sChnNm + "' ");
            strSql.AppendLine("              , '" + sEntranceYmd + "' ");
            strSql.AppendLine("              , '" + sRetireYmd + "' ");
            strSql.AppendLine("              , '" + sDeptCd + "' ");
            strSql.AppendLine("              , '" + sRealDutyDept + "' ");
            strSql.AppendLine("              , '" + sManagerId + "' ");
            strSql.AppendLine("              , '" + sJobKindCd + "' ");
            strSql.AppendLine("              , '" + sJobPositionCd + "' ");
            strSql.AppendLine("              , '" + sGradeCd + "' ");
            strSql.AppendLine("              , '" + sJobDutyCd + "' ");
            strSql.AppendLine("              , '" + sPayStepCd + "' ");
            strSql.AppendLine("              , '" + sGenderCd + "' ");
            strSql.AppendLine("              , '" + sIdtNo + "' ");
            strSql.AppendLine("              , '" + sCntrGb + "' ");
            strSql.AppendLine("              , '" + sRetireResn + "' ");
            strSql.AppendLine("              , '" + sRetireResnCd + "' ");
            strSql.AppendLine("              , '" + sEntranceGb + "' ");
            strSql.AppendLine("              , '" + sEmplGb + "' ");
            strSql.AppendLine("              , '" + sChgBfrId1 + "' ");
            strSql.AppendLine("              , '" + sChgBfrId2 + "' ");
            strSql.AppendLine("              , '" + sForeignerYn + "' ");
            strSql.AppendLine("              , '" + sGDAYYN + "'");
            strSql.AppendLine("              , '" + sJMUNYN + "'");
            strSql.AppendLine("              , convert(varchar(19),getdate(),20) ");
            strSql.AppendLine("              , '" + sEntId + "' ");
            if (_SIGNIMG != null)
            {
                strSql.AppendLine(", @SINIMG");
            }
            strSql.AppendLine("              ) ");
            strSql.AppendLine("     END");

            #region mariaDB
            //strSql.AppendLine(" INSERT INTO HR_EMP_BASIS ");
            //strSql.AppendLine("           ( EMP_ID ");
            //strSql.AppendLine("           , EMP_NM ");
            //strSql.AppendLine("           , ENG_NM ");
            //strSql.AppendLine("           , CHN_NM ");
            //strSql.AppendLine("           , ENTRANCE_YMD ");
            //strSql.AppendLine("           , RETIRE_YMD ");
            //strSql.AppendLine("           , DEPT_CD ");
            //strSql.AppendLine("           , REAL_DUTY_DEPT ");
            //strSql.AppendLine("           , MANAGER_ID ");
            //strSql.AppendLine("           , JOBKIND_CD ");
            //strSql.AppendLine("           , JOBPOSITION_CD ");
            //strSql.AppendLine("           , GRADE_CD ");
            //strSql.AppendLine("           , JOBDUTY_CD ");
            //strSql.AppendLine("           , PAYSTEP_CD ");
            //strSql.AppendLine("           , GENDER_CD ");
            //strSql.AppendLine("           , IDT_NO ");
            //strSql.AppendLine("           , CNTR_GB ");
            //strSql.AppendLine("           , RETIRE_RESN ");
            //strSql.AppendLine("           , RETIRE_RESN_CD ");
            //strSql.AppendLine("           , ENTRANCE_GB ");
            //strSql.AppendLine("           , EMPL_GB ");
            //strSql.AppendLine("           , CHG_BFR_ID1 ");
            //strSql.AppendLine("           , CHG_BFR_ID2 ");
            //strSql.AppendLine("           , FOREIGNER_YN ");
            //strSql.AppendLine("           , ENT_DT ");
            //strSql.AppendLine("           , ENT_ID ");
            //strSql.AppendLine("           , MFY_DT ");
            //strSql.AppendLine("           , MFY_ID ");
            //strSql.AppendLine("           ) ");
            //strSql.AppendLine("      VALUES ");
            //strSql.AppendLine("           ( '" + sEmpId + "' ");
            //strSql.AppendLine("           , '" + sEmpNm + "' ");
            //strSql.AppendLine("           , '" + sEngNm + "' ");
            //strSql.AppendLine("           , '" + sChnNm + "' ");
            //strSql.AppendLine("           , '" + sEntranceYmd + "' ");
            //strSql.AppendLine("           , '" + sRetireYmd + "' ");
            //strSql.AppendLine("           , '" + sDeptCd + "' ");
            //strSql.AppendLine("           , '" + sRealDutyDept + "' ");
            //strSql.AppendLine("           , '" + sManagerId + "' ");
            //strSql.AppendLine("           , '" + sJobKindCd + "' ");
            //strSql.AppendLine("           , '" + sJobPositionCd + "' ");
            //strSql.AppendLine("           , '" + sGradeCd + "' ");
            //strSql.AppendLine("           , '" + sJobDutyCd + "' ");
            //strSql.AppendLine("           , '" + sPayStepCd + "' ");
            //strSql.AppendLine("           , '" + sGenderCd + "' ");
            //strSql.AppendLine("           , '" + sIdtNo + "' ");
            //strSql.AppendLine("           , '" + sCntrGb + "' ");
            //strSql.AppendLine("           , '" + sRetireResn + "' ");
            //strSql.AppendLine("           , '" + sRetireResnCd + "' ");
            //strSql.AppendLine("           , '" + sEntranceGb + "' ");
            //strSql.AppendLine("           , '" + sEmplGb + "' ");
            //strSql.AppendLine("           , '" + sChgBfrId1 + "' ");
            //strSql.AppendLine("           , '" + sChgBfrId2 + "' ");
            //strSql.AppendLine("           , '" + sForeignerYn + "' ");
            //strSql.AppendLine("           , NOW() ");
            //strSql.AppendLine("           , '" + sEntId + "' ");
            //strSql.AppendLine("           , NOW() ");
            //strSql.AppendLine("           , '" + sMfyId + "' ");
            //strSql.AppendLine("           ) ");
            //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
            //strSql.AppendLine("             EMP_NM = '" + sEmpNm + "' ");
            //strSql.AppendLine("           , ENG_NM = '" + sEngNm + "' ");
            //strSql.AppendLine("           , CHN_NM = '" + sChnNm + "' ");
            //strSql.AppendLine("           , ENTRANCE_YMD = '" + sEntranceYmd + "' ");
            //strSql.AppendLine("           , RETIRE_YMD = '" + sRetireYmd + "' ");
            //strSql.AppendLine("           , DEPT_CD = '" + sDeptCd + "' ");
            //strSql.AppendLine("           , REAL_DUTY_DEPT = '" + sRealDutyDept + "' ");
            //strSql.AppendLine("           , MANAGER_ID = '" + sManagerId + "' ");
            //strSql.AppendLine("           , JOBKIND_CD = '" + sJobKindCd + "' ");
            //strSql.AppendLine("           , JOBPOSITION_CD = '" + sJobPositionCd + "' ");
            //strSql.AppendLine("           , GRADE_CD = '" + sGradeCd + "' ");
            //strSql.AppendLine("           , JOBDUTY_CD = '" + sJobDutyCd + "' ");
            //strSql.AppendLine("           , PAYSTEP_CD = '" + sPayStepCd + "' ");
            //strSql.AppendLine("           , GENDER_CD = '" + sGenderCd + "' ");
            //strSql.AppendLine("           , IDT_NO = '" + sIdtNo + "' ");
            //strSql.AppendLine("           , CNTR_GB = '" + sCntrGb + "' ");
            //strSql.AppendLine("           , RETIRE_RESN = '" + sRetireResn + "' ");
            //strSql.AppendLine("           , RETIRE_RESN_CD = '" + sRetireResnCd + "' ");
            //strSql.AppendLine("           , ENTRANCE_GB = '" + sEntranceGb + "' ");
            //strSql.AppendLine("           , EMPL_GB = '" + sEmplGb + "' ");
            //strSql.AppendLine("           , CHG_BFR_ID1 = '" + sChgBfrId1 + "' ");
            //strSql.AppendLine("           , CHG_BFR_ID2 = '" + sChgBfrId2 + "' ");
            //strSql.AppendLine("           , FOREIGNER_YN = '" + sForeignerYn + "' ");
            //strSql.AppendLine("           , MFY_DT = NOW() ");
            //strSql.AppendLine("           , MFY_ID = '" + sEntId + "' ");
            #endregion

            if (_SIGNIMG != null)
            {
                SqlParameter paramImg = new SqlParameter("@SINIMG", SqlDbType.VarBinary);
                paramImg.Value = _SIGNIMG;

                cmd.Parameters.Add(paramImg);
            }
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSql.ToString();
            cmd.ExecuteNonQuery();

            return true;
        }


        #endregion[인사 기본자료 Select, Save]

        #region[인사 개인 정보 Select, Save]

        private void GetPersonalInfo()
        {
            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.EMP_ID ");
            strSql.AppendLine("      , A.MARRIAGE_YN ");
            strSql.AppendLine("      , A.ZIP ");
            strSql.AppendLine("      , A.ADDR ");
            strSql.AppendLine("      , A.DTL_ADDR ");
            strSql.AppendLine("      , A.TEL_NO ");
            strSql.AppendLine("      , A.HP_NO ");
            strSql.AppendLine("      , A.EMAIL ");
            strSql.AppendLine("      , A.BIRTH_YMD ");
            strSql.AppendLine("      , A.MARRIAGE_YMD ");
            strSql.AppendLine("      , A.HEIGHT ");
            strSql.AppendLine("      , A.WEIGHT ");
            strSql.AppendLine("      , A.EYESIGHT ");
            strSql.AppendLine("      , A.CLR_BLDS_YN ");
            strSql.AppendLine("      , A.BLOOD_TYPE ");
            strSql.AppendLine("      , A.HEALTH_STATE ");
            strSql.AppendLine("      , A.RELIGION_CD ");
            strSql.AppendLine("      , A.ACHL_CPCT ");
            strSql.AppendLine("      , A.HOBBY ");
            strSql.AppendLine("      , A.SPECIALTY ");
            strSql.AppendLine("      , A.PATRIOT_YN ");
            strSql.AppendLine("      , A.PATRIOT_NO ");
            strSql.AppendLine("      , A.PATRIOT_RELATION_CD ");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , A.ENT_ID ");
            strSql.AppendLine("      , A.ENT_IP ");
            strSql.AppendLine("      , A.MFY_DT ");
            strSql.AppendLine("      , A.MFY_ID ");
            strSql.AppendLine("      , A.MFY_IP ");
            strSql.AppendLine("   FROM HR_EMP_PERSONAL A ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt.Rows.Count != 0)
            {
                TxtEmployeeId.EditValue = dt.Rows[0]["EMP_ID"].ToString();
                RdgbMarryYn.EditValue = dt.Rows[0]["MARRIAGE_YN"].ToString();
                TxtZipCode.EditValue = dt.Rows[0]["ZIP"].ToString();
                TxtAddr.EditValue = dt.Rows[0]["ADDR"].ToString();
                TxtDtlAddr.EditValue = dt.Rows[0]["DTL_ADDR"].ToString();
                TxtTelNo.EditValue = dt.Rows[0]["TEL_NO"].ToString();
                TxtPhoneNo.EditValue = dt.Rows[0]["HP_NO"].ToString();
                TxtEmail.EditValue = dt.Rows[0]["EMAIL"].ToString();
                
                if(dt.Rows[0]["BIRTH_YMD"].ToString().Length == 8)
                {
                    string sTemp = dt.Rows[0]["BIRTH_YMD"].ToString();
                    string sResult = sTemp.Substring(0, 4) + "-" + sTemp.Substring(4, 2) + "-" + sTemp.Substring(6, 2);
                    DateEditBirth.EditValue = sResult;
                }

                if (dt.Rows[0]["MARRIAGE_YMD"].ToString().Length == 8)
                {
                    string sTemp = dt.Rows[0]["MARRIAGE_YMD"].ToString();
                    string sResult = sTemp.Substring(0, 4) + "-" + sTemp.Substring(4, 2) + "-" + sTemp.Substring(6, 2);
                    DateEditBirth.EditValue = sResult;
                }
                
                TxtHeight.EditValue = dt.Rows[0]["HEIGHT"].ToString();
                TxtWeight.EditValue = dt.Rows[0]["WEIGHT"].ToString();
                LkupReligion.EditValue = dt.Rows[0]["RELIGION_CD"].ToString();
                TxtHobby.EditValue = dt.Rows[0]["HOBBY"].ToString();
                TxtSpecial.EditValue = dt.Rows[0]["SPECIALTY"].ToString();
                RdgbRPYn.EditValue = dt.Rows[0]["PATRIOT_YN"].ToString();
                TxtRPNo.EditValue = dt.Rows[0]["PATRIOT_NO"].ToString();
                TxtRPRelation.EditValue = dt.Rows[0]["PATRIOT_RELATION_CD"].ToString();
                TxtEyeSight.EditValue = dt.Rows[0]["EYESIGHT"].ToString();
                RdgbClrBlind.EditValue = dt.Rows[0]["CLR_BLDS_YN"].ToString();
                LkupBloodType.EditValue = dt.Rows[0]["BLOOD_TYPE"].ToString();
                TxtHealthState.EditValue = dt.Rows[0]["HEALTH_STATE"].ToString();
                TxtAlchol.EditValue = dt.Rows[0]["ACHL_CPCT"].ToString();
            }
            else
            {
                TxtEmployeeId.EditValue = TxtEMPId.EditValue.ToString();
                RdgbMarryYn.EditValue = null;
                TxtZipCode.EditValue = null;
                TxtAddr.EditValue = null;
                TxtDtlAddr.EditValue = null;
                TxtTelNo.EditValue = null;
                TxtPhoneNo.EditValue = null;
                TxtEmail.EditValue = null;
                DateEditBirth.EditValue = null;
                DateEditMarry.EditValue = null;
                TxtHeight.EditValue = null;
                TxtWeight.EditValue = null;
                LkupReligion.EditValue = null;
                TxtHobby.EditValue = null;
                TxtSpecial.EditValue = null;
                RdgbRPYn.EditValue = null;
                TxtRPNo.EditValue = null;
                TxtRPRelation.EditValue = null;
                TxtEyeSight.EditValue = null;
                RdgbClrBlind.EditValue = null;
                LkupBloodType.EditValue = null;
                TxtHealthState.EditValue = null;
                TxtAlchol.EditValue = null;
            }

            
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.EMP_ID ");
            strSql.AppendLine("      , A.SEQ ");
            strSql.AppendLine("      , A.STRT_YMD ");
            strSql.AppendLine("      , A.END_YMD ");
            strSql.AppendLine("      , A.CONTENTS ");
            strSql.AppendLine("      , A.BIGO ");
            strSql.AppendLine("   FROM HR_EMP_BIGO A ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");
            strSql.AppendLine("  ORDER BY SEQ ");

            DataTable dtBigo = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridBigo.DataSource = dtBigo;
        }

        private bool SavePersonalInfo(SqlCommand cmd)
        {
            bool result = false;
            string sEmpId = TxtEMPId.EditValue == null ? string.Empty : TxtEMPId.EditValue.ToString();
            string sMarrigeYn = RdgbMarryYn.EditValue == null ? string.Empty : RdgbMarryYn.EditValue.ToString();
            string sZip = TxtZipCode.EditValue == null ? string.Empty : TxtZipCode.EditValue.ToString();
            string sAddr = TxtAddr.EditValue == null ? string.Empty : TxtAddr.EditValue.ToString();
            string sDtlAddr = TxtDtlAddr.EditValue == null ? string.Empty : TxtDtlAddr.EditValue.ToString();
            string sTelNo = TxtTelNo.EditValue == null ? string.Empty : TxtTelNo.EditValue.ToString();
            string sHpNo = TxtPhoneNo.EditValue == null ? string.Empty : TxtPhoneNo.EditValue.ToString();
            string sEmail = TxtEmail.EditValue == null ? string.Empty : TxtEmail.EditValue.ToString();

            string sBirthYmd = string.Empty;
            if (DateEditBirth.EditValue == null)
            {
                
            }
            else
            {
                sBirthYmd = DateEditBirth.EditValue.ToString().Replace("-", "").Substring(0, 8);
            }

            string sMarrigeYmd = string.Empty;
            if (DateEditMarry.EditValue == null)
            {
                
            }
            else
            {
                sMarrigeYmd = DateEditMarry.EditValue.ToString().Replace("-", "").Substring(0, 8);
            }

            double dHeight = 0;
            string sHeight = TxtHeight.EditValue?.ToString();
            if (!string.IsNullOrEmpty(sHeight))
            {
                dHeight = Convert.ToDouble(sHeight);
            }
            else
            {
                dHeight = 0;
            }

            double dWeight = 0;
            string sWeight = TxtWeight.EditValue?.ToString();
            if (!string.IsNullOrEmpty(sWeight))
            {
                dWeight = Convert.ToDouble(sWeight);
            }
            else
            {
                dHeight = 0;
            }
            string sEyeSight = TxtEyeSight.EditValue?.ToString();
            double dEyeSight = string.IsNullOrEmpty(sEyeSight) ? 0 : Convert.ToDouble(sEyeSight);
            string sClrBldYn = RdgbClrBlind.EditValue?.ToString();
            string sBloodType = LkupBloodType.EditValue?.ToString();
            string sHealthState = TxtHealthState.EditValue?.ToString();
            string sAchlCpct = TxtAlchol.EditValue?.ToString();

            string sReligionCd = LkupReligion.EditValue == null ? string.Empty : LkupReligion.EditValue.ToString();
            string sHobby = TxtHobby.EditValue == null ? string.Empty : TxtHobby.EditValue.ToString();
            string sSpecialty = TxtSpecial.EditValue == null ? string.Empty : TxtSpecial.EditValue.ToString();
            string sPatriotYn = RdgbRPYn.EditValue == null ? string.Empty : RdgbRPYn.EditValue.ToString();
            string sPatriotNo = TxtRPNo.EditValue == null ? string.Empty : TxtRPNo.EditValue.ToString();
            string sPatriotRelationCd = TxtRPRelation.EditValue == null ? string.Empty : TxtRPRelation.EditValue.ToString();
            string sEntId = FmMainToolBar2.drUser["USRCD"]?.ToString();
            string sEntIp = Get_MyIP();
            string sMfyId = FmMainToolBar2.drUser["USRCD"]?.ToString();
            string sMfyIp = Get_MyIP();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine("IF EXISTS(SELECT * FROM HR_EMP_PERSONAL WHERE EMP_ID = '" + sEmpId + "')");
            strSql.AppendLine("   BEGIN");
            strSql.AppendLine("         UPDATE HR_EMP_PERSONAL");
            strSql.AppendLine("            SET MARRIAGE_YN = '" + sMarrigeYn   + "', ZIP = '" + sZip       + "', ADDR = '" + sAddr       + "', DTL_ADDR = '" + sDtlAddr + "' ");
			strSql.AppendLine("              , TEL_NO = '" + sTelNo       + "', HP_NO = '" + sHpNo      + "', EMAIL = '" + sEmail      + "', BIRTH_YMD = '" + sBirthYmd + "' ");
			strSql.AppendLine("              , MARRIAGE_YMD = '" + sMarrigeYmd  + "', HEIGHT = " + dHeight    + " , WEIGHT = " + dWeight      + " , EYESIGHT = " + dEyeSight + " ");
			strSql.AppendLine("              , CLR_BLDS_YN = '" + sClrBldYn    + "', BLOOD_TYPE = '" + sBloodType + "', HEALTH_STATE = '"+ sHealthState + "', RELIGION_CD = '" + sReligionCd + "' ");
			strSql.AppendLine("              , ACHL_CPCT = '" + sAchlCpct    + "', HOBBY = '" + sHobby     + "', SPECIALTY = '" + sSpecialty  + "', PATRIOT_YN = '" + sPatriotYn + "' ");
			strSql.AppendLine("              , PATRIOT_NO = '" + sPatriotNo   + "', PATRIOT_RELATION_CD = '" + sPatriotRelationCd + "' ");
			strSql.AppendLine("              , MFY_DT = convert(varchar(19),getdate(),20), MFY_ID = '" + sEntId + "', MFY_IP = '" + sEntIp + "' ");
            strSql.AppendLine("          WHERE EMP_ID = '" + sEmpId + "'");
            strSql.AppendLine("     END                                 ");
            strSql.AppendLine("ELSE                                     ");
            strSql.AppendLine("   BEGIN                                 ");
            strSql.AppendLine("         INSERT INTO HR_EMP_PERSONAL ");
            strSql.AppendLine("          (");
            strSql.AppendLine("            EMP_ID, MARRIAGE_YN, ZIP, ADDR, DTL_ADDR ");
            strSql.AppendLine("          , TEL_NO, HP_NO, EMAIL, BIRTH_YMD, MARRIAGE_YMD ");
            strSql.AppendLine("          , HEIGHT, WEIGHT, EYESIGHT, CLR_BLDS_YN, BLOOD_TYPE ");
            strSql.AppendLine("          , HEALTH_STATE, RELIGION_CD, ACHL_CPCT, HOBBY, SPECIALTY ");
            strSql.AppendLine("          , PATRIOT_YN, PATRIOT_NO, PATRIOT_RELATION_CD, ENT_DT, ENT_ID ");
            strSql.AppendLine("          , ENT_IP ");
            strSql.AppendLine("          ) ");
            strSql.AppendLine("     VALUES ");
            strSql.AppendLine("          (");
            strSql.AppendLine("            '" + sEmpId       + "', '" + sMarrigeYn  + "', '" + sZip               + "', '" + sAddr     + "', '" + sDtlAddr + "' ");
            strSql.AppendLine("          , '" + sTelNo       + "', '" + sHpNo       + "', '" + sEmail             + "', '" + sBirthYmd + "', '" + sMarrigeYmd + "' ");
            strSql.AppendLine("          , " + dHeight      + ", " + dWeight     + ", " + dEyeSight          + ", '" + sClrBldYn + "', '" + sBloodType + "' ");
            strSql.AppendLine("          , '" + sHealthState + "', '" + sReligionCd + "', '" + sAchlCpct          + "', '" + sHobby    + "', '" + sSpecialty + "' ");
            strSql.AppendLine("          , '" + sPatriotYn   + "', '" + sPatriotNo  + "', '" + sPatriotRelationCd + "', convert(varchar(19),getdate(),20), '" + sEntId + "' ");
            strSql.AppendLine("          , '" + sEntIp       + "'");
            strSql.AppendLine("          ) ");
            strSql.AppendLine("     END");

            #region mariaDB
            //strSql.AppendLine(" INSERT INTO HR_EMP_PERSONAL ");
            //strSql.AppendLine("           ( ");
            //strSql.AppendLine("             EMP_ID      , MARRIAGE_YN, ZIP                , ADDR       , DTL_ADDR ");
            //strSql.AppendLine("           , TEL_NO      , HP_NO      , EMAIL              , BIRTH_YMD  , MARRIAGE_YMD ");
            //strSql.AppendLine("           , HEIGHT      , WEIGHT     , EYESIGHT           , CLR_BLDS_YN, BLOOD_TYPE ");
            //strSql.AppendLine("           , HEALTH_STATE, RELIGION_CD, ACHL_CPCT          , HOBBY      , SPECIALTY ");
            //strSql.AppendLine("           , PATRIOT_YN  , PATRIOT_NO , PATRIOT_RELATION_CD, ENT_DT     , ENT_ID ");
            //strSql.AppendLine("           , ENT_IP      , MFY_DT     , MFY_ID             , MFY_IP ");
            //strSql.AppendLine("           ) ");
            //strSql.AppendLine("      VALUES ");
            //strSql.AppendLine("           ( ");
            //strSql.AppendLine("             '" + sEmpId       + "', '" + sMarrigeYn  + "', '" + sZip               + "', '" + sAddr     + "', '" + sDtlAddr + "' ");
            //strSql.AppendLine("           , '" + sTelNo       + "', '" + sHpNo       + "', '" + sEmail             + "', '" + sBirthYmd + "', '" + sMarrigeYmd + "' ");
            //strSql.AppendLine("           ,  " + dHeight      + " ,  " + dWeight     + " ,  " + dEyeSight          + " , '" + sClrBldYn + "', '" + sBloodType + "' ");
            //strSql.AppendLine("           , '" + sHealthState + "', '" + sReligionCd + "', '" + sAchlCpct          + "', '" + sHobby    + "', '" + sSpecialty + "' ");
            //strSql.AppendLine("           , '" + sPatriotYn   + "', '" + sPatriotNo  + "', '" + sPatriotRelationCd + "', NOW(), '" + sEntId + "' ");
            //strSql.AppendLine("           , '" + sEntIp       + "', NOW(), '" + sEntId + "', '" + sEntIp + "' ");
            //strSql.AppendLine("           ) ");
            //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
            //strSql.AppendLine("             MARRIAGE_YN  = '" + sMarrigeYn   + "', ZIP        = '" + sZip       + "', ADDR         = '" + sAddr       + "', DTL_ADDR    = '" + sDtlAddr + "' ");
            //strSql.AppendLine("           , TEL_NO       = '" + sTelNo       + "', HP_NO      = '" + sHpNo      + "', EMAIL        = '" + sEmail      + "', BIRTH_YMD   = '" + sBirthYmd + "' ");
            //strSql.AppendLine("           , MARRIAGE_YMD = '" + sMarrigeYmd  + "', HEIGHT     =  " + dHeight    + " , WEIGHT       = " + dWeight      + " , EYESIGHT    =  " + dEyeSight + " ");
            //strSql.AppendLine("           , CLR_BLDS_YN  = '" + sClrBldYn    + "', BLOOD_TYPE = '" + sBloodType + "', HEALTH_STATE = '"+ sHealthState + "', RELIGION_CD = '" + sReligionCd + "' ");
            //strSql.AppendLine("           , ACHL_CPCT    = '" + sAchlCpct    + "', HOBBY      = '" + sHobby     + "', SPECIALTY    = '" + sSpecialty  + "', PATRIOT_YN  = '" + sPatriotYn + "' ");
            //strSql.AppendLine("           , PATRIOT_NO   = '" + sPatriotNo   + "', PATRIOT_RELATION_CD = '" + sPatriotRelationCd + "' ");
            //strSql.AppendLine("           , MFY_DT       = NOW(), MFY_ID = '" + sEntId + "', MFY_IP = '" + sEntIp + "' ");
            #endregion

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSql.ToString();
            cmd.ExecuteNonQuery();

            result = true;

            return result;
        }
        
        #endregion[인사 개인 정보 Select, Save]
        
        #region [ 자격증명 부분 Select, Add, Save, Delete ]

        private void GetCertificationInfo()
        {
            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.EMP_ID ");
            strSql.AppendLine("      , A.SEQ ");
            strSql.AppendLine("      , A.CERT_GB ");
            strSql.AppendLine("      , A.CERT_NO ");
            strSql.AppendLine("      , A.CERT_NM ");
            strSql.AppendLine("      , A.ISSUE_YMD ");
            strSql.AppendLine("      , A.END_YMD ");
            strSql.AppendLine("      , A.ISSUE_FACIL ");
            strSql.AppendLine("      , A.ALW_PMNT_YN ");
            strSql.AppendLine("      , A.ALW_PMNT_RATE ");
            strSql.AppendLine("      , A.ALW_PMNT_AMT ");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.ENT_ID AS NUMERIC) IS NULL THEN A.ENT_ID ELSE DBO.FN_USRNM(A.ENT_ID) END AS ENT_ID ");
            strSql.AppendLine("      , A.ENT_IP ");
            strSql.AppendLine("      , A.MFY_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MFY_ID AS NUMERIC) IS NULL THEN A.MFY_ID ELSE DBO.FN_USRNM(A.MFY_ID) END AS MFY_ID ");
            strSql.AppendLine("      , A.MFY_IP ");
            strSql.AppendLine("   FROM HR_EMP_CERT A");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");
            strSql.AppendLine("  ORDER BY SEQ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridCert.DataSource = dt;
        }

        private void AddCertificationInfo()
        {
            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            GridViewCert.AddNewRow();
            GridViewCert.SetFocusedRowCellValue("EMP_ID", sEmpId);
            GridViewCert.SetFocusedRowCellValue("ISSUE_YMD", DateTime.Today.ToString().Substring(0, 10));
            GridViewCert.SetFocusedRowCellValue("END_YMD", DateTime.Today.ToString().Substring(0, 10));
            GridViewCert.SetFocusedRowCellValue("ALW_PMNT_YN", "N");
            //GridViewCert.ShowEditForm();

        }

        private bool SaveCertificationInfo(SqlCommand cmd)
        {
            bool result = false;
            if (GridViewCert.RowCount > 0)
            {
                DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridCert.DataSource);
                DataTable dtMerge = dsSave.Tables[0];

                string sEmpId = string.Empty;
                double dSeq = 0;
                string sCertGb = string.Empty;
                string sCertNo = string.Empty;
                string sCertNm = string.Empty;
                string sIssueYmd = string.Empty;
                string sEndYmd = string.Empty;
                string sIssueFacil = string.Empty;
                string sAlwPmntYn = string.Empty;
                double dAlwPmntRate = 0;
                double dAlwPmntAmt = 0;
                string sEntId = string.Empty;
                string sEntIp = string.Empty;
                string sMfyId = string.Empty;
                string sMfyIp = string.Empty;

                StringBuilder strSql = new StringBuilder();
                for (int i = 0; i < dtMerge.Rows.Count; i++)
                {
                    sEmpId = dtMerge.Rows[i]["EMP_ID"].ToString();

                    if (String.IsNullOrEmpty(dtMerge.Rows[i]["SEQ"].ToString()))
                    {
                        strSql.Clear();
                        strSql.AppendLine(" SELECT MAX(A.SEQ) AS MAX_VALUE");
                        strSql.AppendLine("   FROM HR_EMP_CERT A ");
                        strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");

                        DataTable dtCertChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                        if (dtCertChk.Rows.Count == 0 || String.IsNullOrEmpty(dtCertChk.Rows[0]["MAX_VALUE"].ToString()))
                        {
                            dSeq = 1;
                        }
                        else
                        {
                            dSeq = Convert.ToDouble(dtCertChk.Rows[0]["MAX_VALUE"].ToString()) + 1;
                        }
                    }
                    else
                    {
                        dSeq = Convert.ToDouble(dtMerge.Rows[i]["SEQ"].ToString());
                    }

                    sCertGb = dtMerge.Rows[i]["CERT_GB"].ToString();
                    sCertNo = dtMerge.Rows[i]["CERT_NO"].ToString();
                    sCertNm = dtMerge.Rows[i]["CERT_NM"].ToString();
                    sIssueYmd = dtMerge.Rows[i]["ISSUE_YMD"].ToString().Replace("-", "").Substring(0, 8);
                    sEndYmd = dtMerge.Rows[i]["END_YMD"].ToString().Replace("-", "").Substring(0, 8);
                    sIssueFacil = dtMerge.Rows[i]["ISSUE_FACIL"].ToString();
                    sAlwPmntYn = dtMerge.Rows[i]["ALW_PMNT_YN"].ToString();
                    dAlwPmntRate = String.IsNullOrEmpty(dtMerge.Rows[i]["ALW_PMNT_RATE"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["ALW_PMNT_RATE"].ToString());
                    dAlwPmntAmt = String.IsNullOrEmpty(dtMerge.Rows[i]["ALW_PMNT_AMT"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["ALW_PMNT_AMT"].ToString());
                    sEntId = FmMainToolBar2.drUser["USRCD"]?.ToString();
                    sEntIp = Get_MyIP();
                    sMfyId = FmMainToolBar2.drUser["USRCD"]?.ToString();
                    sMfyIp = Get_MyIP();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" INSERT INTO HR_EMP_CERT ");
                    //strSql.AppendLine(" 		  ( EMP_ID ");
                    //strSql.AppendLine(" 	      , SEQ ");
                    //strSql.AppendLine(" 	      , CERT_GB ");
                    //strSql.AppendLine(" 	      , CERT_NO ");
                    //strSql.AppendLine(" 	      , CERT_NM ");
                    //strSql.AppendLine(" 	      , ISSUE_YMD ");
                    //strSql.AppendLine(" 	      , END_YMD ");
                    //strSql.AppendLine(" 	      , ISSUE_FACIL ");
                    //strSql.AppendLine(" 	      , ALW_PMNT_YN ");
                    //strSql.AppendLine(" 	      , ALW_PMNT_RATE ");
                    //strSql.AppendLine(" 	      , ALW_PMNT_AMT ");
                    //strSql.AppendLine(" 	      , ENT_DT ");
                    //strSql.AppendLine(" 	      , ENT_ID ");
                    //strSql.AppendLine(" 	      , ENT_IP ");
                    //strSql.AppendLine(" 	      , MFY_DT ");
                    //strSql.AppendLine(" 	      , MFY_ID ");
                    //strSql.AppendLine(" 	      , MFY_IP ");
                    //strSql.AppendLine(" 		  ) ");
                    //strSql.AppendLine(" 	 VALUES  ");
                    //strSql.AppendLine("           ( '" + sEmpId + "' ");
                    //strSql.AppendLine("           ,  " + dSeq + " ");
                    //strSql.AppendLine("           , '" + sCertGb + "' ");
                    //strSql.AppendLine("           , '" + sCertNo + "' ");
                    //strSql.AppendLine("           , '" + sCertNm + "' ");
                    //strSql.AppendLine("           , '" + sIssueYmd + "' ");
                    //strSql.AppendLine("           , '" + sEndYmd + "' ");
                    //strSql.AppendLine("           , '" + sIssueFacil + "' ");
                    //strSql.AppendLine("           , '" + sAlwPmntYn + "' ");
                    //strSql.AppendLine("           ,  " + dAlwPmntRate + " ");
                    //strSql.AppendLine("           ,  " + dAlwPmntAmt + " ");
                    //strSql.AppendLine("           ,  NOW() ");
                    //strSql.AppendLine("           ,  '" + sEntId + "' ");
                    //strSql.AppendLine("           ,  '" + sEntIp + "' ");
                    //strSql.AppendLine("           ,  NOW()");
                    //strSql.AppendLine("           ,  '" + sMfyId + "' ");
                    //strSql.AppendLine("           ,  '" + sMfyIp + "' ");
                    //strSql.AppendLine("           )  ");
                    //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                    //strSql.AppendLine("             EMP_ID = '" + sEmpId + "' ");
                    //strSql.AppendLine("           , SEQ = " + dSeq + " ");
                    //strSql.AppendLine("           , CERT_GB = '" + sCertGb + "' ");
                    //strSql.AppendLine("           , CERT_NO = '" + sCertNo + "' ");
                    //strSql.AppendLine("           , CERT_NM = '" + sCertNm + "' ");
                    //strSql.AppendLine("           , ISSUE_YMD = '" + sIssueYmd + "' ");
                    //strSql.AppendLine("           , END_YMD = '" + sEndYmd + "' ");
                    //strSql.AppendLine("           , ISSUE_FACIL = '" + sIssueFacil + "' ");
                    //strSql.AppendLine("           , ALW_PMNT_YN = '" + sAlwPmntYn + "' ");
                    //strSql.AppendLine("           , ALW_PMNT_RATE = " + dAlwPmntRate + " ");
                    //strSql.AppendLine("           , ALW_PMNT_AMT = " + dAlwPmntAmt + " ");
                    //strSql.AppendLine("           , MFY_DT = NOW() ");
                    //strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");
                    //strSql.AppendLine("           , MFY_IP = '" + sMfyIp + "' ");
                    #endregion

                    strSql.AppendLine("IF EXISTS(SELECT* FROM HR_EMP_CERT WHERE EMP_ID = '"+ sEmpId + "' AND SEQ = "+ dSeq + ") ");
                    strSql.AppendLine("   BEGIN                                                          ");
                    strSql.AppendLine("         UPDATE HR_EMP_CERT                                       ");
                    strSql.AppendLine("            SET CERT_GB = '" + sCertGb + "'                       ");
                    strSql.AppendLine("              , CERT_NO = '" + sCertNo + "'                       ");
                    strSql.AppendLine("              , CERT_NM = '" + sCertNm + "'                       ");
                    strSql.AppendLine("              , ISSUE_YMD = '" + sIssueYmd + "'                   ");
                    strSql.AppendLine("              , END_YMD = '" + sEndYmd + "'                       ");
                    strSql.AppendLine("              , ISSUE_FACIL = '" + sIssueFacil + "'               ");
                    strSql.AppendLine("              , ALW_PMNT_YN = '" + sAlwPmntYn + "'                ");
                    strSql.AppendLine("              , ALW_PMNT_RATE = "+dAlwPmntRate +"                     ");
                    strSql.AppendLine("              , ALW_PMNT_AMT = "+dAlwPmntAmt+"                        ");
                    strSql.AppendLine("              , MFY_DT = convert(varchar(19),getdate(),20)                                ");
                    strSql.AppendLine("              , MFY_ID = '" + sMfyId + "'                         ");
                    strSql.AppendLine("              , MFY_IP = '" + sMfyIp + "'                         ");
                    strSql.AppendLine("          WHERE EMP_ID = '"+ sEmpId + "'                                       ");
                    strSql.AppendLine("            AND SEQ = "+ dSeq + "                                           ");
                    strSql.AppendLine("     END                                                          ");
                    strSql.AppendLine("ELSE                                                              ");
                    strSql.AppendLine("   BEGIN                                                          ");
                    strSql.AppendLine("         INSERT INTO HR_EMP_CERT                                  ");
                    strSql.AppendLine("              (EMP_ID                                             ");
                    strSql.AppendLine("              , SEQ                                               ");
                    strSql.AppendLine("              , CERT_GB                                           ");
                    strSql.AppendLine("              , CERT_NO                                           ");
                    strSql.AppendLine("              , CERT_NM                                           ");
                    strSql.AppendLine("              , ISSUE_YMD                                         ");
                    strSql.AppendLine("              , END_YMD                                           ");
                    strSql.AppendLine("              , ISSUE_FACIL                                       ");
                    strSql.AppendLine("              , ALW_PMNT_YN                                       ");
                    strSql.AppendLine("              , ALW_PMNT_RATE                                     ");
                    strSql.AppendLine("              , ALW_PMNT_AMT                                      ");
                    strSql.AppendLine("              , ENT_DT                                            ");
                    strSql.AppendLine("              , ENT_ID                                            ");
                    strSql.AppendLine("              , ENT_IP                                            ");
                    strSql.AppendLine("              )                                                   ");
                    strSql.AppendLine("                                                                  ");
                    strSql.AppendLine("         VALUES                                                   ");
                    strSql.AppendLine("              ('" + sEmpId + "'                                   ");
                    strSql.AppendLine("              , "+dSeq+"                                              ");
                    strSql.AppendLine("              , '" + sCertGb + "'                                 ");
                    strSql.AppendLine("              , '" + sCertNo + "'                                 ");
                    strSql.AppendLine("              , '" + sCertNm + "'                                 ");
                    strSql.AppendLine("              , '" + sIssueYmd + "'                               ");
                    strSql.AppendLine("              , '" + sEndYmd + "'                                 ");
                    strSql.AppendLine("              , '" + sIssueFacil + "'                             ");
                    strSql.AppendLine("              , '" + sAlwPmntYn + "'                              ");
                    strSql.AppendLine("              , " + dAlwPmntRate                                   );
                    strSql.AppendLine("              , " + dAlwPmntAmt                                    );
                    strSql.AppendLine("              , convert(varchar(19),getdate(),20)                                         ");
                    strSql.AppendLine("              , '" + sEntId + "'                                  ");
                    strSql.AppendLine("              , '" + sEntIp + "'                                  ");
                    strSql.AppendLine("              )                                                   ");
                    strSql.AppendLine("     END");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }
                result = true;
            }
            else
            {
                result = true;
            }
            return result;
        }

        private bool DeleteCertificationInfo(SqlCommand cmd)
        {
            string sEmpId = GridViewCert.GetFocusedRowCellValue("EMP_ID")?.ToString();
            string sSeq = GridViewCert.GetFocusedRowCellValue("SEQ")?.ToString();

            if (string.IsNullOrEmpty(sSeq))
            {
                int idx = GridViewRetr.FocusedRowHandle;
                BtnRetr.PerformClick();
                GridViewRetr.FocusedRowHandle = idx;
                return false;
            }

            string sCertNo = GridViewCert.GetFocusedRowCellValue("CERT_NO")?.ToString();
            string sCertNm = GridViewCert.GetFocusedRowCellValue("CERT_NM")?.ToString();

            if(GridViewCert.FocusedRowHandle < 0)
            {
                XtraMessageBox.Show("삭제할 항목이 없습니다.");
                return false;
            }

            if (XtraMessageBox.Show("자격명 : " + sCertNm + "\r\n자격번호 : " + sCertNo
                + "\r\n선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                , "자격/증명 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return false;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" DELETE ");
            strSql.AppendLine("   FROM HR_EMP_CERT ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");
            strSql.AppendLine("    AND SEQ = " + sSeq + " ");

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSql.ToString();
            cmd.ExecuteNonQuery();

            return true;
        }
        #endregion[자격증명 부분 Select, Add, Save, Delete ]

        #region [ 경력사항 부분 Select, Add, Save, Delete ]

        private void GetCareerInfo()
        {
            Cursor = Cursors.WaitCursor;

            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.EMP_ID ");
            strSql.AppendLine("      , A.SEQ ");
            strSql.AppendLine("      , A.STRT_YMD ");
            strSql.AppendLine("      , A.END_YMD ");
            strSql.AppendLine("      , A.DUTY_CO_NM ");
            strSql.AppendLine("      , A.DUTY_JOBDUTY ");
            strSql.AppendLine("      , A.DUTY_CONT ");
            strSql.AppendLine("      , A.CAREER_YCNT ");
            strSql.AppendLine("      , A.CAREER_MCNT ");
            strSql.AppendLine("      , A.CONV_YCNT ");
            strSql.AppendLine("      , A.CONV_MCNT ");
            strSql.AppendLine("      , A.CONV_RATE ");
            strSql.AppendLine("      , ENT_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.ENT_ID AS NUMERIC) IS NULL THEN A.ENT_ID ELSE DBO.FN_USRNM(A.ENT_ID) END AS ENT_ID  ");
            strSql.AppendLine("      , ENT_IP ");
            strSql.AppendLine("      , MFY_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MFY_ID AS NUMERIC) IS NULL THEN A.MFY_ID ELSE DBO.FN_USRNM(A.MFY_ID) END AS MFY_ID ");
            strSql.AppendLine("      , MFY_IP ");
            strSql.AppendLine("   FROM HR_EMP_CAREER A ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");
            strSql.AppendLine("  ORDER BY SEQ ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridCareer.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void AddCareerInfo()
        {
            Cursor = Cursors.WaitCursor;

            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            GridViewCareer.AddNewRow();
            GridViewCareer.SetFocusedRowCellValue("EMP_ID", sEmpId);
            GridViewCareer.SetFocusedRowCellValue("STRT_YMD", DateTime.Now.ToString().Substring(0, 10));
            GridViewCareer.SetFocusedRowCellValue("END_YMD", DateTime.Now.ToString().Substring(0, 10));
            //GridViewCareer.ShowEditForm();

            Cursor = Cursors.Default;
        }

        private bool SaveCareerInfo(SqlCommand cmd)
        {
            bool result = false;

            if (GridViewCareer.RowCount > 0)
            {
                DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridCareer.DataSource);
                DataTable dtMerge = dsSave.Tables[0];

                string sEmpId = string.Empty;
                double dSeq = 0;
                string sStrtYmd = string.Empty;
                string sEndYmd = string.Empty;
                string sDutyCoNm = string.Empty;
                string sDutyJobDuty = string.Empty;
                string sDutyCont = string.Empty;
                double dCareerYcnt = 0;
                double dCareerMcnt = 0;
                double dConvYcnt = 0;
                double dConvMcnt = 0;
                double dConvRate = 0;
                string sEntId = string.Empty;
                string sEntIp = string.Empty;
                string sMfyId = string.Empty;
                string sMfyIp = string.Empty;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < dtMerge.Rows.Count; i++)
                {
                    sEmpId = dtMerge.Rows[i]["EMP_ID"].ToString();

                    if (String.IsNullOrEmpty(dtMerge.Rows[i]["SEQ"].ToString()))
                    {
                        strSql.Clear();
                        strSql.AppendLine(" SELECT MAX(A.SEQ) AS MAX_VALUE");
                        strSql.AppendLine("   FROM HR_EMP_CAREER A ");
                        strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");

                        DataTable dtCareerChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                        if (dtCareerChk.Rows.Count == 0 || String.IsNullOrEmpty(dtCareerChk.Rows[0]["MAX_VALUE"].ToString()))
                        {
                            dSeq = 1;
                        }
                        else
                        {
                            dSeq = Convert.ToDouble(dtCareerChk.Rows[0]["MAX_VALUE"].ToString()) + 1;
                        }
                    }
                    else
                    {
                        dSeq = Convert.ToDouble(dtMerge.Rows[i]["SEQ"].ToString());
                    }

                    sStrtYmd = dtMerge.Rows[i]["STRT_YMD"].ToString().Replace("-", "").Substring(0, 8);
                    sEndYmd = dtMerge.Rows[i]["END_YMD"].ToString().Replace("-", "").Substring(0, 8);
                    sDutyCoNm = dtMerge.Rows[i]["DUTY_CO_NM"].ToString();
                    sDutyJobDuty = dtMerge.Rows[i]["DUTY_JOBDUTY"].ToString();
                    sDutyCont = dtMerge.Rows[i]["DUTY_CONT"].ToString();
                    dCareerYcnt = String.IsNullOrEmpty(dtMerge.Rows[i]["CAREER_YCNT"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["CAREER_YCNT"].ToString());
                    dCareerMcnt = String.IsNullOrEmpty(dtMerge.Rows[i]["CAREER_MCNT"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["CAREER_MCNT"].ToString());
                    dConvYcnt = String.IsNullOrEmpty(dtMerge.Rows[i]["CONV_YCNT"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["CONV_YCNT"].ToString());
                    dConvMcnt = String.IsNullOrEmpty(dtMerge.Rows[i]["CONV_MCNT"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["CONV_MCNT"].ToString());
                    dConvRate = String.IsNullOrEmpty(dtMerge.Rows[i]["CONV_RATE"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["CONV_RATE"].ToString());
                    sEntId = FmMainToolBar2.drUser["USRCD"]?.ToString();
                    sEntIp = Get_MyIP();
                    sMfyId = FmMainToolBar2.drUser["USRCD"]?.ToString();
                    sMfyIp = Get_MyIP();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" INSERT INTO HR_EMP_CAREER ");
                    //strSql.AppendLine(" 		  ( EMP_ID ");
                    //strSql.AppendLine(" 	      , SEQ ");
                    //strSql.AppendLine(" 	      , STRT_YMD ");
                    //strSql.AppendLine(" 	      , END_YMD ");
                    //strSql.AppendLine(" 	      , DUTY_CO_NM ");
                    //strSql.AppendLine(" 	      , DUTY_JOBDUTY ");
                    //strSql.AppendLine(" 	      , DUTY_CONT ");
                    //strSql.AppendLine(" 	      , CAREER_YCNT ");
                    //strSql.AppendLine(" 	      , CAREER_MCNT ");
                    //strSql.AppendLine(" 	      , CONV_YCNT ");
                    //strSql.AppendLine(" 	      , CONV_MCNT ");
                    //strSql.AppendLine(" 	      , CONV_RATE ");
                    //strSql.AppendLine(" 	      , ENT_DT ");
                    //strSql.AppendLine(" 	      , ENT_ID ");
                    //strSql.AppendLine(" 	      , ENT_IP ");
                    //strSql.AppendLine(" 	      , MFY_DT ");
                    //strSql.AppendLine(" 	      , MFY_ID ");
                    //strSql.AppendLine(" 	      , MFY_IP ");
                    //strSql.AppendLine(" 		  ) ");
                    //strSql.AppendLine(" 	 VALUES  ");
                    //strSql.AppendLine("           ( '" + sEmpId + "' ");
                    //strSql.AppendLine("           ,  " + dSeq + " ");
                    //strSql.AppendLine("           , '" + sStrtYmd + "' ");
                    //strSql.AppendLine("           , '" + sEndYmd + "' ");
                    //strSql.AppendLine("           , '" + sDutyCoNm + "' ");
                    //strSql.AppendLine("           , '" + sDutyJobDuty + "' ");
                    //strSql.AppendLine("           , '" + sDutyCont + "' ");
                    //strSql.AppendLine("           , " + dCareerYcnt + " ");
                    //strSql.AppendLine("           , " + dCareerMcnt + " ");
                    //strSql.AppendLine("           ,  " + dConvYcnt + " ");
                    //strSql.AppendLine("           ,  " + dConvMcnt + " ");
                    //strSql.AppendLine("           ,  " + dConvRate + " ");
                    //strSql.AppendLine("           ,  NOW() ");
                    //strSql.AppendLine("           ,  '" + sEntId + "' ");
                    //strSql.AppendLine("           ,  '" + sEntIp + "' ");
                    //strSql.AppendLine("           ,  NOW()");
                    //strSql.AppendLine("           ,  '" + sMfyId + "' ");
                    //strSql.AppendLine("           ,  '" + sMfyIp + "' ");
                    //strSql.AppendLine("           )  ");
                    //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                    //strSql.AppendLine("             STRT_YMD = '" + sStrtYmd + "' ");
                    //strSql.AppendLine("           , END_YMD = '" + sEndYmd + "' ");
                    //strSql.AppendLine("           , DUTY_CO_NM = '" + sDutyCoNm + "' ");
                    //strSql.AppendLine("           , DUTY_JOBDUTY = '" + sDutyJobDuty + "' ");
                    //strSql.AppendLine("           , DUTY_CONT = '" + sDutyCont + "' ");
                    //strSql.AppendLine("           , CAREER_YCNT = " + dCareerYcnt + " ");
                    //strSql.AppendLine("           , CAREER_MCNT = " + dCareerMcnt + " ");
                    //strSql.AppendLine("           , CONV_YCNT = " + dConvYcnt + " ");
                    //strSql.AppendLine("           , CONV_MCNT = " + dConvMcnt + " ");
                    //strSql.AppendLine("           , CONV_RATE = " + dConvRate + " ");
                    //strSql.AppendLine("           , MFY_DT = NOW() ");
                    //strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");
                    //strSql.AppendLine("           , MFY_IP = '" + sMfyIp + "' ");
                    #endregion

                    strSql.AppendLine("IF EXISTS(SELECT* FROM HR_EMP_CAREER WHERE EMP_ID = '"+ sEmpId + "' AND SEQ = "+ dSeq + ")");
                    strSql.AppendLine("   BEGIN                                                           ");
                    strSql.AppendLine("         UPDATE HR_EMP_CAREER                                      ");
                    strSql.AppendLine("            SET STRT_YMD = '" + sStrtYmd + "' ");
	                strSql.AppendLine("              , END_YMD = '" + sEndYmd + "' ");
	                strSql.AppendLine("              , DUTY_CO_NM = '" + sDutyCoNm + "' ");
	                strSql.AppendLine("              , DUTY_JOBDUTY = '" + sDutyJobDuty + "' ");
	                strSql.AppendLine("              , DUTY_CONT = '" + sDutyCont + "' ");
	                strSql.AppendLine("              , CAREER_YCNT = " + dCareerYcnt + " ");
	                strSql.AppendLine("              , CAREER_MCNT = " + dCareerMcnt + " ");
	                strSql.AppendLine("              , CONV_YCNT = " + dConvYcnt + " ");
	                strSql.AppendLine("              , CONV_MCNT = " + dConvMcnt + " ");
	                strSql.AppendLine("              , CONV_RATE = " + dConvRate + " ");
	                strSql.AppendLine("              , MFY_DT = convert(varchar(19),getdate(),20) ");
	                strSql.AppendLine("              , MFY_ID = '" + sMfyId + "' ");
	                strSql.AppendLine("              , MFY_IP = '" + sMfyIp + "' ");
                    strSql.AppendLine("          WHERE EMP_ID = '"+ sEmpId + "'");
                    strSql.AppendLine("            AND SEQ = "+ dSeq + "");
                    strSql.AppendLine("     END               ");
                    strSql.AppendLine("ELSE                   ");
                    strSql.AppendLine("   BEGIN               ");
                    strSql.AppendLine("         INSERT INTO HR_EMP_CAREER ");
                    strSql.AppendLine("              (EMP_ID ");
                    strSql.AppendLine("              , SEQ ");
                    strSql.AppendLine("              , STRT_YMD ");
                    strSql.AppendLine("              , END_YMD ");
                    strSql.AppendLine("              , DUTY_CO_NM ");
                    strSql.AppendLine("              , DUTY_JOBDUTY ");
                    strSql.AppendLine("              , DUTY_CONT ");
                    strSql.AppendLine("              , CAREER_YCNT ");
                    strSql.AppendLine("              , CAREER_MCNT ");
                    strSql.AppendLine("              , CONV_YCNT ");
                    strSql.AppendLine("              , CONV_MCNT ");
                    strSql.AppendLine("              , CONV_RATE ");
                    strSql.AppendLine("              , ENT_DT ");
                    strSql.AppendLine("              , ENT_ID ");
                    strSql.AppendLine("              , ENT_IP ");
                    strSql.AppendLine("              ) ");
                    strSql.AppendLine("         VALUES  ");
                    strSql.AppendLine("              ('" + sEmpId + "' ");
                    strSql.AppendLine("              , " + dSeq + " ");
                    strSql.AppendLine("              , '" + sStrtYmd + "' ");
                    strSql.AppendLine("              , '" + sEndYmd + "' ");
                    strSql.AppendLine("              , '" + sDutyCoNm + "' ");
                    strSql.AppendLine("              , '" + sDutyJobDuty + "' ");
                    strSql.AppendLine("              , '" + sDutyCont + "' ");
                    strSql.AppendLine("              , " + dCareerYcnt + " ");
                    strSql.AppendLine("              , " + dCareerMcnt + " ");
                    strSql.AppendLine("              , " + dConvYcnt + " ");
                    strSql.AppendLine("              , " + dConvMcnt + " ");
                    strSql.AppendLine("              , " + dConvRate + " ");
                    strSql.AppendLine("              , convert(varchar(19),getdate(),20) ");
                    strSql.AppendLine("              , '" + sEntId + "' ");
                    strSql.AppendLine("              , '" + sEntIp + "' ");
                    strSql.AppendLine("              )  ");
                    strSql.AppendLine("     END");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                }

                result = true;
            }
            else
            {
                result = true;
            }
            return result;
        }

        private bool DeleteCareerInfo(SqlCommand cmd)
        {
            StringBuilder strSql = new StringBuilder();

            string sEmpId = GridViewCareer.GetFocusedRowCellValue("EMP_ID")?.ToString();
            string sSeq = GridViewCareer.GetFocusedRowCellValue("SEQ")?.ToString();

            if (string.IsNullOrEmpty(sSeq))
            {
                int idx = GridViewRetr.FocusedRowHandle;
                BtnRetr.PerformClick();
                GridViewRetr.FocusedRowHandle = idx;
                return false;
            }

            string sStrtYmd = GridViewCareer.GetFocusedRowCellValue("STRT_YMD")?.ToString();
            string sEndYmd = GridViewCareer.GetFocusedRowCellValue("END_YMD")?.ToString();
            string sCareerNm = GridViewCareer.GetFocusedRowCellValue("DUTY_CO_NM")?.ToString();

            if (GridViewCareer.FocusedRowHandle < 0)
            {
                XtraMessageBox.Show("삭제할 항목이 없습니다.");
                return false;
            }

            if (XtraMessageBox.Show("시작일자 : " + sStrtYmd + "\r\n종료일자 : " + sEndYmd + "\r\n직장명 : " + sCareerNm
                + " \r\n선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                , "경력 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return false;
            }

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" DELETE ");
            strSql.AppendLine("   FROM HR_EMP_CAREER ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");
            strSql.AppendLine("    AND SEQ = " + sSeq + " ");

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSql.ToString();
            cmd.ExecuteNonQuery();

            return true;
        }

        # endregion [ 경력사항 부분 Select, Add, Save, Delete ]

        #region [ 가족사항 부분 Select, Add, Save, Delete ]

        private void GetFamilyInfo()
        {
            Cursor = Cursors.WaitCursor;

            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.EMP_ID ");
            strSql.AppendLine("      , A.SEQ ");
            strSql.AppendLine("      , A.IDT_NO ");
            strSql.AppendLine("      , A.FAML_RELATION_CD ");
            strSql.AppendLine("      , A.FAML_NM ");
            strSql.AppendLine("      , A.COHABIT_YN ");
            strSql.AppendLine("      , A.DDCT_YN ");
            strSql.AppendLine("      , A.ACADEMIC_GB ");
            strSql.AppendLine("      , A.HEALTH_INSR_YN ");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.ENT_ID AS NUMERIC) IS NULL THEN A.ENT_ID ELSE DBO.FN_USRNM(A.ENT_ID) END AS ENT_ID ");
            strSql.AppendLine("      , A.ENT_IP ");
            strSql.AppendLine("      , A.MFY_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MFY_ID AS NUMERIC) IS NULL THEN A.MFY_ID ELSE DBO.FN_USRNM(A.MFY_ID) END AS MFY_ID ");
            strSql.AppendLine("      , A.MFY_IP ");
            strSql.AppendLine("   FROM HR_EMP_FAMILY A ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");
            strSql.AppendLine("  ORDER BY SEQ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if(dt != null)
            {
                foreach(DataRow row in dt.Rows)
                {
                    string sIdtNo = row["IDT_NO"]?.ToString();

                    if (!string.IsNullOrEmpty(sIdtNo))
                    {
                        row["IDT_NO"] = ComnEtcFunc.Decrypt(sIdtNo, ComnEtcFunc._SECRET_KEY2);
                    }
                }
            }

            GridFamily.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void AddFamilyInfo()
        {
            Cursor = Cursors.WaitCursor;

            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            GridViewFamily.AddNewRow();
            GridViewFamily.SetFocusedRowCellValue("EMP_ID", sEmpId);
            //GridViewFamily.ShowEditForm();

            Cursor = Cursors.Default;
        }

        private bool SaveFamilyInfo(SqlCommand cmd)
        {
            bool result = false;

            DataTable dtChk = (DataTable)GridFamily.DataSource;
            for(int i = 0; i < dtChk.Rows.Count; i++)
            {
                if (String.IsNullOrEmpty(dtChk.Rows[i]["IDT_NO"].ToString()))
                {
                    XtraMessageBox.Show("가족관계에서 식별번호를 입력하세요.");
                    GridViewFamily.FocusedRowHandle = i;
                    return false;
                }
            }

            if (GridViewFamily.RowCount > 0)
            {
                DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridFamily.DataSource);
                DataTable dtMerge = dsSave.Tables[0];

                string sEmpId = string.Empty;
                double dSeq = 0;
                string sIdtNo = string.Empty;
                string sFamlRelationCd = string.Empty;
                string sFamlNm = string.Empty;
                string sCohabitYn = string.Empty;
                string sDdctYn = string.Empty;
                string sAcademicGb = string.Empty;
                string sHealthInsrYn = string.Empty;
                string sEntId = string.Empty;
                string sEntIp = string.Empty;
                string sMfyId = string.Empty;
                string sMfyIp = string.Empty;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < dtMerge.Rows.Count; i++)
                {
                    sEmpId = dtMerge.Rows[i]["EMP_ID"].ToString();

                    if (String.IsNullOrEmpty(dtMerge.Rows[i]["SEQ"].ToString()))
                    {
                        strSql.Clear();
                        strSql.AppendLine(" SELECT MAX(A.SEQ) AS MAX_VALUE");
                        strSql.AppendLine("   FROM HR_EMP_FAMILY A ");
                        strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");

                        DataTable dtCareerChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                        if (dtCareerChk.Rows.Count == 0 || String.IsNullOrEmpty(dtCareerChk.Rows[0]["MAX_VALUE"].ToString()))
                        {
                            dSeq = 1;
                        }
                        else
                        {
                            dSeq = Convert.ToDouble(dtCareerChk.Rows[0]["MAX_VALUE"].ToString()) + 1;
                        }
                    }
                    else
                    {
                        dSeq = Convert.ToDouble(dtMerge.Rows[i]["SEQ"].ToString());
                    }

                    sIdtNo = dtMerge.Rows[i]["IDT_NO"].ToString();
                    if (!string.IsNullOrEmpty(sIdtNo))
                    {
                        sIdtNo = ComnEtcFunc.Encrypt(sIdtNo, ComnEtcFunc._SECRET_KEY2);
                    }
                    sFamlRelationCd = dtMerge.Rows[i]["FAML_RELATION_CD"].ToString();
                    sFamlNm = dtMerge.Rows[i]["FAML_NM"].ToString();
                    sCohabitYn = dtMerge.Rows[i]["COHABIT_YN"].ToString();
                    sDdctYn = dtMerge.Rows[i]["DDCT_YN"].ToString();
                    sAcademicGb = dtMerge.Rows[i]["ACADEMIC_GB"].ToString();
                    sHealthInsrYn = dtMerge.Rows[i]["HEALTH_INSR_YN"].ToString();
                    sEntId = FmMainToolBar2.drUser["USRCD"]?.ToString();
                    sEntIp = Get_MyIP();
                    sMfyId = FmMainToolBar2.drUser["USRCD"]?.ToString();
                    sMfyIp = Get_MyIP();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" INSERT INTO HR_EMP_FAMILY ");
                    //strSql.AppendLine(" 		  ( EMP_ID ");
                    //strSql.AppendLine(" 	      , SEQ ");
                    //strSql.AppendLine(" 	      , IDT_NO ");
                    //strSql.AppendLine(" 	      , FAML_RELATION_CD ");
                    //strSql.AppendLine(" 	      , FAML_NM ");
                    //strSql.AppendLine(" 	      , COHABIT_YN ");
                    //strSql.AppendLine(" 	      , DDCT_YN ");
                    //strSql.AppendLine(" 	      , ACADEMIC_GB ");
                    //strSql.AppendLine(" 	      , HEALTH_INSR_YN ");
                    //strSql.AppendLine(" 	      , ENT_DT ");
                    //strSql.AppendLine(" 	      , ENT_ID ");
                    //strSql.AppendLine(" 	      , ENT_IP ");
                    //strSql.AppendLine(" 	      , MFY_DT ");
                    //strSql.AppendLine(" 	      , MFY_ID ");
                    //strSql.AppendLine(" 	      , MFY_IP ");
                    //strSql.AppendLine(" 		  ) ");
                    //strSql.AppendLine(" 	 VALUES  ");
                    //strSql.AppendLine("           ( '" + sEmpId + "' ");
                    //strSql.AppendLine("           , " + dSeq + " ");
                    //strSql.AppendLine("           , '" + sIdtNo + "' ");
                    //strSql.AppendLine("           , '" + sFamlRelationCd + "' ");
                    //strSql.AppendLine("           , '" + sFamlNm + "' ");
                    //strSql.AppendLine("           , '" + sCohabitYn + "' ");
                    //strSql.AppendLine("           , '" + sDdctYn + "' ");
                    //strSql.AppendLine("           , '" + sAcademicGb + "' ");
                    //strSql.AppendLine("           , '" + sHealthInsrYn + "' ");
                    //strSql.AppendLine("           ,  NOW() ");
                    //strSql.AppendLine("           ,  '" + sEntId + "' ");
                    //strSql.AppendLine("           ,  '" + sEntIp + "' ");
                    //strSql.AppendLine("           ,  NOW()");
                    //strSql.AppendLine("           ,  '" + sMfyId + "' ");
                    //strSql.AppendLine("           ,  '" + sMfyIp + "' ");
                    //strSql.AppendLine("           )  ");
                    //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                    //strSql.AppendLine("             IDT_NO = '" + sIdtNo + "' ");
                    //strSql.AppendLine("           , FAML_RELATION_CD = '" + sFamlRelationCd + "' ");
                    //strSql.AppendLine("           , FAML_NM = '" + sFamlNm + "' ");
                    //strSql.AppendLine("           , COHABIT_YN = '" + sCohabitYn + "' ");
                    //strSql.AppendLine("           , DDCT_YN = '" + sDdctYn + "' ");
                    //strSql.AppendLine("           , ACADEMIC_GB = '" + sAcademicGb + "' ");
                    //strSql.AppendLine("           , HEALTH_INSR_YN = '" + sHealthInsrYn + "' ");
                    //strSql.AppendLine("           , MFY_DT = NOW() ");
                    //strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");
                    //strSql.AppendLine("           , MFY_IP = '" + sMfyIp + "' ");
                    #endregion

                    strSql.AppendLine("IF EXISTS(SELECT* FROM HR_EMP_FAMILY WHERE EMP_ID = '"+ sEmpId + "' AND SEQ = "+ dSeq + ")");
                    strSql.AppendLine("   BEGIN                                                           ");
                    strSql.AppendLine("         UPDATE HR_EMP_FAMILY");
                    strSql.AppendLine("            SET IDT_NO = '" + sIdtNo + "' ");
	                strSql.AppendLine("              , FAML_RELATION_CD = '" + sFamlRelationCd + "' ");
	                strSql.AppendLine("              , FAML_NM = '" + sFamlNm + "' ");
	                strSql.AppendLine("              , COHABIT_YN = '" + sCohabitYn + "' ");
	                strSql.AppendLine("              , DDCT_YN = '" + sDdctYn + "' ");
	                strSql.AppendLine("              , ACADEMIC_GB = '" + sAcademicGb + "' ");
	                strSql.AppendLine("              , HEALTH_INSR_YN = '" + sHealthInsrYn + "' ");
	                strSql.AppendLine("              , MFY_DT = convert(varchar(19),getdate(),20) ");
	                strSql.AppendLine("              , MFY_ID = '" + sMfyId + "' ");
	                strSql.AppendLine("              , MFY_IP = '" + sMfyIp + "' ");
                    strSql.AppendLine("          WHERE EMP_ID = '"+ sEmpId + "'");
                    strSql.AppendLine("            AND SEQ = "+ dSeq + "    ");
                    strSql.AppendLine("     END                   ");
                    strSql.AppendLine("ELSE                       ");
                    strSql.AppendLine("   BEGIN                   ");
                    strSql.AppendLine("         INSERT INTO HR_EMP_FAMILY ");
                    strSql.AppendLine("                (EMP_ID ");
                    strSql.AppendLine("              , SEQ ");
                    strSql.AppendLine("              , IDT_NO ");
                    strSql.AppendLine("              , FAML_RELATION_CD ");
                    strSql.AppendLine("              , FAML_NM ");
                    strSql.AppendLine("              , COHABIT_YN ");
                    strSql.AppendLine("              , DDCT_YN ");
                    strSql.AppendLine("              , ACADEMIC_GB ");
                    strSql.AppendLine("              , HEALTH_INSR_YN ");
                    strSql.AppendLine("              , ENT_DT ");
                    strSql.AppendLine("              , ENT_ID ");
                    strSql.AppendLine("              , ENT_IP ");
                    strSql.AppendLine("              ) ");
                    strSql.AppendLine("         VALUES  ");
                    strSql.AppendLine("              ('" + sEmpId + "' ");
                    strSql.AppendLine("              , " + dSeq + " ");
                    strSql.AppendLine("              , '" + sIdtNo + "' ");
                    strSql.AppendLine("              , '" + sFamlRelationCd + "' ");
                    strSql.AppendLine("              , '" + sFamlNm + "' ");
                    strSql.AppendLine("              , '" + sCohabitYn + "' ");
                    strSql.AppendLine("              , '" + sDdctYn + "' ");
                    strSql.AppendLine("              , '" + sAcademicGb + "' ");
                    strSql.AppendLine("              , '" + sHealthInsrYn + "' ");
                    strSql.AppendLine("              , convert(varchar(19),getdate(),20) ");
                    strSql.AppendLine("              , '" + sEntId + "' ");
                    strSql.AppendLine("              , '" + sEntIp + "' ");
                    strSql.AppendLine("              )  ");
                    strSql.AppendLine("     END");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }
                result = true;
            }

            return result;
        }

        private bool DeleteFamilyInfo(SqlCommand cmd)
        {
            string sEmpId = GridViewFamily.GetFocusedRowCellValue("EMP_ID")?.ToString();
            string sSeq = GridViewFamily.GetFocusedRowCellValue("SEQ")?.ToString();

            if (string.IsNullOrEmpty(sSeq))
            {
                int idx = GridViewRetr.FocusedRowHandle;
                BtnRetr.PerformClick();
                GridViewRetr.FocusedRowHandle = idx;

                return false;
            }

            string sIdtNo = GridViewFamily.GetFocusedRowCellValue("IDT_NO")?.ToString();
            string sFamlNm = GridViewFamily.GetFocusedRowCellValue("FAML_NM")?.ToString();

            if(GridViewFamily.FocusedRowHandle < 0)
            {
                XtraMessageBox.Show("삭제할 항목이 없습니다.");
                return false;
            }

            if (XtraMessageBox.Show("이름 : " + sFamlNm + "\r\n식별번호 : " + sIdtNo 
                + " 선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                , "가족 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return false;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" DELETE ");
            strSql.AppendLine("   FROM HR_EMP_FAMILY ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");
            strSql.AppendLine("    AND SEQ = " + sSeq + " ");

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSql.ToString();
            cmd.ExecuteNonQuery();

            return true;
        }

        #endregion [ 가족사항 부분 Select, Add, Save, Delete ]

        #region [ 학력 부분 Select, Add, Save, Delete ]

        private void GetEducationInfo()
        {
            Cursor = Cursors.WaitCursor;

            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.EMP_ID ");
            strSql.AppendLine("      , A.STRT_YMD ");
            strSql.AppendLine("      , A.END_YMD ");
            strSql.AppendLine("      , A.ACADEMIC_GB ");
            strSql.AppendLine("      , A.EDU_FACIL_NM ");
            strSql.AppendLine("      , A.GRDU_GB ");
            strSql.AppendLine("      , A.MAJOR ");
            strSql.AppendLine("      , A.SUB_MAJOR ");
            strSql.AppendLine("      , A.NOTE ");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.ENT_ID AS NUMERIC) IS NULL THEN A.ENT_ID ELSE DBO.FN_USRNM(A.ENT_ID) END AS ENT_ID ");
            strSql.AppendLine("      , A.ENT_IP ");
            strSql.AppendLine("      , A.MFY_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MFY_ID AS NUMERIC) IS NULL THEN A.MFY_ID ELSE DBO.FN_USRNM(A.MFY_ID) END AS MFY_ID");
            strSql.AppendLine("      , A.MFY_IP ");
            strSql.AppendLine("   FROM HR_EMP_ACADEMIC A ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");
            strSql.AppendLine("  ORDER BY STRT_YMD, END_YMD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridEdu.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void AddEducationInfo()
        {
            Cursor = Cursors.WaitCursor;

            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            GridViewEdu.AddNewRow();
            GridViewEdu.SetFocusedRowCellValue("EMP_ID", sEmpId);
            GridViewEdu.SetFocusedRowCellValue("STRT_YMD", DateTime.Now.ToString().Substring(0, 10));
            GridViewEdu.SetFocusedRowCellValue("END_YMD", DateTime.Now.ToString().Substring(0, 10));
            //GridViewEdu.ShowEditForm();

            Cursor = Cursors.Default;
        }

        private bool SaveEducationInfo(SqlCommand cmd)
        {
            bool result = false;

            DataTable dtChk = (DataTable)GridEdu.DataSource;
            
            if (GridViewEdu.RowCount > 0)
            {
                DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridEdu.DataSource);
                DataTable dtMerge = dsSave.Tables[0];

                string sEmpId = string.Empty;
                string sStrtYmd = string.Empty;
                string sEndYmd = string.Empty;
                string sAcademicGb = string.Empty;
                string sEduFacilNm = string.Empty;
                string sGrduGb = string.Empty;
                string sMajor = string.Empty;
                string sSubMajor = string.Empty;
                string sNote = string.Empty;
                string sEntId = string.Empty;
                string sEntIp = string.Empty;
                string sMfyId = string.Empty;
                string sMfyIp = string.Empty;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < dtMerge.Rows.Count; i++)
                {
                    sEmpId = dtMerge.Rows[i]["EMP_ID"].ToString();
                    sStrtYmd = dtMerge.Rows[i]["STRT_YMD"].ToString().Replace("-", "");
                    if (!string.IsNullOrEmpty(sStrtYmd))
                        sStrtYmd = sStrtYmd.Substring(0, 8);
                    sEndYmd = dtMerge.Rows[i]["END_YMD"].ToString().Replace("-", "");
                    if (!string.IsNullOrEmpty(sEndYmd))
                        sEndYmd = sEndYmd.Substring(0, 8);
                    sAcademicGb = dtMerge.Rows[i]["ACADEMIC_GB"].ToString();
                    sEduFacilNm = dtMerge.Rows[i]["EDU_FACIL_NM"].ToString();
                    sGrduGb = dtMerge.Rows[i]["GRDU_GB"].ToString();
                    sMajor = dtMerge.Rows[i]["MAJOR"].ToString();
                    sSubMajor = dtMerge.Rows[i]["SUB_MAJOR"].ToString();
                    sNote = dtMerge.Rows[i]["NOTE"].ToString();
                    sEntId = FmMainToolBar2.drUser["USRCD"]?.ToString();
                    sEntIp = Get_MyIP();
                    sMfyId = FmMainToolBar2.drUser["USRCD"]?.ToString();
                    sMfyIp = Get_MyIP();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    #region mariaDB
                    //strSql.AppendLine(" INSERT INTO HR_EMP_ACADEMIC ");
                    //strSql.AppendLine(" 		  ( EMP_ID ");
                    //strSql.AppendLine(" 	      , STRT_YMD ");
                    //strSql.AppendLine(" 	      , END_YMD ");
                    //strSql.AppendLine(" 	      , ACADEMIC_GB ");
                    //strSql.AppendLine(" 	      , EDU_FACIL_NM ");
                    //strSql.AppendLine(" 	      , GRDU_GB ");
                    //strSql.AppendLine(" 	      , MAJOR ");
                    //strSql.AppendLine(" 	      , SUB_MAJOR ");
                    //strSql.AppendLine(" 	      , NOTE ");
                    //strSql.AppendLine(" 	      , ENT_DT ");
                    //strSql.AppendLine(" 	      , ENT_ID ");
                    //strSql.AppendLine(" 	      , ENT_IP ");
                    //strSql.AppendLine(" 	      , MFY_DT ");
                    //strSql.AppendLine(" 	      , MFY_ID ");
                    //strSql.AppendLine(" 	      , MFY_IP ");
                    //strSql.AppendLine(" 		  ) ");
                    //strSql.AppendLine(" 	 VALUES  ");
                    //strSql.AppendLine("           ( '" + sEmpId + "' ");
                    //strSql.AppendLine("           , '" + sStrtYmd + "' ");
                    //strSql.AppendLine("           , '" + sEndYmd + "' ");
                    //strSql.AppendLine("           , '" + sAcademicGb + "' ");
                    //strSql.AppendLine("           , '" + sEduFacilNm + "' ");
                    //strSql.AppendLine("           , '" + sGrduGb + "' ");
                    //strSql.AppendLine("           , '" + sMajor + "' ");
                    //strSql.AppendLine("           , '" + sSubMajor + "' ");
                    //strSql.AppendLine("           , '" + sNote + "' ");
                    //strSql.AppendLine("           ,  NOW() ");
                    //strSql.AppendLine("           ,  '" + sEntId + "' ");
                    //strSql.AppendLine("           ,  '" + sEntIp + "' ");
                    //strSql.AppendLine("           ,  NOW()");
                    //strSql.AppendLine("           ,  '" + sMfyId + "' ");
                    //strSql.AppendLine("           ,  '" + sMfyIp + "' ");
                    //strSql.AppendLine("           )  ");
                    //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                    //strSql.AppendLine("             STRT_YMD = '" + sStrtYmd + "' ");
                    //strSql.AppendLine("           , END_YMD = '" + sEndYmd + "' ");
                    //strSql.AppendLine("           , ACADEMIC_GB = '" + sAcademicGb + "' ");
                    //strSql.AppendLine("           , EDU_FACIL_NM = '" + sEduFacilNm + "' ");
                    //strSql.AppendLine("           , GRDU_GB = '" + sGrduGb + "' ");
                    //strSql.AppendLine("           , MAJOR = '" + sMajor + "' ");
                    //strSql.AppendLine("           , SUB_MAJOR = '" + sSubMajor + "' ");
                    //strSql.AppendLine("           , NOTE = '" + sNote + "' ");
                    //strSql.AppendLine("           , MFY_DT = NOW() ");
                    //strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");
                    //strSql.AppendLine("           , MFY_IP = '" + sMfyIp + "' ");
                    #endregion
                    strSql.AppendLine("IF EXISTS(SELECT* FROM HR_EMP_ACADEMIC WHERE EMP_ID = '"+ sEmpId + "' AND STRT_YMD = '" + sStrtYmd + "' AND END_YMD = '" + sEndYmd + "') ");
                    strSql.AppendLine("   BEGIN                        ");
                    strSql.AppendLine("         UPDATE HR_EMP_ACADEMIC ");
                    strSql.AppendLine("            SET ACADEMIC_GB = '" + sAcademicGb + "' ");
	                strSql.AppendLine("              , EDU_FACIL_NM = '" + sEduFacilNm + "' ");
	                strSql.AppendLine("              , GRDU_GB = '" + sGrduGb + "' ");
	                strSql.AppendLine("              , MAJOR = '" + sMajor + "' ");
	                strSql.AppendLine("              , SUB_MAJOR = '" + sSubMajor + "' ");
	                strSql.AppendLine("              , NOTE = '" + sNote + "' ");
	                strSql.AppendLine("              , MFY_DT = convert(varchar(19),getdate(),20) ");
	                strSql.AppendLine("              , MFY_ID = '" + sMfyId + "' ");
	                strSql.AppendLine("              , MFY_IP = '" + sMfyIp + "' ");
                    strSql.AppendLine("          WHERE EMP_ID = '"+ sEmpId + "'");
                    strSql.AppendLine("            AND STRT_YMD = '" + sStrtYmd + "' ");
                    strSql.AppendLine("            AND END_YMD = '" + sEndYmd + "' ");
                    strSql.AppendLine("     END     ");
                    strSql.AppendLine("ELSE         ");
                    strSql.AppendLine("   BEGIN     ");
                    strSql.AppendLine("         INSERT INTO HR_EMP_ACADEMIC ");
                    strSql.AppendLine("          (EMP_ID ");
                    strSql.AppendLine("          , STRT_YMD ");
                    strSql.AppendLine("          , END_YMD ");
                    strSql.AppendLine("          , ACADEMIC_GB ");
                    strSql.AppendLine("          , EDU_FACIL_NM ");
                    strSql.AppendLine("          , GRDU_GB ");
                    strSql.AppendLine("          , MAJOR ");
                    strSql.AppendLine("          , SUB_MAJOR ");
                    strSql.AppendLine("          , NOTE ");
                    strSql.AppendLine("          , ENT_DT ");
                    strSql.AppendLine("          , ENT_ID ");
                    strSql.AppendLine("          , ENT_IP ");
                    strSql.AppendLine("          ) ");
                    strSql.AppendLine("     VALUES  ");
                    strSql.AppendLine("          ('" + sEmpId + "' ");
                    strSql.AppendLine("          , '" + sStrtYmd + "' ");
                    strSql.AppendLine("          , '" + sEndYmd + "' ");
                    strSql.AppendLine("          , '" + sAcademicGb + "' ");
                    strSql.AppendLine("          , '" + sEduFacilNm + "' ");
                    strSql.AppendLine("          , '" + sGrduGb + "' ");
                    strSql.AppendLine("          , '" + sMajor + "' ");
                    strSql.AppendLine("          , '" + sSubMajor + "' ");
                    strSql.AppendLine("          , '" + sNote + "' ");
                    strSql.AppendLine("          , convert(varchar(19),getdate(),20) ");
                    strSql.AppendLine("          , '" + sEntId + "' ");
                    strSql.AppendLine("          , '" + sEntIp + "' ");
                    strSql.AppendLine("          )  ");
                    strSql.AppendLine("     END");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }
                result = true;
            }
            else
            {
                result = true;
            }

            return result;
        }

        private bool DeleteEducationInfo(SqlCommand cmd)
        {
            string sEmpId = GridViewEdu.GetFocusedRowCellValue("EMP_ID")?.ToString();
            string sStrtYmd = GridViewEdu.GetFocusedRowCellValue("STRT_YMD")?.ToString();
            string sEndYmd = GridViewEdu.GetFocusedRowCellValue("END_YMD")?.ToString();

            if (string.IsNullOrEmpty(sStrtYmd) || string.IsNullOrEmpty(sEndYmd))
            {
                int idx = GridViewRetr.FocusedRowHandle;
                BtnRetr.PerformClick();
                GridViewRetr.FocusedRowHandle = idx;
                return false;
            }

            if(GridViewEdu.FocusedRowHandle < 0)
            {
                XtraMessageBox.Show("삭제할 항목이 없습니다.");
                return false;
            }

            string sEduNm = GridViewEdu.GetFocusedRowCellValue("EDU_FACIL_NM")?.ToString();

            if (XtraMessageBox.Show("시작일자 : " + sStrtYmd + "\r\n종료일자 : " + sEndYmd + "\r\n교육기관 : " + sEduNm
                + "\r\n선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                , "학력 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return false;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" DELETE ");
            strSql.AppendLine("   FROM HR_EMP_ACADEMIC ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");
            strSql.AppendLine("    AND STRT_YMD = '" + sStrtYmd + "' ");
            strSql.AppendLine("    AND END_YMD = '" + sEndYmd + "' ");

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSql.ToString();
            cmd.ExecuteNonQuery();

            return true;
        }

        #endregion[ 학력 부분 Select, Add, Save, Delete ]

        #region [ 급여 부분 Select, Add, Save ]

        private void GetPaymentInfo()
        {
            Cursor = Cursors.WaitCursor;

            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.EMP_ID ");
            strSql.AppendLine("      , A.SLR_BSS_AMT  ");
            strSql.AppendLine("      , A.YAKAMT");
            strSql.AppendLine("      , A.SLR_ANL_AMT  ");
            strSql.AppendLine("      , A.SLR_DLY_AMT  ");
            strSql.AppendLine("      , A.SLR_TIM_AMT  ");
            strSql.AppendLine("      , A.BASTIME1     ");
            strSql.AppendLine("      , A.BASTIME2     ");
            strSql.AppendLine("      , A.SUDNG1       ");
            strSql.AppendLine("      , A.SUDNG2       ");
            strSql.AppendLine("      , A.SUDNG3       ");
            strSql.AppendLine("      , A.SUDNG4       ");
            strSql.AppendLine("      , A.SUDNG5       ");
            strSql.AppendLine("      , A.SUDNG6       ");
            strSql.AppendLine("      , A.SUDNG7       ");
            strSql.AppendLine("      , A.GONGJ1       ");
            strSql.AppendLine("      , A.GONGJ2       ");
            strSql.AppendLine("      , A.GONGJ3       ");
            strSql.AppendLine("      , A.GONGJ4       ");
            strSql.AppendLine("      , A.GONGJ5       ");
            strSql.AppendLine("      , A.GONGJ6       ");
            strSql.AppendLine("      , A.GONGJ7       ");
            strSql.AppendLine("      , A.DEAL_BANK_CD ");
            strSql.AppendLine("      , A.DEAL_BANK_NM ");
            strSql.AppendLine("      , A.ACNT_HDR     ");
            strSql.AppendLine("      , A.PMNT_ACNT_FST");
            strSql.AppendLine("      , A.PMNT_ACNT_SCD");
            strSql.AppendLine("      , A.PMNT_ACNT_TRD");
            strSql.AppendLine("      , A.NOTE         ");
            strSql.AppendLine("   FROM HR_EMP_SALARY A ");
            strSql.AppendLine("  WHERE A.EMP_ID = '" + sEmpId + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt == null || dt.Rows.Count == 0)
            {
                TxtYAKAMT.EditValue = string.Empty;
                TxtBasicSalary.EditValue = string.Empty;
                TxtAnnualSalary.EditValue = string.Empty;
                TxtDailySalary.EditValue = string.Empty;
                TxtSlrTimAmt.EditValue = string.Empty;
                TxtBTime1.EditValue = string.Empty;
                TxtBTime2.EditValue = string.Empty;

                TxtSubng1.EditValue = string.Empty;

                TxtGongj1.EditValue = string.Empty;
                TxtGongj2.EditValue = string.Empty;
                TxtGongj3.EditValue = string.Empty;
                TxtGongj4.EditValue = string.Empty;
                TxtGONGJ5.EditValue = string.Empty;
                TxtGongj6.EditValue = string.Empty;

                LkupDealBankCd.EditValue = string.Empty;
                TxtBankHolder.EditValue = string.Empty;
                TxtAcntNo1.EditValue = string.Empty;
                TxtAcntNo2.EditValue = string.Empty;
                TxtAcntNo3.EditValue = string.Empty;
                TxtNote.EditValue = string.Empty;
            }
            else
            {
                TxtYAKAMT.EditValue = dt.Rows[0]["YAKAMT"]?.ToString();
                TxtBasicSalary.EditValue = dt.Rows[0]["SLR_BSS_AMT"]?.ToString();
                TxtAnnualSalary.EditValue = dt.Rows[0]["SLR_ANL_AMT"]?.ToString();
                TxtDailySalary.EditValue = dt.Rows[0]["SLR_DLY_AMT"]?.ToString();
                TxtSlrTimAmt.EditValue = dt.Rows[0]["SLR_TIM_AMT"]?.ToString();
                TxtBTime1.EditValue = dt.Rows[0]["BASTIME1"]?.ToString();
                TxtBTime2.EditValue = dt.Rows[0]["BASTIME2"]?.ToString();
                TxtSubng1.EditValue = dt.Rows[0]["SUDNG1"]?.ToString();

                TxtGongj1.EditValue = dt.Rows[0]["GONGJ1"]?.ToString();
                TxtGongj2.EditValue = dt.Rows[0]["GONGJ2"]?.ToString();
                TxtGongj3.EditValue = dt.Rows[0]["GONGJ3"]?.ToString();
                TxtGongj4.EditValue = dt.Rows[0]["GONGJ4"]?.ToString();
                TxtGONGJ5.EditValue = dt.Rows[0]["GONGJ5"]?.ToString();
                TxtGongj6.EditValue = dt.Rows[0]["GONGJ6"]?.ToString();

                LkupDealBankCd.EditValue =dt.Rows[0]["DEAL_BANK_CD"]?.ToString();
                TxtBankHolder.EditValue = dt.Rows[0]["ACNT_HDR"]?.ToString();
                TxtAcntNo1.EditValue = dt.Rows[0]["PMNT_ACNT_FST"]?.ToString();
                TxtAcntNo2.EditValue = dt.Rows[0]["PMNT_ACNT_SCD"]?.ToString();
                TxtAcntNo3.EditValue = dt.Rows[0]["PMNT_ACNT_TRD"]?.ToString();
                TxtNote.EditValue = dt.Rows[0]["NOTE"]?.ToString();

                _CHGTEXT1 = dt.Rows[0]["YAKAMT"]?.ToString();
                _CHGTEXT2 = dt.Rows[0]["BASTIME1"]?.ToString();
                _CHGTEXT3 = dt.Rows[0]["BASTIME2"]?.ToString();
            }                          

            Cursor = Cursors.Default;
        }

        private bool SavePaymentInfo(SqlCommand cmd)
        {
            bool result = false;
            string sEMP_ID = string.Empty;
            string sYAKAMT = string.Empty;
            string sSLR_BSS_AMT = string.Empty;
            string sSLR_ANL_AMT = string.Empty;
            string sSLR_DLY_AMT = string.Empty;
            string sSLR_TIM_AMT = string.Empty;
            string sBASTIME1 = string.Empty;
            string sBASTIME2 = string.Empty;
            string sSUDNG1 = string.Empty;
            string sSUDNG2 = string.Empty;
            string sSUDNG3 = string.Empty;
            string sSUDNG4 = string.Empty;
            string sSUDNG5 = string.Empty;
            string sSUDNG6 = string.Empty;
            string sSUDNG7 = string.Empty;
            string sGONGJ1 = string.Empty;
            string sGONGJ2 = string.Empty;
            string sGONGJ3 = string.Empty;
            string sGONGJ4 = string.Empty;
            string sGONGJ5 = string.Empty;
            string sGONGJ6 = string.Empty;
            string sGONGJ7 = string.Empty;
            string sDEAL_BANK_CD = string.Empty;
            string sDEAL_BANK_NM = string.Empty;
            string sACNT_HDR = string.Empty;
            string sPMNT_ACNT_FST = string.Empty;
            string sPMNT_ACNT_SCD = string.Empty;
            string sPMNT_ACNT_TRD = string.Empty;
            string sNOTE = string.Empty;
            string sUSRCD = FmMainToolBar2.drUser["USRCD"]?.ToString();

            double dYAKAMT = 0;
            double dSLR_BSS_AMT = 0;
            double dSLR_ANL_AMT = 0;
            double dSLR_DLY_AMT = 0;
            double dSLR_TIM_AMT = 0;
            double dBASTIME1 = 0;
            double dBASTIME2 = 0;
            double dSUDNG1 = 0;
            double dSUDNG2 = 0;
            double dSUDNG3 = 0;
            double dSUDNG4 = 0;
            double dSUDNG5 = 0;
            double dSUDNG6 = 0;
            double dSUDNG7 = 0;
            double dGONGJ1 = 0;
            double dGONGJ2 = 0;
            double dGONGJ3 = 0;
            double dGONGJ4 = 0;
            double dGONGJ5 = 0;
            double dGONGJ6 = 0;
            double dGONGJ7 = 0;

            StringBuilder strSql = new StringBuilder();

            sEMP_ID = GridViewRetr.GetFocusedRowCellValue("EMP_ID")?.ToString();
            if (string.IsNullOrEmpty(sEMP_ID))
            {
                XtraMessageBox.Show("사원번호가 없습니다.");
                return false;
            }

            sYAKAMT = TxtYAKAMT.EditValue?.ToString();
            sSLR_BSS_AMT = TxtBasicSalary.EditValue?.ToString();
            sSLR_ANL_AMT = TxtAnnualSalary.EditValue?.ToString();
            sSLR_DLY_AMT = TxtDailySalary.EditValue?.ToString();
            sSLR_TIM_AMT = TxtSlrTimAmt.EditValue?.ToString();
            sBASTIME1 = TxtBTime1.EditValue?.ToString();
            sBASTIME2 = TxtBTime2.EditValue?.ToString();

            double.TryParse(sYAKAMT, out dYAKAMT);
            double.TryParse(sSLR_BSS_AMT, out dSLR_BSS_AMT);
            double.TryParse(sSLR_ANL_AMT, out dSLR_ANL_AMT);
            double.TryParse(sSLR_DLY_AMT, out dSLR_DLY_AMT);
            double.TryParse(sSLR_TIM_AMT, out dSLR_TIM_AMT);
            double.TryParse(sBASTIME1, out dBASTIME1);
            double.TryParse(sBASTIME2, out dBASTIME2);

            sSUDNG1 = TxtSubng1.EditValue?.ToString();

            double.TryParse(sSUDNG1, out dSUDNG1);

            sGONGJ1 = TxtGongj1.EditValue?.ToString();
            sGONGJ2 = TxtGongj2.EditValue?.ToString();
            sGONGJ3 = TxtGongj3.EditValue?.ToString();
            sGONGJ4 = TxtGongj4.EditValue?.ToString();
            sGONGJ5 = TxtGONGJ5.EditValue?.ToString();
            sGONGJ6 = TxtGongj6.EditValue?.ToString();

            double.TryParse(sGONGJ1, out dGONGJ1);
            double.TryParse(sGONGJ2, out dGONGJ2);
            double.TryParse(sGONGJ3, out dGONGJ3);
            double.TryParse(sGONGJ4, out dGONGJ4);
            double.TryParse(sGONGJ5, out dGONGJ5);
            double.TryParse(sGONGJ6, out dGONGJ6);

            sDEAL_BANK_CD = LkupDealBankCd.EditValue?.ToString();
            sDEAL_BANK_NM = LkupDealBankCd.Text;
            sACNT_HDR = TxtBankHolder.EditValue?.ToString();
            sPMNT_ACNT_FST = TxtAcntNo1.EditValue?.ToString();
            sPMNT_ACNT_SCD = TxtAcntNo2.EditValue?.ToString();
            sPMNT_ACNT_TRD = TxtAcntNo3.EditValue?.ToString();
            sNOTE = TxtNote.EditValue?.ToString();

            strSql.Clear();
            strSql.AppendLine("SET IDENTITY_INSERT HR_EMP_SALARY ON");
            strSql.AppendLine("IF EXISTS(SELECT* FROM HR_EMP_SALARY WHERE EMP_ID = '"+ sEMP_ID + "')");
            strSql.AppendLine("   BEGIN                                               ");
            strSql.AppendLine("         UPDATE HR_EMP_SALARY");
            strSql.AppendLine("            SET SLR_BSS_AMT = " + dSLR_BSS_AMT);
            strSql.AppendLine("              , YAKAMT = " + dYAKAMT);
            strSql.AppendLine("              , SLR_ANL_AMT  = "+ dSLR_ANL_AMT);
            strSql.AppendLine("              , SLR_DLY_AMT  = "+ dSLR_DLY_AMT);
            strSql.AppendLine("              , SLR_TIM_AMT  = "+ dSLR_TIM_AMT);
            strSql.AppendLine("              , BASTIME1     = "+ dBASTIME1);
            strSql.AppendLine("              , BASTIME2     = "+ dBASTIME2);
            strSql.AppendLine("              , SUDNG1       = "+ dSUDNG1 );
            strSql.AppendLine("              , SUDNG2       = "+ dSUDNG2 );
            strSql.AppendLine("              , SUDNG3       = "+ dSUDNG3 );
            strSql.AppendLine("              , SUDNG4       = "+ dSUDNG4 );
            strSql.AppendLine("              , SUDNG5       = "+ dSUDNG5 );
            strSql.AppendLine("              , SUDNG6       = "+ dSUDNG6 );
            strSql.AppendLine("              , SUDNG7       = "+ dSUDNG7 );
            strSql.AppendLine("              , GONGJ1       = "+ dGONGJ1 );
            strSql.AppendLine("              , GONGJ2       = "+ dGONGJ2 );
            strSql.AppendLine("              , GONGJ3       = "+ dGONGJ3 );
            strSql.AppendLine("              , GONGJ4       = "+ dGONGJ4 );
            strSql.AppendLine("              , GONGJ5       = "+ dGONGJ5 );
            strSql.AppendLine("              , GONGJ6       = "+ dGONGJ6 );
            strSql.AppendLine("              , GONGJ7       = "+ dGONGJ7 );
            strSql.AppendLine("              , DEAL_BANK_CD = '"+ sDEAL_BANK_CD + "'");
            strSql.AppendLine("              , DEAL_BANK_NM = '"+ sDEAL_BANK_NM + "'");
            strSql.AppendLine("              , ACNT_HDR     = '"+ sACNT_HDR + "'");
            strSql.AppendLine("              , PMNT_ACNT_FST= '"+ sPMNT_ACNT_FST + "'");
            strSql.AppendLine("              , PMNT_ACNT_SCD= '"+ sPMNT_ACNT_SCD + "'");
            strSql.AppendLine("              , PMNT_ACNT_TRD= '"+ sPMNT_ACNT_TRD + "'");
            strSql.AppendLine("              , NOTE         = '"+ sNOTE + "'");
            strSql.AppendLine("              , MFY_DT = convert(varchar(19),getdate(),20) ");
	        strSql.AppendLine("              , MFY_ID = '" + sUSRCD + "' ");
            strSql.AppendLine("          WHERE EMP_ID = '"+ sEMP_ID + "'");
            strSql.AppendLine("     END                    ");
            strSql.AppendLine("ELSE                        ");
            strSql.AppendLine("   BEGIN                    ");
            strSql.AppendLine("         INSERT INTO HR_EMP_SALARY ");
            strSql.AppendLine("              (EMP_ID");
            strSql.AppendLine("              , YAKAMT");
            strSql.AppendLine("              , SLR_BSS_AMT");
            strSql.AppendLine("              , SLR_ANL_AMT  ");
            strSql.AppendLine("              , SLR_DLY_AMT  ");
            strSql.AppendLine("              , SLR_TIM_AMT  ");
            strSql.AppendLine("              , BASTIME1     ");
            strSql.AppendLine("              , BASTIME2     ");
            strSql.AppendLine("              , SUDNG1       ");
            strSql.AppendLine("              , SUDNG2       ");
            strSql.AppendLine("              , SUDNG3       ");
            strSql.AppendLine("              , SUDNG4       ");
            strSql.AppendLine("              , SUDNG5       ");
            strSql.AppendLine("              , SUDNG6       ");
            strSql.AppendLine("              , SUDNG7       ");
            strSql.AppendLine("              , GONGJ1       ");
            strSql.AppendLine("              , GONGJ2       ");
            strSql.AppendLine("              , GONGJ3       ");
            strSql.AppendLine("              , GONGJ4       ");
            strSql.AppendLine("              , GONGJ5       ");
            strSql.AppendLine("              , GONGJ6       ");
            strSql.AppendLine("              , GONGJ7       ");
            strSql.AppendLine("              , DEAL_BANK_CD ");
            strSql.AppendLine("              , DEAL_BANK_NM ");
            strSql.AppendLine("              , ACNT_HDR     ");
            strSql.AppendLine("              , PMNT_ACNT_FST");
            strSql.AppendLine("              , PMNT_ACNT_SCD");
            strSql.AppendLine("              , PMNT_ACNT_TRD");
            strSql.AppendLine("              , NOTE         ");
            strSql.AppendLine("              , ENT_DT ");
            strSql.AppendLine("              , ENT_ID )");
            strSql.AppendLine("         VALUES ");
            strSql.AppendLine("              ('" + sEMP_ID + "' ");
            strSql.AppendLine("              , " + dYAKAMT);
            strSql.AppendLine("              , " + dSLR_BSS_AMT);
            strSql.AppendLine("              , " + dSLR_ANL_AMT);
            strSql.AppendLine("              , " + dSLR_DLY_AMT);
            strSql.AppendLine("              , " + dSLR_TIM_AMT);
            strSql.AppendLine("              , " + dBASTIME1);
            strSql.AppendLine("              , " + dBASTIME2);
            strSql.AppendLine("              , " + dSUDNG1);
            strSql.AppendLine("              , " + dSUDNG2);
            strSql.AppendLine("              , " + dSUDNG3);
            strSql.AppendLine("              , " + dSUDNG4);
            strSql.AppendLine("              , " + dSUDNG5);
            strSql.AppendLine("              , " + dSUDNG6);
            strSql.AppendLine("              , " + dSUDNG7);
            strSql.AppendLine("              , " + dGONGJ1);
            strSql.AppendLine("              , " + dGONGJ2);
            strSql.AppendLine("              , " + dGONGJ3);
            strSql.AppendLine("              , " + dGONGJ4);
            strSql.AppendLine("              , " + dGONGJ5);
            strSql.AppendLine("              , " + dGONGJ6);
            strSql.AppendLine("              , " + dGONGJ7);
            strSql.AppendLine("              , '" + sDEAL_BANK_CD + "'");
            strSql.AppendLine("              , '" + sDEAL_BANK_NM + "'");
            strSql.AppendLine("              , '" + sACNT_HDR + "'");
            strSql.AppendLine("              , '" + sPMNT_ACNT_FST + "'");
            strSql.AppendLine("              , '" + sPMNT_ACNT_SCD + "'");
            strSql.AppendLine("              , '" + sPMNT_ACNT_TRD + "'");
            strSql.AppendLine("              , '" + sNOTE + "'");
            strSql.AppendLine("              , convert(varchar(19),getdate(),20) ");
            strSql.AppendLine("              , '" + sUSRCD + "' )");
            strSql.AppendLine("     END");
            strSql.AppendLine("SET IDENTITY_INSERT HR_EMP_SALARY OFF");

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSql.ToString();
            cmd.ExecuteNonQuery();

            result = true;

            return result;
        }

        //private void TxtYAKAMT_EditValueChanged(object sender, EventArgs e)
        //{
        //    SetSudng1();
        //}

        //private void TxtBTime1_EditValueChanged(object sender, EventArgs e)
        //{
        //    SetSudng1();
        //}

        //private void TxtBTime2_EditValueChanged(object sender, EventArgs e)
        //{
        //    SetSudng1();
        //}

        string _CHGTEXT1 = string.Empty;
        private void TxtYAKAMT_Leave(object sender, EventArgs e)
        {
            TextEdit textEdit = (TextEdit)sender;

            if (!_CHGTEXT1.Equals(textEdit.EditValue?.ToString()))
            {
                SetSudng1();
                _CHGTEXT1 = textEdit.EditValue?.ToString();
            }
        }

        string _CHGTEXT2 = string.Empty;
        private void TxtBTime1_Leave(object sender, EventArgs e)
        {
            TextEdit textEdit = (TextEdit)sender;

            if (!_CHGTEXT2.Equals(textEdit.EditValue?.ToString()))
            {
                SetSudng1();
                _CHGTEXT2 = textEdit.EditValue?.ToString();
            }
        }

        string _CHGTEXT3 = string.Empty;
        private void TxtBTime2_Leave(object sender, EventArgs e)
        {
            TextEdit textEdit = (TextEdit)sender;

            if (!_CHGTEXT3.Equals(textEdit.EditValue?.ToString()))
            {
                SetSudng1();
                _CHGTEXT3 = textEdit.EditValue?.ToString();
            }
        }

        //기본급,고정연장수당 계산
        private void SetSudng1()
        {
            string sAmt = TxtYAKAMT.EditValue?.ToString();
            string sBtime1 = TxtBTime1.EditValue?.ToString();
            string sBtime2 = TxtBTime2.EditValue?.ToString();

            double dAmt = 0;
            double dBtime1 = 0;
            double dBtime2 = 0;

            double.TryParse(sAmt, out dAmt); //약정급여
            double.TryParse(sBtime1, out dBtime1); //기본근무시간
            double.TryParse(sBtime2, out dBtime2); //연장근무시간

            double dVal1 = (dAmt * dBtime1) / (dBtime1 + dBtime2); //기본급
            dVal1 = Math.Round(dVal1, 0, MidpointRounding.AwayFromZero);
            double dVal2 = (dAmt * dBtime2) / (dBtime1 + dBtime2); //고정연장수당
            dVal2 = Math.Round(dVal2, 0, MidpointRounding.AwayFromZero);
            double dVal3 = Math.Round(dVal1 / dBtime1, 0, MidpointRounding.AwayFromZero);//시급

            if (double.IsNaN(dVal1))
            {
                dVal1 = 0;
            }

            if (double.IsNaN(dVal2))
            {
                dVal2 = 0;
            }

            if (double.IsNaN(dVal3))
            {
                dVal3 = 0;
            }

            TxtBasicSalary.EditValue = dVal1;
            TxtSubng1.EditValue = dVal2;
            TxtSlrTimAmt.EditValue = dVal3;
        }

        //주민세 계산
        private void TxtGongj1_EditValueChanged(object sender, EventArgs e)
        {
            string sGongj1 = TxtGongj1.EditValue?.ToString();
            double dGongj1 = 0;

            if (double.TryParse(sGongj1, out dGongj1))
            {
                double dVal = dGongj1 / 10;
                dVal = Math.Truncate(dVal / 10) * 10;

                if (double.IsNaN(dVal))
                {
                    dVal = 0;
                }

                TxtGongj2.EditValue = dVal;
            }
        }
        #endregion[급여 부분 Select, Add, Save]

        #region [ 병역 부분 Select, Add, Save ]

        private void GetMilitaryInfo()
        {
            Cursor = Cursors.WaitCursor;

            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.EMP_ID ");
            strSql.AppendLine("      , A.MTR_CY_SIGN ");
            strSql.AppendLine("      , A.MTR_HTH_LVL ");
            strSql.AppendLine("      , A.NON_SRV_RESN ");
            strSql.AppendLine("      , A.MTR_CLS_FCTN ");
            strSql.AppendLine("      , A.MTR_BRCH ");
            strSql.AppendLine("      , A.MTR_RANK ");
            strSql.AppendLine("      , A.MTR_NO ");
            strSql.AppendLine("      , A.MTR_ELMT_DT ");
            strSql.AppendLine("      , A.MTR_DCHG_DT ");
            strSql.AppendLine("      , A.MTR_PSTN ");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , A.ENT_ID ");
            strSql.AppendLine("      , A.ENT_IP ");
            strSql.AppendLine("      , A.MFY_DT ");
            strSql.AppendLine("      , A.MFY_ID ");
            strSql.AppendLine("      , A.MFY_IP ");
            strSql.AppendLine("   FROM HR_EMP_MILITARY A ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count == 0)
            {
                TxtCySign.EditValue = string.Empty;
                TxtHthLevel.EditValue = string.Empty;
                TxtNonSrvResn.EditValue = string.Empty;
                TxtClsFctn.EditValue = string.Empty;
                TxtBrch.EditValue = string.Empty;
                TxtRank.EditValue = string.Empty;
                TxtMtrNo.EditValue = string.Empty;
                DateEditElmt.EditValue = string.Empty;
                DateEditDchg.EditValue = string.Empty;
                TxtMtrPstn.EditValue = string.Empty;
            }
            else
            {

                string sMtrCySign = dt.Rows[0]["MTR_CY_SIGN"]?.ToString();
                string sMtrHthLvl = dt.Rows[0]["MTR_HTH_LVL"]?.ToString();
                string sNonSrvResn = dt.Rows[0]["NON_SRV_RESN"]?.ToString();
                string sMtrClsFctn = dt.Rows[0]["MTR_CLS_FCTN"]?.ToString();
                string sMtrBrch = dt.Rows[0]["MTR_BRCH"]?.ToString();
                string sMtrRank = dt.Rows[0]["MTR_RANK"]?.ToString();
                string sMtrNo = dt.Rows[0]["MTR_NO"]?.ToString();

                string sMtrElmtDt = dt.Rows[0]["MTR_ELMT_DT"]?.ToString();
                string sMtrElmDtFinal = string.Empty;
                if (!string.IsNullOrEmpty(sMtrElmtDt) && sMtrElmtDt.Length == 8)
                {
                    sMtrElmDtFinal = sMtrElmtDt.Substring(0, 4) + "-" + sMtrElmtDt.Substring(4, 2) + "-" + sMtrElmtDt.Substring(6, 2);
                }

                string sMtrDchgDt = dt.Rows[0]["MTR_DCHG_DT"]?.ToString();
                string sMtrDchgDtFinal = string.Empty;
                if (!string.IsNullOrEmpty(sMtrDchgDt) && sMtrElmtDt.Length == 8)
                {
                    sMtrDchgDtFinal = sMtrDchgDt.Substring(0, 4) + "-" + sMtrDchgDt.Substring(4, 2) + "-" + sMtrDchgDt.Substring(6, 2);
                }
                string sMtrPstn = dt.Rows[0]["MTR_PSTN"]?.ToString();


                TxtCySign.EditValue = sMtrCySign;
                TxtHthLevel.EditValue = sMtrHthLvl;
                TxtNonSrvResn.EditValue = sNonSrvResn;
                TxtClsFctn.EditValue = sMtrClsFctn;
                TxtBrch.EditValue = sMtrBrch;
                TxtRank.EditValue = sMtrRank;
                TxtMtrNo.EditValue = sMtrNo;
                DateEditElmt.EditValue = sMtrElmDtFinal;
                DateEditDchg.EditValue = sMtrDchgDtFinal;
                TxtMtrPstn.EditValue = sMtrPstn;
            }

            Cursor = Cursors.Default;
        }

        private bool SaveMilitaryInfo(SqlCommand cmd)
        {
            bool result = false;
            string sEmpId = string.Empty;
            string sMtrCySign = string.Empty;
            string sMtrHthLvl = string.Empty;
            string sNonSrvResn = string.Empty;
            string sMtrClsFctn = string.Empty;
            string sMtrBrch = string.Empty;
            string sMtrRank = string.Empty;
            string sMtrNo = string.Empty;
            string sMtrElmtDt = string.Empty;
            string sMtrDchgDt = string.Empty;
            string sMtrPstn = string.Empty;
            string sEntId = string.Empty;
            string sEntIp = string.Empty;
            string sMfyId = string.Empty;
            string sMfyIp = string.Empty;

            StringBuilder strSql = new StringBuilder();

            sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID")?.ToString();
            if (string.IsNullOrEmpty(sEmpId))
            {
                XtraMessageBox.Show("사원번호가 없습니다.");
                return false;
            }

            sMtrCySign = TxtCySign.EditValue?.ToString();
            sMtrHthLvl = TxtHthLevel.EditValue?.ToString();
            sNonSrvResn = TxtNonSrvResn.EditValue?.ToString();
            sMtrClsFctn = TxtClsFctn.EditValue?.ToString();
            sMtrBrch = TxtBrch.EditValue?.ToString();
            sMtrRank = TxtRank.EditValue?.ToString();
            sMtrNo = TxtMtrNo.EditValue?.ToString();
            sMtrElmtDt = DateEditElmt.EditValue?.ToString();
            if(!sMtrElmtDt.Equals(""))
                sMtrElmtDt = sMtrElmtDt.Replace("-", "").Substring(0, 8);

            sMtrDchgDt = DateEditDchg.EditValue?.ToString();
            if(!sMtrDchgDt.Equals(""))
                sMtrDchgDt = sMtrDchgDt.Replace("-", "").Substring(0, 8);

            sMtrPstn = TxtMtrPstn.EditValue?.ToString();
            sEntId = FmMainToolBar2.drUser["USRCD"]?.ToString();
            sEntIp = Get_MyIP();
            sMfyId = FmMainToolBar2.drUser["USRCD"]?.ToString();
            sMfyIp = Get_MyIP();
            
            strSql.Clear();
            strSql.AppendLine(" ");
            #region mariaDB
            //strSql.AppendLine(" INSERT INTO HR_EMP_MILITARY  ");
            //strSql.AppendLine(" 		  ( EMP_ID ");
            //strSql.AppendLine("           , MTR_CY_SIGN ");
            //strSql.AppendLine("           , MTR_HTH_LVL ");
            //strSql.AppendLine("           , NON_SRV_RESN ");
            //strSql.AppendLine("           , MTR_CLS_FCTN ");
            //strSql.AppendLine("           , MTR_BRCH ");
            //strSql.AppendLine("           , MTR_RANK ");
            //strSql.AppendLine("           , MTR_NO ");
            //strSql.AppendLine("           , MTR_ELMT_DT ");
            //strSql.AppendLine("           , MTR_DCHG_DT ");
            //strSql.AppendLine("           , MTR_PSTN ");
            //strSql.AppendLine("           , ENT_DT ");
            //strSql.AppendLine("           , ENT_ID ");
            //strSql.AppendLine("           , ENT_IP ");
            //strSql.AppendLine("           , MFY_DT ");
            //strSql.AppendLine("           , MFY_ID ");
            //strSql.AppendLine("           , MFY_IP ) ");
            //strSql.AppendLine("      VALUES ");
            //strSql.AppendLine("           ( " + sEmpId + " ");
            //strSql.AppendLine("           , '" + sMtrCySign + "' ");
            //strSql.AppendLine("           , '" + sMtrHthLvl + "' ");
            //strSql.AppendLine("           , '" +sNonSrvResn + "' ");
            //strSql.AppendLine("           , '" +sMtrClsFctn + "' ");
            //strSql.AppendLine("           , '" +sMtrBrch + "' ");
            //strSql.AppendLine("           , '" +sMtrRank + "' ");
            //strSql.AppendLine("           , '" +sMtrNo + "' ");
            //strSql.AppendLine("           , '" +sMtrElmtDt + "' ");
            //strSql.AppendLine("           , '" +sMtrDchgDt + "' ");
            //strSql.AppendLine("           , '" +sMtrPstn + "' ");
            //strSql.AppendLine("           , NOW() ");
            //strSql.AppendLine("           , '" +sEntId + "' ");
            //strSql.AppendLine("           , '" +sEntIp + "' ");
            //strSql.AppendLine("           , NOW() ");
            //strSql.AppendLine("           , '" +sMfyId + "' ");
            //strSql.AppendLine("           , '" +sMfyIp + "') ");
            //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
            //strSql.AppendLine("             MTR_CY_SIGN = '" + sMtrCySign + "' ");
            //strSql.AppendLine("           , MTR_HTH_LVL = '" + sMtrHthLvl + "' ");
            //strSql.AppendLine("           , NON_SRV_RESN = '" + sNonSrvResn + "' ");
            //strSql.AppendLine("           , MTR_CLS_FCTN = '" + sMtrClsFctn + "' ");
            //strSql.AppendLine("           , MTR_BRCH = '" + sMtrBrch + "' ");
            //strSql.AppendLine("           , MTR_RANK = '" + sMtrRank + "' ");
            //strSql.AppendLine("           , MTR_NO = '" + sMtrNo + "' ");
            //strSql.AppendLine("           , MTR_ELMT_DT = '" + sMtrElmtDt + "' ");
            //strSql.AppendLine("           , MTR_DCHG_DT = '" + sMtrDchgDt + "' ");
            //strSql.AppendLine("           , MTR_PSTN = '" + sMtrPstn + "' ");
            //strSql.AppendLine("           , MFY_DT = NOW() ");
            //strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");
            //strSql.AppendLine("           , MFY_IP = '" + sMfyIp + "' ");
            #endregion

            strSql.AppendLine("IF EXISTS(SELECT* FROM HR_EMP_MILITARY WHERE EMP_ID = '"+ sEmpId + "')");
            strSql.AppendLine("   BEGIN                               ");
            strSql.AppendLine("         UPDATE HR_EMP_MILITARY        ");
            strSql.AppendLine("            SET MTR_CY_SIGN = '" + sMtrCySign + "' ");
	        strSql.AppendLine("              , MTR_HTH_LVL = '" + sMtrHthLvl + "' ");
	        strSql.AppendLine("              , NON_SRV_RESN = '" + sNonSrvResn + "' ");
	        strSql.AppendLine("              , MTR_CLS_FCTN = '" + sMtrClsFctn + "' ");
	        strSql.AppendLine("              , MTR_BRCH = '" + sMtrBrch + "' ");
	        strSql.AppendLine("              , MTR_RANK = '" + sMtrRank + "' ");
	        strSql.AppendLine("              , MTR_NO = '" + sMtrNo + "' ");
	        strSql.AppendLine("              , MTR_ELMT_DT = '" + sMtrElmtDt + "' ");
	        strSql.AppendLine("              , MTR_DCHG_DT = '" + sMtrDchgDt + "' ");
	        strSql.AppendLine("              , MTR_PSTN = '" + sMtrPstn + "' ");
	        strSql.AppendLine("              , MFY_DT = convert(varchar(19),getdate(),20) ");
	        strSql.AppendLine("              , MFY_ID = '" + sMfyId + "' ");
	        strSql.AppendLine("              , MFY_IP = '" + sMfyIp + "' ");
            strSql.AppendLine("          WHERE EMP_ID = '"+ sEmpId + "' ");
            strSql.AppendLine("     END                    ");
            strSql.AppendLine("ELSE                        ");
            strSql.AppendLine("   BEGIN");
            strSql.AppendLine("         INSERT INTO HR_EMP_MILITARY  ");
            strSql.AppendLine("              (EMP_ID ");
            strSql.AppendLine("              , MTR_CY_SIGN ");
            strSql.AppendLine("              , MTR_HTH_LVL ");
            strSql.AppendLine("              , NON_SRV_RESN ");
            strSql.AppendLine("              , MTR_CLS_FCTN ");
            strSql.AppendLine("              , MTR_BRCH ");
            strSql.AppendLine("              , MTR_RANK ");
            strSql.AppendLine("              , MTR_NO ");
            strSql.AppendLine("              , MTR_ELMT_DT ");
            strSql.AppendLine("              , MTR_DCHG_DT ");
            strSql.AppendLine("              , MTR_PSTN ");
            strSql.AppendLine("              , ENT_DT ");
            strSql.AppendLine("              , ENT_ID ");
            strSql.AppendLine("              , ENT_IP ) ");
            strSql.AppendLine("            VALUES ");
            strSql.AppendLine("                 (" + sEmpId + " ");
            strSql.AppendLine("                 , '" + sMtrCySign + "' ");
            strSql.AppendLine("                 , '" + sMtrHthLvl + "' ");
            strSql.AppendLine("                 , '" +sNonSrvResn + "' ");
            strSql.AppendLine("                 , '" +sMtrClsFctn + "' ");
            strSql.AppendLine("                 , '" +sMtrBrch + "' ");
            strSql.AppendLine("                 , '" +sMtrRank + "' ");
            strSql.AppendLine("                 , '" +sMtrNo + "' ");
            strSql.AppendLine("                 , '" +sMtrElmtDt + "' ");
            strSql.AppendLine("                 , '" +sMtrDchgDt + "' ");
            strSql.AppendLine("                 , '" +sMtrPstn + "' ");
            strSql.AppendLine("                 , convert(varchar(19),getdate(),20) ");
            strSql.AppendLine("                 , '" +sEntId + "' ");
            strSql.AppendLine("                 , '" +sEntIp + "') ");
            strSql.AppendLine("     END");

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSql.ToString();
            cmd.ExecuteNonQuery();

            result = true;

            return result;
        }

        #endregion[병역 부분 Select, Add, Save]

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (XtTControl.SelectedTabPage.Name.Equals("XtTPerson"))
            {
                if (XtraMessageBox.Show("새로운 사원을 추가하시겠습니까? \r\n 현재 선택한 사원정보를 수정하셨다면 No를 선택하시고 저장버튼을 눌러주세요."
                    , "사원 추가여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {

                    /*
                     * #00001 전체 최대값으로 채번 삭제
                     */
                    //StringBuilder strSql = new StringBuilder();

                    //strSql.Clear();
                    //strSql.AppendLine(" SELECT MAX(A.EMP_ID) AS MAX_VALUE ");
                    //strSql.AppendLine("   FROM HR_EMP_BASIS A ");

                    //DataTable dt = MySqlDb.GetDataTable(MySqlDb.dbCon, strSql.ToString());

                    //double sResult = 0;
                    //if (String.IsNullOrEmpty(dt.Rows[0]["MAX_VALUE"].ToString()))
                    //{
                    //    sResult = 1;
                    //}
                    //else
                    //{
                    //    sResult = Convert.ToDouble(dt.Rows[0]["MAX_VALUE"]) + 15;
                    //}
                    
                    GridViewRetr.AddNewRow();
                    //GridViewRetr.SetFocusedRowCellValue("EMP_ID", sResult.ToString());

                    GetBasisInfo();
                    GetPersonalInfo();
                    GetCareerInfo();
                    GetCertificationInfo();
                    GetFamilyInfo();
                    GetEducationInfo();
                    GetPaymentInfo();
                    GetMilitaryInfo();

                    ChkAutoNo.Enabled = true;
                    TxtEMPId.EditValue = "";
                }
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPCarrer"))
            {
                AddCareerInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPCert"))
            {
                AddCertificationInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPFamily"))
            {
                AddFamilyInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPEdctn"))
            {
                AddEducationInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPPay"))
            {

            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPMltrSvc"))
            {

            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                bool result = false;

                if (XtTControl.SelectedTabPage.Name.Equals("XtTPerson"))
                {
                    /*
                     * #00001 중복체크 추가
                     */
                    if (!SaveBasisInfo(cmd))
                    {
                        DBConn.dbTran.Rollback();
                        DBConn.dbTran = null;

                        return;
                    }

                    result = SavePersonalInfo(cmd);

                   
                    if (ComGrid.GridDataSourceUpdate(GridViewBigo))
                    {
                        result = SaveBigo(cmd);
                    }

                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPCarrer"))
                {
                    if (ComGrid.GridDataSourceUpdate(GridViewCareer))
                    {
                        result = SaveCareerInfo(cmd);
                    }
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPCert"))
                {
                    if (ComGrid.GridDataSourceUpdate(GridViewCert))
                    {
                        result = SaveCertificationInfo(cmd);
                    }
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPFamily"))
                {
                    if (ComGrid.GridDataSourceUpdate(GridViewFamily))
                    {
                        result = SaveFamilyInfo(cmd);
                    }
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPEdctn"))
                {
                    if (ComGrid.GridDataSourceUpdate(GridViewEdu))
                    {
                        result = SaveEducationInfo(cmd);
                    }
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPPay"))
                {
                    result = SavePaymentInfo(cmd);
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPMltrSvc"))
                {
                    result = SaveMilitaryInfo(cmd);   
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                if (result)
                {
                    MessageBox.Show("저장을 완료했습니다.");

                    int idx = GridViewRetr.FocusedRowHandle;
                    BtnRetr_Click(null, null);
                    GridViewRetr.FocusedRowHandle = idx;
                }

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

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (XtTControl.SelectedTabPage.Name.Equals("XtTPerson"))
            {
                DelAllData();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPPay"))
            {
                DelAllData();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPMltrSvc"))
            {
                DelAllData();
            }
            else
            {
                DelTabData();
            }

            
        }

        private void DelAllData()
        {
            try
            {
                string sEmpid = GridViewRetr.GetFocusedRowCellValue("EMP_ID")?.ToString();
                string sEmpNm = GridViewRetr.GetFocusedRowCellValue("EMP_NM")?.ToString();

                if (XtraMessageBox.Show("사원번호 : " + sEmpid + "\r\n사원명 : " + sEmpNm
                    + "\r\n선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                    , "인사관리 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM HR_EMP_BASIS    ");
                strSql.AppendLine("  WHERE EMP_ID = '"+ sEmpid + "'          ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM HR_EMP_ACADEMIC ");
                strSql.AppendLine("  WHERE EMP_ID = '" + sEmpid + "'          ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM HR_EMP_CAREER   ");
                strSql.AppendLine("  WHERE EMP_ID = '" + sEmpid + "'          ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM HR_EMP_CERT     ");
                strSql.AppendLine("  WHERE EMP_ID = '" + sEmpid + "'          ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM HR_EMP_FAMILY   ");
                strSql.AppendLine("  WHERE EMP_ID = '" + sEmpid + "'          ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM HR_EMP_MILITARY ");
                strSql.AppendLine("  WHERE EMP_ID = '" + sEmpid + "'          ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM HR_EMP_PERSONAL ");
                strSql.AppendLine("  WHERE EMP_ID = '" + sEmpid + "'          ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM HR_EMP_SALARY   ");
                strSql.AppendLine("  WHERE EMP_ID = '" + sEmpid + "'          ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("삭제를 완료했습니다.");

                int idx = GridViewRetr.FocusedRowHandle;
                BtnRetr_Click(null, null);
                GridViewRetr.FocusedRowHandle = idx-1;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void DelTabData()
        {
            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                if (XtTControl.SelectedTabPage.Name.Equals("XtTPCarrer"))
                {
                    if (!DeleteCareerInfo(cmd))
                    {
                        DBConn.dbTran.Rollback();
                        DBConn.dbTran = null;
                        return;
                    }
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPCert"))
                {
                    if (!DeleteCertificationInfo(cmd))
                    {
                        DBConn.dbTran.Rollback();
                        DBConn.dbTran = null;
                        return;
                    }
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPFamily"))
                {
                    if (!DeleteFamilyInfo(cmd))
                    {
                        DBConn.dbTran.Rollback();
                        DBConn.dbTran = null;
                        return;
                    }
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPEdctn"))
                {
                    if (!DeleteEducationInfo(cmd))
                    {
                        DBConn.dbTran.Rollback();
                        DBConn.dbTran = null;
                        return;
                    }
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("삭제를 완료했습니다.");

                int idx = GridViewRetr.FocusedRowHandle;
                BtnRetr_Click(null, null);
                GridViewRetr.FocusedRowHandle = idx;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void XtTControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            if (e.Page.Name.Equals("XtTPerson"))
            {
                GetBasisInfo();
                GetPersonalInfo();
            }
            else if (e.Page.Name.Equals("XtTPCarrer"))
            {
                GetCareerInfo();
            }
            else if (e.Page.Name.Equals("XtTPCert"))
            {
                GetCertificationInfo();
            }
            else if (e.Page.Name.Equals("XtTPFamily"))
            {
                GetFamilyInfo();
            }
            else if (e.Page.Name.Equals("XtTPEdctn"))
            {
                GetEducationInfo();
            }
            else if (e.Page.Name.Equals("XtTPPay"))
            {
                GetPaymentInfo();
            }
            else if (e.Page.Name.Equals("XtTPMltrSvc"))
            {
                GetMilitaryInfo();
            }
        }
        
        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (e.FocusedRowHandle < 0)
            {
                return;
            }
            else
            {
                GetBasisInfo();

                if (XtTControl.SelectedTabPage.Name.Equals("XtTPerson"))
                {
                    GetPersonalInfo();
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPCarrer"))
                {
                    GetCareerInfo();
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPCert"))
                {
                    GetCertificationInfo();
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPFamily"))
                {
                    GetFamilyInfo();
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPEdctn"))
                {
                    GetEducationInfo();
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPPay"))
                {
                    GetPaymentInfo();
                }
                else if (XtTControl.SelectedTabPage.Name.Equals("XtTPMltrSvc"))
                {
                    GetMilitaryInfo();
                }

                if (string.IsNullOrEmpty(GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString()))
                {
                    ChkAutoNo.Enabled = true;
                }
                else
                {
                    ChkAutoNo.EditValue = "Y";
                    TxtEMPId.ReadOnly = true;
                    ChkAutoNo.Enabled = false;
                }
            }
        }

        private void RdgbRetrGb_EditValueChanged(object sender, EventArgs e)
        {
            GetGridRetr();
            GetBasisInfo();

            if (XtTControl.SelectedTabPage.Name.Equals("XtTPerson"))
            {
                GetPersonalInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPCarrer"))
            {
                GetCareerInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPCert"))
            {
                GetCertificationInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPFamily"))
            {
                GetFamilyInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPEdctn"))
            {
                GetEducationInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPPay"))
            {
                GetPaymentInfo();
            }
            else if (XtTControl.SelectedTabPage.Name.Equals("XtTPMltrSvc"))
            {
                GetMilitaryInfo();
            }
        }

        public string Get_MyIP()
        {
            string strHostName = "";
            strHostName = System.Net.Dns.GetHostName();
            IPHostEntry ipEntry = System.Net.Dns.GetHostEntry(strHostName);
            IPAddress[] addr = ipEntry.AddressList;
            return addr[addr.Length - 1].ToString();
        }

        private void GridViewCert_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName.Equals("ISSUE_YMD"))
            {
                if(e.Value.ToString().Length == 8)
                {
                    string sTemp = e.Value.ToString();
                    string sResult = sTemp.Substring(0, 4) + "-" + sTemp.Substring(4, 2) + "-" + sTemp.Substring(6, 2);
                    e.DisplayText = sResult;
                }
            }
            if (e.Column.FieldName.Equals("END_YMD"))
            {
                if (e.Value.ToString().Length == 8)
                {
                    string sTemp = e.Value.ToString();
                    string sResult = sTemp.Substring(0, 4) + "-" + sTemp.Substring(4, 2) + "-" + sTemp.Substring(6, 2);
                    e.DisplayText = sResult;
                }
            }
        }

        private void GridViewCareer_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName.Equals("STRT_YMD"))
            {
                if (e.Value.ToString().Length == 8)
                {
                    string sTemp = e.Value.ToString();
                    string sResult = sTemp.Substring(0, 4) + "-" + sTemp.Substring(4, 2) + "-" + sTemp.Substring(6, 2);
                    e.DisplayText = sResult;
                }
            }
            if (e.Column.FieldName.Equals("END_YMD"))
            {
                if (e.Value.ToString().Length == 8)
                {
                    string sTemp = e.Value.ToString();
                    string sResult = sTemp.Substring(0, 4) + "-" + sTemp.Substring(4, 2) + "-" + sTemp.Substring(6, 2);
                    e.DisplayText = sResult;
                }
            }
        }

        private void LkupEditDeptCd_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                GetGridRetr();
            }
            else if(e.KeyCode == Keys.Tab)
            {
                GetGridRetr();
            }
        }
        
        private void RdgbMarryYn_EditValueChanged(object sender, EventArgs e)
        {
            string sChk = RdgbMarryYn.EditValue?.ToString();

            if (!string.IsNullOrEmpty(sChk))
            {
                if (RdgbMarryYn.EditValue.ToString().Equals("N"))
                {
                    DateEditMarry.ReadOnly = true;
                }
                else if (RdgbMarryYn.EditValue.ToString().Equals("Y"))
                {
                    DateEditMarry.ReadOnly = false;
                }
            }
        }

        private void RdgbEmplGb_EditValueChanged(object sender, EventArgs e)
        {
            string sChk = RdgbEmplGb.EditValue?.ToString();

            if (!string.IsNullOrEmpty(sChk))
            {
                if (RdgbEmplGb.EditValue.ToString().Equals("Y"))
                {
                    LkupRetireResnCd.ReadOnly = true;
                    TxtRetireResn.ReadOnly = true;
                }
                else if (RdgbEmplGb.EditValue.ToString().Equals("N"))
                {
                    LkupRetireResnCd.ReadOnly = false;
                    TxtRetireResn.ReadOnly = false;
                }
            }
        }
        
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void AccPersonnelRecordsDev_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {
                BtnDelete_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                //BtnExcel_Click(null, null);
            }
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridFamily_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void XtTControl_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridCareer_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
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

        private void AccPersonnelRecordsDev_TextChanged(object sender, EventArgs e)
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

        /*
         * #00001
         */
        private void ChkAutoNo_CheckedChanged(object sender, EventArgs e)
        {
            string sChkAuto = ChkAutoNo.EditValue?.ToString();
            if (sChkAuto.Equals("Y")){
                TxtEMPId.ReadOnly = true;
            }
            else
            {
                TxtEMPId.ReadOnly = false;
            }
        }

        /*
         * #00001
         */
        private void LkupDeptCd_EditValueChanged(object sender, EventArgs e)
        {
            string sChkAuto = ChkAutoNo.EditValue?.ToString();
            string sDeptCd = LkupDeptCd.EditValue?.ToString();

            if (sChkAuto.Equals("N"))
                return;

            string sEmpid = GridViewRetr.GetFocusedRowCellValue("EMP_ID")?.ToString();

            if (sEmpid == null || !sEmpid.Equals(""))
                return;

            if (sDeptCd != null && !sDeptCd.Equals(""))
            {
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine("SELECT MAX(A.EMP_ID)+1 AS MAX_VALUE");
                strSql.AppendLine("  FROM HR_EMP_BASIS A");
                strSql.AppendLine(" WHERE EMP_ID LIKE '" + sDeptCd.Substring(0,1) + "%'");
                strSql.AppendLine(" AND EMPL_GB = 'Y'");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt != null)
                {
                    if (!string.IsNullOrEmpty(dt.Rows[0]["MAX_VALUE"].ToString()))
                        TxtEMPId.EditValue = dt.Rows[0]["MAX_VALUE"];
                    else
                        TxtEMPId.EditValue = sDeptCd.Substring(0, 1) + "000";
                }
            }
        }

        private void RepoFamilyGLkupFamlRelatCd_EditValueChanged(object sender, EventArgs e)
        {
            LookUpEditBase lkup = (LookUpEditBase)sender;

            string sVal = lkup.EditValue?.ToString();

            GridViewFamily.SetFocusedRowCellValue(GridColFamilyFamlRelationCd, sVal);
        }

        byte[] _SIGNIMG = null;
        private void BtnGetImg_Click(object sender, EventArgs e)
        {
            try
            {
                using (XtraOpenFileDialog openFileDialog = new XtraOpenFileDialog())
                {
                    openFileDialog.Title = "사인 이미지를 선택해주세요.";
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png";
                    openFileDialog.FilterIndex = 1;
                    openFileDialog.RestoreDirectory = true;
                    openFileDialog.Multiselect = false;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Get the path of specified file
                        string filePath = openFileDialog.FileName;

                        Image img = Image.FromFile(filePath);

                        picImg.EditValue = img;

                        MemoryStream ms = new MemoryStream();
                        img.Save(ms, img.RawFormat);
                        _SIGNIMG = ms.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        //자격증명 일자
        private void RepoCertDateEditIssue_EditValueChanged(object sender, EventArgs e)
        {
            DateEdit dateEdit = (DateEdit)sender;

            string sDate = dateEdit.EditValue?.ToString();

            if (DateTime.TryParse(sDate, out DateTime dt))
            {
                GridViewCert.SetRowCellValue(GridViewCert.FocusedRowHandle, GridViewCert.FocusedColumn, dt.ToString("yyyy-MM-dd"));
            }
        }
        private void RepoCertDateEditIssue_Leave(object sender, EventArgs e)
        {
            DateEdit dateEdit = (DateEdit)sender;

            string sDate = dateEdit.EditValue?.ToString();

            if (DateTime.TryParse(sDate, out DateTime dt))
            {
                GridViewCert.SetRowCellValue(GridViewCert.FocusedRowHandle, GridViewCert.FocusedColumn, dt.ToString("yyyy-MM-dd"));
            }
        }

        //학력사항 일자
        private void RepoEduDateEditStrtEnt_EditValueChanged(object sender, EventArgs e)
        {
            DateEdit dateEdit = (DateEdit)sender;

            string sDate = dateEdit.EditValue?.ToString();

            if (DateTime.TryParse(sDate, out DateTime dt))
            {
                GridViewEdu.SetRowCellValue(GridViewEdu.FocusedRowHandle, GridViewEdu.FocusedColumn, dt.ToString("yyyy-MM-dd"));
            }
        }
        private void RepoEduDateEditStrtEnt_Leave(object sender, EventArgs e)
        {
            DateEdit dateEdit = (DateEdit)sender;

            string sDate = dateEdit.EditValue?.ToString();

            if (DateTime.TryParse(sDate, out DateTime dt))
            {
                GridViewEdu.SetRowCellValue(GridViewEdu.FocusedRowHandle, GridViewEdu.FocusedColumn, dt.ToString("yyyy-MM-dd"));
            }
        }

        //경력사항 일자
        private void RepoCareerDateEditStrtEnd_EditValueChanged(object sender, EventArgs e)
        {
            DateEdit dateEdit = (DateEdit)sender;

            string sDate = dateEdit.EditValue?.ToString();

            if (DateTime.TryParse(sDate, out DateTime dt))
            {
                GridViewCareer.SetRowCellValue(GridViewCareer.FocusedRowHandle, GridViewCareer.FocusedColumn, dt.ToString("yyyy-MM-dd"));
            }
        }
        private void RepoCareerDateEditStrtEnd_Leave(object sender, EventArgs e)
        {
            DateEdit dateEdit = (DateEdit)sender;

            string sDate = dateEdit.EditValue?.ToString();

            if (DateTime.TryParse(sDate, out DateTime dt))
            {
                GridViewCareer.SetRowCellValue(GridViewCareer.FocusedRowHandle, GridViewCareer.FocusedColumn, dt.ToString("yyyy-MM-dd"));
            }
        }

        private void TxtSubng1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
                e.SuppressKeyPress = true;
        }

        private void TxtSaNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void btn_Add_1_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            string sEmpId = GridViewRetr.GetFocusedRowCellValue("EMP_ID") == null ? string.Empty : GridViewRetr.GetFocusedRowCellValue("EMP_ID").ToString();

            GridViewBigo.AddNewRow();
            GridViewBigo.SetFocusedRowCellValue("EMP_ID", sEmpId);
            GridViewBigo.SetFocusedRowCellValue("STRT_YMD", DateTime.Now.ToString().Substring(0, 10));
            GridViewBigo.SetFocusedRowCellValue("END_YMD", DateTime.Now.ToString().Substring(0, 10));
            //GridViewCareer.ShowEditForm();

            Cursor = Cursors.Default;
        }

        
        private bool SaveBigo(SqlCommand cmd)
        {
            bool result = false;

            if (GridViewBigo.RowCount > 0)
            {
                DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridBigo.DataSource);
                DataTable dtMerge9 = dsSave.Tables[0];

                string sEmpId = string.Empty;
                double dSeq = 0;
                string sStrtYmd = string.Empty;
                string sEndYmd = string.Empty;
                string sContents = string.Empty;
                string sBigo = string.Empty;
                
                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < dtMerge9.Rows.Count; i++)
                {
                    sEmpId = dtMerge9.Rows[i]["EMP_ID"].ToString();

                    if (String.IsNullOrEmpty(dtMerge9.Rows[i]["SEQ"].ToString()))
                    {
                        strSql.Clear();
                        strSql.AppendLine(" SELECT MAX(A.SEQ) AS MAX_VALUE");
                        strSql.AppendLine("   FROM HR_EMP_BIGO A ");
                        strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");

                        DataTable dtBigoChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                        if (dtBigoChk.Rows.Count == 0 || String.IsNullOrEmpty(dtBigoChk.Rows[0]["MAX_VALUE"].ToString()))
                        {
                            dSeq = 1;
                        }
                        else
                        {
                            dSeq = Convert.ToDouble(dtBigoChk.Rows[0]["MAX_VALUE"].ToString()) + 1;
                        }
                    }
                    else
                    {
                        dSeq = Convert.ToDouble(dtMerge9.Rows[i]["SEQ"].ToString());
                    }

                    sStrtYmd = dtMerge9.Rows[i]["STRT_YMD"].ToString().Replace("-", "").Substring(0, 8);
                    sEndYmd = dtMerge9.Rows[i]["END_YMD"].ToString().Replace("-", "").Substring(0, 8);
                    sContents = dtMerge9.Rows[i]["CONTENTS"].ToString();
                    sBigo = dtMerge9.Rows[i]["BIGO"].ToString();
                    
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    
                    strSql.AppendLine("IF EXISTS(SELECT* FROM HR_EMP_BIGO WHERE EMP_ID = '" + sEmpId + "' AND SEQ = " + dSeq + ")");
                    strSql.AppendLine("   BEGIN                                                           ");
                    strSql.AppendLine("         UPDATE HR_EMP_BIGO                                      ");
                    strSql.AppendLine("            SET STRT_YMD = '" + sStrtYmd + "' ");
                    strSql.AppendLine("              , END_YMD = '" + sEndYmd + "' ");
                    strSql.AppendLine("              , CONTENTS = '" + sContents + "' ");
                    strSql.AppendLine("              , BIGO = '" + sBigo + "' ");
                    strSql.AppendLine("          WHERE EMP_ID = '" + sEmpId + "'");
                    strSql.AppendLine("            AND SEQ = " + dSeq + "");
                    strSql.AppendLine("     END               ");
                    strSql.AppendLine("ELSE                   ");
                    strSql.AppendLine("   BEGIN               ");
                    strSql.AppendLine("         INSERT INTO HR_EMP_BIGO ");
                    strSql.AppendLine("              (EMP_ID ");
                    strSql.AppendLine("              , SEQ ");
                    strSql.AppendLine("              , STRT_YMD ");
                    strSql.AppendLine("              , END_YMD ");
                    strSql.AppendLine("              , CONTENTS ");
                    strSql.AppendLine("              , BIGO ");
                    strSql.AppendLine("              ) ");
                    strSql.AppendLine("         VALUES  ");
                    strSql.AppendLine("              ('" + sEmpId + "' ");
                    strSql.AppendLine("              , " + dSeq + " ");
                    strSql.AppendLine("              , '" + sStrtYmd + "' ");
                    strSql.AppendLine("              , '" + sEndYmd + "' ");
                    strSql.AppendLine("              , '" + sContents + "' ");
                    strSql.AppendLine("              , '" + sBigo + "' ");
                    strSql.AppendLine("              )  ");
                    strSql.AppendLine("     END");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                }

                result = true;
            }
            else
            {
                result = true;
            }
            return result;
        }

        private void GridViewBigo_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName.Equals("STRT_YMD"))
            {
                if (e.Value.ToString().Length == 8)
                {
                    string sTemp = e.Value.ToString();
                    string sResult = sTemp.Substring(0, 4) + "-" + sTemp.Substring(4, 2) + "-" + sTemp.Substring(6, 2);
                    e.DisplayText = sResult;
                }
            }
            if (e.Column.FieldName.Equals("END_YMD"))
            {
                if (e.Value.ToString().Length == 8)
                {
                    string sTemp = e.Value.ToString();
                    string sResult = sTemp.Substring(0, 4) + "-" + sTemp.Substring(4, 2) + "-" + sTemp.Substring(6, 2);
                    e.DisplayText = sResult;
                }
            }
        }

        private void repositoryItemDateEdit1_EditValueChanged(object sender, EventArgs e)
        {
            DateEdit dateEdit = (DateEdit)sender;

            string sDate = dateEdit.EditValue?.ToString();

            if (DateTime.TryParse(sDate, out DateTime dt))
            {
                GridViewCareer.SetRowCellValue(GridViewCareer.FocusedRowHandle, GridViewCareer.FocusedColumn, dt.ToString("yyyy-MM-dd"));
            }
        }

        private void btn_Del_1_Click(object sender, EventArgs e)
        {
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            StringBuilder strSql = new StringBuilder();

            string sEmpid = GridViewBigo.GetFocusedRowCellValue("EMP_ID")?.ToString();
            double sSeq = Convert.ToDouble(GridViewBigo.GetFocusedRowCellValue("SEQ")?.ToString());
            int focusedRow = GridViewBigo.FocusedRowHandle;

            if (XtraMessageBox.Show("삭제하시겠습니까?", "비고삭제", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            strSql.Clear();
            strSql.AppendLine(" DELETE FROM HR_EMP_BIGO   ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpid + "'          ");
            strSql.AppendLine("  AND   SEQ    = '" + sSeq + "'          ");

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = strSql.ToString();
            cmd.ExecuteNonQuery();

            GridViewBigo.DeleteRow(focusedRow);
            //GridViewBigo.GridControl.RefreshDataSource();
        }
    }
}