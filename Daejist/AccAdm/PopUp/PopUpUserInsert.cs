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
using System.Data.SqlClient;

/*
 * 작성자: 모름
 * 작성일: 모름
 * --------------------History---------------------
 * 수정자: 정은영
 * 수정일자: 2021-05-27
 * ID: #001
 * 내용: 1. 사용자명 키다운(엔터)이벤트 > 리브이벤트로 변경
 *       2. FmMainToolBar2.UserID > dUserID
 *       3. 인사코드 추가
 * 
 * 수정자: 정은영
 * 수정일자: 2021-05-28
 * ID: #002
 * 내용: 1. 사용자수정시 권한등급 바꼈을때만 권한 적용.
 *       2. ID중복 체크 추가
 */

namespace AccAdm
{
    public partial class PopUpUserInsert : DevExpress.XtraEditors.XtraForm
    {
        public PopUpUserInsert()
        {
            InitializeComponent();
        }

        public DataRow drUserInfo { get; set; }

        private void PopUpUserInsert_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            DataTable dtDept = GetLookUpData("1", "1", "", "Y");
            DataTable dtJobGrade = GetLookUpData("2", "1", "", "Y");
            DataTable dtEmpNm = GetLookUpData("3", "", "", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupDept, dtDept, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupJobGrade, dtJobGrade, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupEmpNm, dtEmpNm, "CD", "NM", "Y");

            if (drUserInfo != null)
            {
                string sUserCD = drUserInfo["USRCD"]?.ToString();
                string sUserID = drUserInfo["USRID"]?.ToString();
                string sUserNM = drUserInfo["USRNM"]?.ToString();
                string sPassWD = drUserInfo["PASSWD"]?.ToString();
                if (!string.IsNullOrEmpty(sPassWD))
                {
                    sPassWD = ComnEtcFunc.Decrypt(sPassWD, ComnEtcFunc._SECRET_KEY2);
                }

                string sDeptCd = drUserInfo["DEPTCD"]?.ToString();
                string sJobGrade = drUserInfo["JKWICD"]?.ToString();
                string sMoblCd = drUserInfo["MOBLNO"]?.ToString();
                string sUseYn = drUserInfo["USEYN"]?.ToString();
                string sUserLvl = drUserInfo["USLVL"]?.ToString();
                string sRemark = drUserInfo["RK"]?.ToString();
                string sIsptYn = drUserInfo["ISPT_YN"]?.ToString();

                TxtUserId.EditValue = sUserID;
                LkupEmpNm.Text = sUserNM;
                TxtPw.EditValue = sPassWD;
                LkupDept.EditValue = sDeptCd;
                LkupJobGrade.EditValue = sJobGrade;
                TxtCallNo.EditValue = sMoblCd;
                ChkUse.EditValue = sUseYn;
                CboGrade.EditValue = sUserLvl;
                TxtMemo.EditValue = sRemark;
                ChkIsptYn.EditValue = sIsptYn;
            }
            
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

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.DEPT_CD AS CD");
                strSql.AppendLine("      , A.DEPT_NM AS NM");
                strSql.AppendLine("      , A.DEPT_CD AS SEQ");
                strSql.AppendLine("   FROM ACC_DEPT_CD A");
                strSql.AppendLine("  WHERE DEPT_CD <> '0000'");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'GRADE_CD'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.EMP_ID AS CD");
                strSql.AppendLine("      , A.EMP_NM AS NM");
                strSql.AppendLine("      , A.EMP_ID AS SEQ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A");
                strSql.AppendLine("  WHERE A.EMPL_GB = 'Y'");
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
        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            TxtCallNo.Text = "";
            TxtMemo.Text = "";
            TxtPw.Text = "";
            TxtUserId.Text = "";
            LkupEmpNm.Text = "";
            LkupDept.EditValue = null;
            LkupJobGrade.EditValue = null;
            CboGrade.EditValue = null;
            ChkUse.EditValue = "N";

        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                string sUserID = TxtUserId.EditValue?.ToString();
                string sPassWD = TxtPw.EditValue?.ToString();
                if (string.IsNullOrEmpty(sUserID) || string.IsNullOrEmpty(sPassWD))
                {
                    XtraMessageBox.Show("사용자ID와 PW는 필수항목입니다");
                    TxtUserId.Focus();
                    return;
                }

                string sGrade = CboGrade.EditValue?.ToString();
                if (string.IsNullOrEmpty(sGrade))
                {
                    XtraMessageBox.Show("권한 등급을 선택해주세요.");
                    CboGrade.Focus();
                    return;
                }

                if (!string.IsNullOrEmpty(sPassWD))
                {
                    sPassWD = ComnEtcFunc.Encrypt(sPassWD, ComnEtcFunc._SECRET_KEY2);
                }

                double dUserCD = 0;

                string sUserNm = LkupEmpNm.Text;
                string sInsaNo = LkupEmpNm.EditValue?.ToString();
                string sDeptCd = LkupDept.EditValue?.ToString();
                string sJobGrade = LkupJobGrade.EditValue?.ToString();
                string sMoblCd = TxtCallNo.EditValue?.ToString();
                string sUseYn = ChkUse.EditValue?.ToString();
                string sRemark = TxtMemo.EditValue?.ToString();
                string sIsptYn = ChkIsptYn.EditValue?.ToString();
                string sEntId = FmMainToolBar2.UserID;
                string sMfyId = FmMainToolBar2.UserID;

                Cursor = Cursors.WaitCursor;


                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT COUNT(*) AS CNT ");
                strSql.AppendLine("   FROM ZUSRLST A");
                strSql.AppendLine("  WHERE USRID = '"+ sUserID + "'");

                DataTable chkUsrId = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                double dChkUsrid = double.Parse(chkUsrId.Rows[0]["CNT"]?.ToString());
                string sUsinfoid = string.Empty;

                //#002
                if (drUserInfo != null)
                    sUsinfoid = drUserInfo["USRID"].ToString();

                if (dChkUsrid > 0 && !sUserID.Equals(sUsinfoid))
                {
                    XtraMessageBox.Show("중복된 ID를 사용할 수 없습니다.");
                    TxtUserId.Focus();
                    TxtUserId.SelectAll();
                    Cursor = Cursors.Default;
                    return;
                }

                if (drUserInfo != null)
                {
                    dUserCD = Convert.ToDouble(drUserInfo["USRCD"]);
                }
                else
                {
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT MAX(A.USRCD) AS MAX_VALUE ");
                    strSql.AppendLine("   FROM ZUSRLST A");

                    DataTable dtMax = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    if (dtMax.Rows.Count == 0)
                    {
                        dUserCD = 1000;
                    }
                    else if (dtMax.Rows.Count > 0)
                    {
                        dUserCD = Convert.ToDouble(dtMax.Rows[0]["MAX_VALUE"]) + 100;
                    }
                }

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                #region LOG
                //string sLOG = string.Empty;
                //string sEDIT_KIND = string.Empty;

                //strSql.Clear();
                //strSql.AppendLine(" ");
                //strSql.AppendLine(" SELECT * ");
                //strSql.AppendLine("   FROM ZUSRLST ");
                //strSql.AppendLine("  WHERE USRCD = " + dUserCD + " ");

                //DataTable dtPrv = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                //if(dtPrv.Rows.Count > 0)
                //{
                //    sLOG = string.Concat("[사용자관리]사용자코드 : ", dtPrv.Rows[0]["USRCD"],
                //        ", 사용자ID : ", dtPrv.Rows[0]["USRID"], " ▶ ", sUserID,
                //        ", 패스워드 : ", dtPrv.Rows[0]["PASSWD"], " ▶ ", sPassWD,
                //        ", 유저명 : ", dtPrv.Rows[0]["USRNM"], " ▶ ", sUserNm,
                //        ", 부서코드 : ", dtPrv.Rows[0]["DEPTCD"], " ▶ ", sDeptCd,
                //        ", 직위코드 : ", dtPrv.Rows[0]["JKWICD"], " ▶ ", sJobGrade,
                //        //#001 인사코드추가   
                //        ", 인사코드 : ", dtPrv.Rows[0]["INSANO"], " ▶ ", sInsaNo,
                //        ", 연락처 : ", dtPrv.Rows[0]["MOBLNO"], " ▶ ", sMoblCd,
                //        ", 사용여부 : ", dtPrv.Rows[0]["USEYN"], " ▶ ", sUseYn,
                //        ", 검수자여부 : ", dtPrv.Rows[0]["ISPT_YN"], " ▶ ", sIsptYn);
                //}
                //else
                //{
                //    sLOG = string.Format("[사용자관리]" +
                //           ", 사용자ID : {0}" +
                //           ", 패스워드 : {1}" +
                //           ", 유저명 : {2}" +
                //           ", 부서코드 : {3}" +
                //           ", 직위코드 : {4}" +
                //           //#001 인사코드추가
                //           ", 인사코드 : {5}" +
                //           ", 연락처 : {6}" +
                //           ", 사용여부 : {7}" +
                //           ", 검수자여부 : {8}"
                //           , sUserID
                //           , sPassWD
                //           , sUserNm
                //           , sDeptCd
                //           , sJobGrade
                //           //#001 인사코드추가
                //           , sInsaNo
                //           , sMoblCd
                //           , sUseYn
                //           , sIsptYn);
                //}
                #endregion

                strSql.Clear();
                strSql.AppendLine("");
                strSql.AppendLine(" IF EXISTS ( SELECT * FROM ZUSRLST WHERE USRCD = "+ dUserCD.ToString() + " )");
                strSql.AppendLine("    BEGIN ");
                strSql.AppendLine("          UPDATE ZUSRLST ");
                strSql.AppendLine("             SET USRID   = '"+ sUserID+"'");
                strSql.AppendLine("               , USRNM   = '"+ sUserNm + "'   ");
                strSql.AppendLine("               , PASSWD  = '"+ sPassWD + "'  ");
                strSql.AppendLine("               , DEPTCD  = '"+ sDeptCd + "'  ");
                strSql.AppendLine("               , JKWICD  = '"+ sJobGrade + "' ");
                strSql.AppendLine("               , INSANO  = '"+ sInsaNo + "'  ");
                strSql.AppendLine("               , MOBLNO  = '"+ sMoblCd + "'  ");
                strSql.AppendLine("               , USEYN   = '"+ sUseYn + "'   ");
                strSql.AppendLine("               , USLVL   = '"+ sGrade + "'   ");
                strSql.AppendLine("               , RK      = '"+ sRemark + "'      ");
                strSql.AppendLine("               , ISPT_YN = '"+ sIsptYn + "' ");
                strSql.AppendLine("               , MDATE   = GETDATE()   ");
                strSql.AppendLine("               , MUSER   = '"+ FmMainToolBar2.UserID + "'   ");
                strSql.AppendLine("           WHERE USRCD = "+ dUserCD.ToString() + " ");
                strSql.AppendLine("      END ");
                strSql.AppendLine(" ELSE ");
                strSql.AppendLine("    BEGIN ");
                strSql.AppendLine("           INSERT INTO ZUSRLST ");
                strSql.AppendLine("                     ( USRCD , USRID  , USRNM  ");
                strSql.AppendLine("                     , PASSWD, DEPTCD , JKWICD  ");
                strSql.AppendLine("                     , MOBLNO, INSANO , USEYN  ");
                strSql.AppendLine("                     , USLVL , RK     , ISPT_YN ");
                strSql.AppendLine("                     , CDATE , CUSER ) ");
                strSql.AppendLine("               VALUES( "+ dUserCD.ToString() + " , '"+ sUserID + "'  , '"+sUserNm+"'  ");
                strSql.AppendLine("                     , '"+ sPassWD + "', '"+ sDeptCd + "' , '"+ sJobGrade + "'  ");
                strSql.AppendLine("                     , '"+ sMoblCd + "', '"+ sInsaNo + "', '"+ sUseYn + "'  ");
                strSql.AppendLine("                     , '"+ sGrade + "' , '"+ sRemark + "'    , '"+ sIsptYn + "'");
                strSql.AppendLine("                     , GETDATE() , '"+ FmMainToolBar2.UserID + "' ); ");
                strSql.AppendLine(" ");
                strSql.AppendLine("            ");
                strSql.AppendLine("      END ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT PGMID ");
                strSql.AppendLine("      , PGMNM ");
                strSql.AppendLine("      , SYSGU ");
                strSql.AppendLine("      , PGGRP ");
                strSql.AppendLine("      , PGGUB ");
                strSql.AppendLine("      , PGTAG ");
                strSql.AppendLine("      , PGLVL ");
                strSql.AppendLine("      , USEYN ");
                strSql.AppendLine("      , RK ");
                strSql.AppendLine("      , CDATE ");
                strSql.AppendLine("      , CUSER ");
                strSql.AppendLine("      , MDATE ");
                strSql.AppendLine("      , MUSER ");
                strSql.AppendLine("   FROM ZPGMLST ");
                strSql.AppendLine("  WHERE USEYN = 'Y' ");

                DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                //#002
                string sUslvl = string.Empty;

                if (drUserInfo != null)
                    sUslvl = drUserInfo["USLVL"]?.ToString();

                if (!sGrade.Equals(sUslvl))
                {
                    for (int i = 0; i < dtChk.Rows.Count; i++)
                    {
                        string sDtGrade = dtChk.Rows[i]["PGLVL"]?.ToString();
                        double dPgmGrade = !string.IsNullOrEmpty(sDtGrade) ? Convert.ToDouble(sDtGrade) : 0;
                        double dUserGrade = Convert.ToDouble(sGrade);

                        #region mariaDB
                        //strSql.Clear();
                        //strSql.AppendLine(" ");
                        //strSql.AppendLine(" INSERT INTO ZPGMAUT");
                        //strSql.AppendLine("           ( USRCD ");
                        //strSql.AppendLine("           , PGMID ");
                        //strSql.AppendLine("           , USE_Y ");
                        //strSql.AppendLine("           , ADD_Y ");
                        //strSql.AppendLine("           , UPD_Y ");
                        //strSql.AppendLine("           , DEL_Y ");
                        //strSql.AppendLine("           , PRT_Y ");
                        //strSql.AppendLine("           , XLS_Y ");
                        //strSql.AppendLine("           , CDATE ");
                        //strSql.AppendLine("           , CUSER ");
                        //strSql.AppendLine("           , MDATE ");
                        //strSql.AppendLine("           , MUSER) ");
                        //strSql.AppendLine("     VALUES( " + dUserCD + " ");
                        //strSql.AppendLine("           , '" + dtChk.Rows[i]["PGMID"].ToString() + "' ");
                        //if (dUserGrade == 0)
                        //{
                        //    strSql.AppendLine("           , 'Y' ");
                        //    strSql.AppendLine("           , 'Y' ");
                        //    strSql.AppendLine("           , 'Y' ");
                        //    strSql.AppendLine("           , 'Y' ");
                        //    strSql.AppendLine("           , 'Y' ");
                        //    strSql.AppendLine("           , 'Y' ");
                        //}
                        //else if (dPgmGrade >= dUserGrade)
                        //{
                        //    strSql.AppendLine("           , 'Y' ");
                        //    strSql.AppendLine("           , 'Y' ");
                        //    strSql.AppendLine("           , 'N' ");
                        //    strSql.AppendLine("           , 'Y' ");
                        //    strSql.AppendLine("           , 'Y' ");
                        //    strSql.AppendLine("           , 'Y' ");
                        //}
                        //else if (dPgmGrade <= dUserGrade)
                        //{
                        //    strSql.AppendLine("           , 'N' ");
                        //    strSql.AppendLine("           , 'N' ");
                        //    strSql.AppendLine("           , 'N' ");
                        //    strSql.AppendLine("           , 'N' ");
                        //    strSql.AppendLine("           , 'N' ");
                        //    strSql.AppendLine("           , 'N' ");
                        //}
                        //strSql.AppendLine("           , NOW() ");
                        //strSql.AppendLine("           , '" + sEntId + "' ");
                        //strSql.AppendLine("           , NOW() ");
                        //strSql.AppendLine("           , '" + sMfyId + "' ");
                        //strSql.AppendLine("           ) ");
                        //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                        //strSql.AppendLine("             PGMID = '" + dtChk.Rows[i]["PGMID"].ToString() + "' ");
                        //if (dUserGrade == 0)
                        //{
                        //    strSql.AppendLine("           , USE_Y = 'Y' ");
                        //    strSql.AppendLine("           , ADD_Y = 'Y' ");
                        //    strSql.AppendLine("           , UPD_Y = 'Y' ");
                        //    strSql.AppendLine("           , DEL_Y = 'Y' ");
                        //    strSql.AppendLine("           , PRT_Y = 'Y' ");
                        //    strSql.AppendLine("           , XLS_Y = 'Y' ");
                        //}
                        //else if (dPgmGrade >= dUserGrade)
                        //{
                        //    strSql.AppendLine("           , USE_Y = 'Y' ");
                        //    strSql.AppendLine("           , ADD_Y = 'Y' ");
                        //    strSql.AppendLine("           , UPD_Y = 'N' ");
                        //    strSql.AppendLine("           , DEL_Y = 'Y' ");
                        //    strSql.AppendLine("           , PRT_Y = 'Y' ");
                        //    strSql.AppendLine("           , XLS_Y = 'Y' ");
                        //}
                        //else if (dPgmGrade <= dUserGrade)
                        //{
                        //    strSql.AppendLine("           , USE_Y = 'N' ");
                        //    strSql.AppendLine("           , ADD_Y = 'N' ");
                        //    strSql.AppendLine("           , UPD_Y = 'N' ");
                        //    strSql.AppendLine("           , DEL_Y = 'N' ");
                        //    strSql.AppendLine("           , PRT_Y = 'N' ");
                        //    strSql.AppendLine("           , XLS_Y = 'N' ");
                        //}
                        //strSql.AppendLine("           , MDATE = NOW() ");
                        //strSql.AppendLine("           , MUSER = '" + sMfyId + "' ");
                        #endregion

                        strSql.Clear();
                        strSql.AppendLine("IF EXISTS(SELECT* FROM ZPGMAUT WHERE USRCD = "+dUserCD+" AND PGMID = '"+dtChk.Rows[i]["PGMID"].ToString()+"')");
                        strSql.AppendLine("   BEGIN                                                                                                     ");
                        strSql.AppendLine("         UPDATE ZPGMAUT                                                                                      ");
                        if (dUserGrade == 0)
                        {
                            strSql.AppendLine("         SET USE_Y = 'Y' ");
                            strSql.AppendLine("           , ADD_Y = 'Y' ");
                            strSql.AppendLine("           , UPD_Y = 'Y' ");
                            strSql.AppendLine("           , DEL_Y = 'Y' ");
                            strSql.AppendLine("           , PRT_Y = 'Y' ");
                            strSql.AppendLine("           , XLS_Y = 'Y' ");
                        }
                        else if (dPgmGrade >= dUserGrade)
                        {
                            strSql.AppendLine("         SET USE_Y = 'Y' ");
                            strSql.AppendLine("           , ADD_Y = 'Y' ");
                            strSql.AppendLine("           , UPD_Y = 'N' ");
                            strSql.AppendLine("           , DEL_Y = 'Y' ");
                            strSql.AppendLine("           , PRT_Y = 'Y' ");
                            strSql.AppendLine("           , XLS_Y = 'Y' ");
                        }
                        else if (dPgmGrade <= dUserGrade)
                        {
                            strSql.AppendLine("         SET USE_Y = 'N' ");
                            strSql.AppendLine("           , ADD_Y = 'N' ");
                            strSql.AppendLine("           , UPD_Y = 'N' ");
                            strSql.AppendLine("           , DEL_Y = 'N' ");
                            strSql.AppendLine("           , PRT_Y = 'N' ");
                            strSql.AppendLine("           , XLS_Y = 'N' ");
                        }
                        strSql.AppendLine("                 , MDATE = GETDATE()                               ");
                        strSql.AppendLine("                 , MUSER = '0'                                     ");
                        strSql.AppendLine("             WHERE USRCD = "+dUserCD+"                             ");
                        strSql.AppendLine("               AND PGMID = '"+dtChk.Rows[i]["PGMID"].ToString()+"' ");
                        strSql.AppendLine("      END                                                          ");
                        strSql.AppendLine("   ELSE                                                            ");
                        strSql.AppendLine("      BEGIN                                                        ");
                        strSql.AppendLine("                INSERT INTO ZPGMAUT                                ");
                        strSql.AppendLine("                  (USRCD                                           ");
                        strSql.AppendLine("                  , PGMID                                          ");
                        strSql.AppendLine("                  , USE_Y                                          ");
                        strSql.AppendLine("                  , ADD_Y                                          ");
                        strSql.AppendLine("                  , UPD_Y                                          ");
                        strSql.AppendLine("                  , DEL_Y                                          ");
                        strSql.AppendLine("                  , PRT_Y                                          ");
                        strSql.AppendLine("                  , XLS_Y                                          ");
                        strSql.AppendLine("                  , CDATE                                          ");
                        strSql.AppendLine("                  , CUSER)                                         ");
                        strSql.AppendLine("            VALUES("+dUserCD+"                                     ");
                        strSql.AppendLine("                  , '" + dtChk.Rows[i]["PGMID"].ToString() + "'    ");
                        if (dUserGrade == 0)
                        {
                            strSql.AppendLine("           , 'Y' ");
                            strSql.AppendLine("           , 'Y' ");
                            strSql.AppendLine("           , 'Y' ");
                            strSql.AppendLine("           , 'Y' ");
                            strSql.AppendLine("           , 'Y' ");
                            strSql.AppendLine("           , 'Y' ");
                        }
                        else if (dPgmGrade >= dUserGrade)
                        {
                            strSql.AppendLine("           , 'Y' ");
                            strSql.AppendLine("           , 'Y' ");
                            strSql.AppendLine("           , 'N' ");
                            strSql.AppendLine("           , 'Y' ");
                            strSql.AppendLine("           , 'Y' ");
                            strSql.AppendLine("           , 'Y' ");
                        }
                        else if (dPgmGrade <= dUserGrade)
                        {
                            strSql.AppendLine("           , 'N' ");
                            strSql.AppendLine("           , 'N' ");
                            strSql.AppendLine("           , 'N' ");
                            strSql.AppendLine("           , 'N' ");
                            strSql.AppendLine("           , 'N' ");
                            strSql.AppendLine("           , 'N' ");
                        }
                        strSql.AppendLine("               , GETDATE()");
                        strSql.AppendLine("               , '0'      ");
                        strSql.AppendLine("               )          ");
                        strSql.AppendLine("   END                    ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                Cursor = Cursors.Default;

                XtraMessageBox.Show("저장을 완료했습니다.");

                UserMgt fm = this.Owner as UserMgt;
                fm._USRID = sUserID;
                DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
                Cursor = Cursors.Default;
            }

            #region[2020-12-29 로그 관련 로직 변경이전]

            //string sUserID = TxtUserId.EditValue?.ToString();
            //string sPassWD = TxtPw.EditValue.ToString();
            //if (string.IsNullOrEmpty(sUserID) || string.IsNullOrEmpty(sPassWD))
            //{
            //    XtraMessageBox.Show("사용자ID와 PW는 필수항목입니다");
            //    return;
            //}

            //string sGrade = CboGrade.EditValue?.ToString();
            //if (string.IsNullOrEmpty(sGrade))
            //{
            //    XtraMessageBox.Show("권한 등급을 선택해주세요.");
            //    return;
            //}
            
            //try
            //{
            //    double dUserCD = 0;

            //    string sUserNm = LkupEmpNm.Text;
            //    string sInsaNo = LkupEmpNm.EditValue?.ToString();
            //    string sDeptCd = LkupDept.EditValue?.ToString();
            //    string sJobGrade = LkupJobGrade.EditValue?.ToString();
            //    string sMoblCd = TxtCallNo.EditValue?.ToString();
            //    string sUseYn = ChkUse.EditValue?.ToString();
            //    string sRemark = TxtMemo.EditValue?.ToString();
            //    string sIsptYn = ChkIsptYn.EditValue?.ToString();
            //    string sEntId = FmMainToolBar2.UserID;
            //    string sMfyId = FmMainToolBar2.UserID;

            //    Cursor = Cursors.WaitCursor;


            //    StringBuilder strSql = new StringBuilder();

            //    if (drUserInfo != null)
            //    {
            //        dUserCD = Convert.ToDouble(drUserInfo["USRCD"]);
            //    }
            //    else
            //    {
            //        strSql.Clear();
            //        strSql.AppendLine(" ");
            //        strSql.AppendLine(" SELECT MAX(A.USRCD) AS MAX_VALUE ");
            //        strSql.AppendLine("   FROM ZUSRLST A");

            //        DataTable dtMax = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            //        if (dtMax.Rows.Count == 0)
            //        {
            //            dUserCD = 1000;
            //        }
            //        else if (dtMax.Rows.Count > 0)
            //        {
            //            dUserCD = Convert.ToDouble(dtMax.Rows[0]["MAX_VALUE"]) + 100;
            //        }
            //    }

            //    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            //    MySqlCommand cmd = DBConn.dbCon.CreateCommand();
            //    cmd.Transaction = DBConn.dbTran;

            //    string sLOG = string.Empty;
            //    string sEDIT_KIND = string.Empty;

            //    strSql.Clear();
            //    strSql.AppendLine(" ");
            //    strSql.AppendLine(" SELECT COUNT(*) AS CNT ");
            //    strSql.AppendLine("   FROM ZUSRLST ");
            //    strSql.AppendLine("  WHERE USRCD = @USRCD ");

            //    cmd.CommandType = CommandType.Text;
            //    cmd.CommandText = strSql.ToString();
            //    cmd.Parameters.AddWithValue("@USRCD", dUserCD);
            //    object cnt = cmd.ExecuteScalar();
            //    cmd.Parameters.Clear();
            //    int iCnt = 0;
            //    int.TryParse(cnt?.ToString(), out iCnt);
            //    if (iCnt == 0)
            //    {
            //        sEDIT_KIND = "I";
            //    }
            //    else
            //    {
            //        sEDIT_KIND = "U";
            //    }

            //    sLOG = string.Format("[사용자관리]" +
            //            "사용자ID : {0}" +
            //            "패스워드 : {1}" +
            //            "유저명 : {2}" +
            //            "부서코드 : {3}" +
            //            "직위코드 : {4}" +
            //            "연락처 : {5}" +
            //            "사용여부 : {6}" +
            //            "검수자여부 : {7}"
            //            , sUserID
            //            , sPassWD
            //            , sUserNm
            //            , sDeptCd
            //            , sJobGrade
            //            , sMoblCd
            //            , sUseYn
            //            , sIsptYn);

            //    strSql.Clear();
            //    strSql.AppendLine("");
            //    strSql.AppendLine(" INSERT INTO ZUSRLST");
            //    strSql.AppendLine("           ( ");
            //    strSql.AppendLine("             USRCD ");
            //    strSql.AppendLine("           , USRID ");
            //    strSql.AppendLine("           , USRNM ");
            //    strSql.AppendLine("           , PASSWD ");
            //    strSql.AppendLine("           , DEPTCD ");
            //    strSql.AppendLine("           , JKWICD ");
            //    strSql.AppendLine("           , MOBLNO ");
            //    strSql.AppendLine("           , INSANO ");
            //    strSql.AppendLine("           , USEYN ");
            //    strSql.AppendLine("           , USLVL ");
            //    strSql.AppendLine("           , RK ");
            //    strSql.AppendLine("           , ISPT_YN ");
            //    strSql.AppendLine("           , CDATE ");
            //    strSql.AppendLine("           , CUSER ");
            //    strSql.AppendLine("           ) ");
            //    strSql.AppendLine("     VALUES( ");
            //    strSql.AppendLine("             " + dUserCD + " ");
            //    strSql.AppendLine("           , '" + sUserID + "' ");
            //    strSql.AppendLine("           , '" + sUserNm + "' ");
            //    strSql.AppendLine("           , '" + sPassWD + "' ");
            //    strSql.AppendLine("           , '" + sDeptCd + "' ");
            //    strSql.AppendLine("           , '" + sJobGrade + "' ");
            //    strSql.AppendLine("           , '" + sMoblCd + "' ");
            //    strSql.AppendLine("           , '" + sInsaNo + "' ");
            //    strSql.AppendLine("           , '" + sUseYn + "' ");
            //    strSql.AppendLine("           , '" + sGrade + "' ");
            //    strSql.AppendLine("           , '" + sRemark + "' ");
            //    strSql.AppendLine("           , '" + sIsptYn + "' ");
            //    strSql.AppendLine("           , NOW() ");
            //    strSql.AppendLine("           , '" + sEntId + "' ");
            //    strSql.AppendLine("           ) ");
            //    strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
            //    strSql.AppendLine("             USRID = '" + sUserID + "' ");
            //    strSql.AppendLine("           , USRNM = '" + sUserNm + "' ");
            //    strSql.AppendLine("           , PASSWD = '" + sPassWD + "' ");
            //    strSql.AppendLine("           , DEPTCD = '" + sDeptCd + "' ");
            //    strSql.AppendLine("           , JKWICD = '" + sJobGrade + "' ");
            //    strSql.AppendLine("           , MOBLNO = '" + sMoblCd + "' ");
            //    strSql.AppendLine("           , USEYN = '" + sUseYn + "' ");
            //    strSql.AppendLine("           , USLVL = '" + sGrade + "' ");
            //    strSql.AppendLine("           , RK = '" + sRemark + "' ");
            //    strSql.AppendLine("           , ISPT_YN = '" + sIsptYn + "' ");
            //    strSql.AppendLine("           , MDATE = NOW() ");
            //    strSql.AppendLine("           , MUSER = '" + sMfyId + "' ");

            //    cmd.CommandType = CommandType.Text;
            //    cmd.CommandText = strSql.ToString();
            //    cmd.ExecuteNonQuery();

            //    strSql.Clear();
            //    strSql.AppendLine(" ");
            //    strSql.AppendLine(" SELECT PGMID ");
            //    strSql.AppendLine("      , PGMNM ");
            //    strSql.AppendLine("      , SYSGU ");
            //    strSql.AppendLine("      , PGGRP ");
            //    strSql.AppendLine("      , PGGUB ");
            //    strSql.AppendLine("      , PGTAG ");
            //    strSql.AppendLine("      , PGLVL ");
            //    strSql.AppendLine("      , USEYN ");
            //    strSql.AppendLine("      , RK ");
            //    strSql.AppendLine("      , CDATE ");
            //    strSql.AppendLine("      , CUSER ");
            //    strSql.AppendLine("      , MDATE ");
            //    strSql.AppendLine("      , MUSER ");
            //    strSql.AppendLine("   FROM ZPGMLST ");
            //    strSql.AppendLine("  WHERE USEYN = 'Y' ");

            //    DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            //    ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, "1", sEDIT_KIND, "UserMgt", sLOG, cmd);

            //    for (int i = 0; i < dtChk.Rows.Count; i++)
            //    {
            //        string sDtGrade = dtChk.Rows[i]["PGLVL"]?.ToString();
            //        double dPgmGrade = !string.IsNullOrEmpty(sDtGrade) ? Convert.ToDouble(sDtGrade) : 0;
            //        double dUserGrade = Convert.ToDouble(sGrade);

            //        strSql.Clear();
            //        strSql.AppendLine(" ");
            //        strSql.AppendLine(" INSERT INTO ZPGMAUT");
            //        strSql.AppendLine("           ( USRCD ");
            //        strSql.AppendLine("           , PGMID ");
            //        strSql.AppendLine("           , USE_Y ");
            //        strSql.AppendLine("           , ADD_Y ");
            //        strSql.AppendLine("           , UPD_Y ");
            //        strSql.AppendLine("           , DEL_Y ");
            //        strSql.AppendLine("           , PRT_Y ");
            //        strSql.AppendLine("           , XLS_Y ");
            //        strSql.AppendLine("           , CDATE ");
            //        strSql.AppendLine("           , CUSER ");
            //        strSql.AppendLine("           , MDATE ");
            //        strSql.AppendLine("           , MUSER) ");
            //        strSql.AppendLine("     VALUES( " + dUserCD + " ");
            //        strSql.AppendLine("           , '" + dtChk.Rows[i]["PGMID"].ToString() + "' ");
            //        if(dUserGrade == 0)
            //        {
            //            strSql.AppendLine("           , 'Y' ");
            //            strSql.AppendLine("           , 'Y' ");
            //            strSql.AppendLine("           , 'Y' ");
            //            strSql.AppendLine("           , 'Y' ");
            //            strSql.AppendLine("           , 'Y' ");
            //            strSql.AppendLine("           , 'Y' ");
            //        }
            //        else if (dPgmGrade >= dUserGrade)
            //        {
            //            strSql.AppendLine("           , 'Y' ");
            //            strSql.AppendLine("           , 'Y' ");
            //            strSql.AppendLine("           , 'N' ");
            //            strSql.AppendLine("           , 'Y' ");
            //            strSql.AppendLine("           , 'Y' ");
            //            strSql.AppendLine("           , 'Y' ");
            //        }
            //        else if(dPgmGrade <= dUserGrade)
            //        {
            //            strSql.AppendLine("           , 'N' ");
            //            strSql.AppendLine("           , 'N' ");
            //            strSql.AppendLine("           , 'N' ");
            //            strSql.AppendLine("           , 'N' ");
            //            strSql.AppendLine("           , 'N' ");
            //            strSql.AppendLine("           , 'N' ");
            //        }
            //        strSql.AppendLine("           , NOW() ");
            //        strSql.AppendLine("           , '" + sEntId + "' ");
            //        strSql.AppendLine("           , NOW() ");
            //        strSql.AppendLine("           , '" + sMfyId + "' ");
            //        strSql.AppendLine("           ) ");
            //        strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
            //        strSql.AppendLine("             PGMID = '" + dtChk.Rows[i]["PGMID"].ToString() + "' ");
            //        if (dUserGrade == 0)
            //        {
            //            strSql.AppendLine("           , USE_Y = 'Y' ");
            //            strSql.AppendLine("           , ADD_Y = 'Y' ");
            //            strSql.AppendLine("           , UPD_Y = 'Y' ");
            //            strSql.AppendLine("           , DEL_Y = 'Y' ");
            //            strSql.AppendLine("           , PRT_Y = 'Y' ");
            //            strSql.AppendLine("           , XLS_Y = 'Y' ");
            //        }
            //        else if (dPgmGrade >= dUserGrade)
            //        {
            //            strSql.AppendLine("           , USE_Y = 'Y' ");
            //            strSql.AppendLine("           , ADD_Y = 'Y' ");
            //            strSql.AppendLine("           , UPD_Y = 'N' ");
            //            strSql.AppendLine("           , DEL_Y = 'Y' ");
            //            strSql.AppendLine("           , PRT_Y = 'Y' ");
            //            strSql.AppendLine("           , XLS_Y = 'Y' ");
            //        }
            //        else if (dPgmGrade <= dUserGrade)
            //        {
            //            strSql.AppendLine("           , USE_Y = 'N' ");
            //            strSql.AppendLine("           , ADD_Y = 'N' ");
            //            strSql.AppendLine("           , UPD_Y = 'N' ");
            //            strSql.AppendLine("           , DEL_Y = 'N' ");
            //            strSql.AppendLine("           , PRT_Y = 'N' ");
            //            strSql.AppendLine("           , XLS_Y = 'N' ");
            //        }
            //        strSql.AppendLine("           , MDATE = NOW() ");
            //        strSql.AppendLine("           , MUSER = '" + sMfyId + "' ");

            //        cmd.CommandType = CommandType.Text;
            //        cmd.CommandText = strSql.ToString();
            //        cmd.ExecuteNonQuery();
            //    }
                
            //    DBConn.dbTran.Commit();
            //    DBConn.dbTran = null;

            //    Cursor = Cursors.Default;

            //    XtraMessageBox.Show("저장을 완료했습니다.");
            //}
            //catch (Exception ex)
            //{
            //    DBConn.dbTran.Rollback();
            //    DBConn.dbTran = null;
            //    MessageBox.Show(ex.Message);
            //}
            //BtnClear_Click(null, null);
            //this.DialogResult = DialogResult.OK;

            //Dispose();

            #endregion[2020-12-29 로그관련 로직변경 이전 소스]
        }

        private void PopUpUserInsert_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnClear_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {

            }
            else if (e.KeyCode == Keys.F5)
            {
               
            }
            else if (e.KeyCode == Keys.F8)
            {
                
            }
        }

        //private void LkupEmpNm_KeyDown(object sender, KeyEventArgs e)
        //{
        //    if(e.KeyCode == Keys.Enter)
        //    {
        //        string sEmpId = LkupEmpNm.EditValue?.ToString();

        //        if (!string.IsNullOrEmpty(sEmpId))
        //        {
        //            StringBuilder strSql = new StringBuilder();

        //            strSql.Clear();
        //            strSql.AppendLine(" ");
        //            strSql.AppendLine(" SELECT A.DEPT_CD ");
        //            strSql.AppendLine(" 	 , A.GRADE_CD ");
        //            strSql.AppendLine(" 	 , B.HP_NO ");
        //            strSql.AppendLine("   FROM HR_EMP_BASIS A  ");
        //            strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_PERSONAL B ");
        //            strSql.AppendLine("     ON B.EMP_ID = A.EMP_ID ");
        //            strSql.AppendLine("  WHERE A.EMP_ID = '" + sEmpId + "' ");

        //            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

        //            if(dt.Rows.Count > 0)
        //            {
        //                string sDeptCd = dt.Rows[0]["DEPT_CD"]?.ToString();
        //                string sGradeCd = dt.Rows[0]["GRADE_CD"]?.ToString();
        //                string sHPNo = dt.Rows[0]["HP_NO"]?.ToString();

        //                LkupDept.EditValue = sDeptCd;
        //                LkupJobGrade.EditValue = sGradeCd;
        //                TxtCallNo.EditValue = sHPNo;
        //            }
        //        }
        //    }
        //}

        //#001
        private void LkupEmpNm_Leave(object sender, EventArgs e)
        {
            string sEmpId = LkupEmpNm.EditValue?.ToString();

            if (!string.IsNullOrEmpty(sEmpId))
            {
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT A.DEPT_CD ");
                strSql.AppendLine(" 	 , A.GRADE_CD ");
                strSql.AppendLine(" 	 , B.HP_NO ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A  ");
                strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_PERSONAL B ");
                strSql.AppendLine("     ON B.EMP_ID = A.EMP_ID ");
                strSql.AppendLine("  WHERE A.EMP_ID = '" + sEmpId + "' ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt.Rows.Count > 0)
                {
                    string sDeptCd = dt.Rows[0]["DEPT_CD"]?.ToString();
                    string sGradeCd = dt.Rows[0]["GRADE_CD"]?.ToString();
                    string sHPNo = dt.Rows[0]["HP_NO"]?.ToString();

                    LkupDept.EditValue = sDeptCd;
                    LkupJobGrade.EditValue = sGradeCd;
                    TxtCallNo.EditValue = sHPNo;
                }
            }
        }

        private void CboGrade_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                BtnSave.Focus();
            }
        }
    }
}