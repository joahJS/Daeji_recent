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

namespace AccAdm
{
    public partial class PD01001F01 : DevExpress.XtraEditors.XtraForm
    {
        public PD01001F01()
        {
            InitializeComponent();
        }

        public string USRCD;
        public ProdPlanMgt PARENT_FORM;
        private void PD01001F01_Load(object sender, EventArgs e)
        {
            DateEditYmd.EditValue = DateTime.Now;
        }
        
        private void BtnCheck_Click(object sender, EventArgs e)
        {
            string sYmd = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            if (string.IsNullOrEmpty(sYmd))
            {
                XtraMessageBox.Show("기준일자를 선택하세요.");
                return;
            }

            /*
             * 해당 일자에 해당 유저의 생산정보 개수 체크
             * 1 이상일 경우 결재여부 체크 
             * 결재가 아닐 경우 수정할 수 있게 A리턴
             * 결재일 경우 수정할 수 없게 B 리턴
             * 아무 값도 없을경우 등록할 수 있게 C 리턴
             */
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT COUNT(A.CNT) AS CNT ");
            strSql.AppendLine("   FROM ( SELECT 1 AS CNT ");
            strSql.AppendLine("            FROM MAKE_M X1 ");
            strSql.AppendLine("           WHERE X1.MDATE = '" + sYmd + "' ");
            strSql.AppendLine("             AND X1.MUSER_ID = '" + USRCD + "' ) A  ");

            //strSql.AppendLine(" SELECT COUNT(*) AS CNT ");
            //strSql.AppendLine("   FROM MAKE_M A  ");
            //strSql.AppendLine("  WHERE A.MDATE = '" + sYmd + "' ");
            //strSql.AppendLine("    AND A.MUSER_ID = '" + USRCD + "' ");

            DataTable dt_Chk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            int iCnt = Convert.ToInt32(dt_Chk.Rows[0]["CNT"].ToString());
            if(iCnt > 0) //생산정보 존재하므로 결재부분 체크
            {
                strSql.Clear();
                strSql.AppendLine(" SELECT CASE WHEN SIGN1 = 'Y' THEN 'Y' ELSE 'N' END AS YN ");
                strSql.AppendLine("   FROM MAKE_S ");
                strSql.AppendLine("  WHERE MDATE = '" + sYmd + "' ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                if(dt.Rows.Count > 0)
                {
                    string sYn = dt.Rows[0]["YN"]?.ToString();

                    strSql.Clear();
                    strSql.AppendLine(" SELECT A.MAKENO ");
                    strSql.AppendLine("   FROM MAKE_M A ");
                    strSql.AppendLine("  WHERE MDATE = '" + sYmd + "' ");
                    strSql.AppendLine("    AND MUSER_ID = '" + USRCD + "' ");

                    dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    double dMakeNo = Convert.ToDouble(dt.Rows[0]["MAKENO"]);

                    

                    if (sYn.Equals("Y")) //결재의 경우 B리턴
                    {
                        XtraMessageBox.Show(string.Format("{0}-{1}-{2}의 생산 건은 결재승인 상태입니다.", sYmd.Substring(0, 4), sYmd.Substring(4, 2), sYmd.Substring(6, 2)));
                        PARENT_FORM.RESULT_GB = "B";
                        PARENT_FORM.RESULT_MAKENO = dMakeNo;
                        PARENT_FORM.RESULT_YMD = sYmd;
                    }
                    else //미결재일 경우 A리턴
                    {
                        PARENT_FORM.RESULT_GB = "A";
                        PARENT_FORM.RESULT_MAKENO = dMakeNo;
                        PARENT_FORM.RESULT_YMD = sYmd;
                    }
                    DialogResult = DialogResult.OK;
                }
                else //존재하지 않으므로 미결재판단 A리턴
                {
                    strSql.Clear();
                    strSql.AppendLine(" SELECT A.MAKENO ");
                    strSql.AppendLine("   FROM MAKE_M A ");
                    strSql.AppendLine("  WHERE MDATE = '" + sYmd + "' ");
                    strSql.AppendLine("    AND MUSER_ID = '" + USRCD + "' ");

                    dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    double dMakeNo = Convert.ToDouble(dt.Rows[0]["MAKENO"]);

                    PARENT_FORM.RESULT_GB = "A";
                    PARENT_FORM.RESULT_MAKENO = dMakeNo;
                    PARENT_FORM.RESULT_YMD = sYmd;
                    DialogResult = DialogResult.OK;
                }
            }
            else //생산정보 미존재 판단, 결재여부 체크
            {
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT A.REAL_DUTY_DEPT ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A ");
                strSql.AppendLine("  WHERE EMP_ID = '" + USRCD + "' ");
                strSql.AppendLine("    AND EMPL_GB = 'Y' ");
                DataTable dtDept = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                string sDept = dtDept.Rows[0]["REAL_DUTY_DEPT"]?.ToString();

                //결재여부 체크
                strSql.Clear();
                if (sDept == "4100") { strSql.AppendLine(" SELECT CASE WHEN SIGN1 = 'Y' THEN 'Y' ELSE 'N' END AS YN "); }
                if (sDept == "4150") { strSql.AppendLine(" SELECT CASE WHEN SIGN1a = 'Y' THEN 'Y' ELSE 'N' END AS YN "); }
                if (sDept == "4200") { strSql.AppendLine(" SELECT CASE WHEN SIGN2 = 'Y' THEN 'Y' ELSE 'N' END AS YN "); }
                if (sDept == "4300") { strSql.AppendLine(" SELECT CASE WHEN SIGN3 = 'Y' THEN 'Y' ELSE 'N' END AS YN "); }
                if (sDept == "2000") { strSql.AppendLine(" SELECT CASE WHEN SIGN4 = 'Y' THEN 'Y' ELSE 'N' END AS YN "); }
                strSql.AppendLine("   FROM MAKE_S ");
                strSql.AppendLine("  WHERE MDATE = '" + sYmd + "' ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt.Rows.Count > 0)
                {
                    string sYn = dt.Rows[0]["YN"]?.ToString();
                    if (sYn.Equals("Y")) //결재완료 시 등록불가
                    {
                        XtraMessageBox.Show(string.Format("{0}-{1}-{2}의 생산 건은 결재승인 상태이므로 등록할 수 없습니다.\r\n기준일자를 다시 설정하세요.", sYmd.Substring(0, 4), sYmd.Substring(4, 2), sYmd.Substring(6, 2)));
                        DateEditYmd.SelectAll();
                        DateEditYmd.Focus();
                        return;
                    }
                    else //미결재 시 등록
                    {
                        PARENT_FORM.RESULT_GB = "C";
                        PARENT_FORM.RESULT_YMD = sYmd;
                        DialogResult = DialogResult.OK;
                    }
                }
                else //존재하지 않으므로 미결재 판단 등록
                {
                    PARENT_FORM.RESULT_GB = "C";
                    PARENT_FORM.RESULT_YMD = sYmd;
                    DialogResult = DialogResult.OK;
                }
            }

        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void PD01001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
                BtnCheck_Click(null, null);
            else if (e.KeyCode == Keys.Escape)
                BtnCancel_Click(null, null);
        }
        
        private void DateEditYmd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DateEditYmd.DoValidate();
                BtnCheck_Click(null, null);
            }
        }
        
    }
}