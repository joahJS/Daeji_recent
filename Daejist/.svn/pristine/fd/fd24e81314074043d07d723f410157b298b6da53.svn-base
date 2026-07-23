using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using System.Data;
using System.Text;
using ComLib;

namespace AccAdm
{
    public partial class RptCertOfEmp : DevExpress.XtraReports.UI.XtraReport
    {
        //public DataRow DrCertInfo;
        public RptCertOfEmp()
        {
            InitializeComponent();
            XRLblYearAndSeq.DataBindings.Add("Text", null, "ISSUE_CONTENT");
            XRLblDept.DataBindings.Add("Text", null, "DEPT_NM");
            XRLblName.DataBindings.Add("Text", null, "EMP_NM");
            XRLblIdtNo.DataBindings.Add("Text", null, "IDT_NO");
            XRLblAddr.DataBindings.Add("Text", null, "ADDR");
            XRLblEntranceYmd.DataBindings.Add("Text", null, "ENTRANCE_YMD");
            XRLblGrade.DataBindings.Add("Text", null, "GRADE_NM");
            XRLblPurp.DataBindings.Add("Text", null, "PURP");
            XRLblIssueYmd.DataBindings.Add("Text", null, "ISSUE_YMD");

            //SetYearAndSeq();
            //SetDeptName();
            //SetEmpName();
            //SetIdentificationNumber();
            //SetAddress();
            //SetEntranceYmd();
            //SetGrade();
            //SetPurpose();
            //SetIssueYmd();
        }

        //private void SetYearAndSeq()
        //{
        //    string sYear = DrCertInfo["ISSUE_YEAR"]?.ToString();
        //    string sSeq = DrCertInfo["ISSUE_SEQ"]?.ToString();

        //    string sResult = "제 " + sYear + "-" + sSeq + "호";

        //    XRLblYearAndSeq.Text = sResult;
        //}

        //private void SetDeptName()
        //{
        //    string sDeptCd = DrCertInfo["DEPT_CD"]?.ToString();

        //    if (!string.IsNullOrEmpty(sDeptCd))
        //    {
        //        StringBuilder strSql = new StringBuilder();

        //        strSql.Clear();
        //        strSql.AppendLine(" ");
        //        strSql.AppendLine(" SELECT A.DEPT_NM ");
        //        strSql.AppendLine("   FROM ACC_DEPT_CD A ");
        //        strSql.AppendLine("  WHERE DEPT_CD = '" + sDeptCd + "' ");

        //        DataTable dt = MySqlDb.GetDataTable(MySqlDb.dbCon, strSql.ToString());
        //        string sResult = dt.Rows[0]["DEPT_NM"]?.ToString();
        //        XRLblDept.Text = sResult;
        //    }
        //}

        //private void SetEmpName()
        //{
        //    string sEmpNm = DrCertInfo["EMP_NM"]?.ToString();
        //    XRLblName.Text = sEmpNm;
        //}

        //private void SetIdentificationNumber()
        //{
        //    string sIdtNo = DrCertInfo["IDT_NO"]?.ToString();
        //    if(!string.IsNullOrEmpty(sIdtNo) & sIdtNo.Length == 13)
        //    {
        //        string sResult = sIdtNo.Substring(0, 6) + "-" + sIdtNo.Substring(6, 7);
        //        XRLblIdtNo.Text = sResult;
        //    }
        //    else
        //    {
        //        XRLblIdtNo.Text = sIdtNo;
        //    }
        //}

        //private void SetAddress()
        //{
        //    string sAddr = DrCertInfo["ADDR"]?.ToString();
        //    XRLblAddr.Text = sAddr;
        //}

        //private void SetEntranceYmd()
        //{
        //    string sYmd = DrCertInfo["ENTRANCE_YMD"]?.ToString();
        //    XRLblEntranceYmd.Text = sYmd;
        //}

        //private void SetGrade()
        //{
        //    string sGrade = DrCertInfo["GRADE_CD"]?.ToString();

        //    if (!string.IsNullOrEmpty(sGrade))
        //    {
        //        StringBuilder strSql = new StringBuilder();

        //        strSql.Clear();
        //        strSql.AppendLine(" ");
        //        strSql.AppendLine(" SELECT A.COM_NM ");
        //        strSql.AppendLine("   FROM COM_BASE_CD A  ");
        //        strSql.AppendLine("  WHERE CD_GB = 'GRADE_CD' ");
        //        strSql.AppendLine("    AND COM_CD = '" + sGrade + "' ");
                
        //        DataTable dt = MySqlDb.GetDataTable(MySqlDb.dbCon, strSql.ToString());
        //        string sResult = dt.Rows[0]["COM_NM"]?.ToString();
        //        XRLblGrade.Text = sResult;
        //    }
        //}

        //private void SetPurpose()
        //{
        //    string sPurp = DrCertInfo["ISSUE_PPS"]?.ToString();
        //    XRLblPurp.Text = sPurp;
        //}

        //private void SetIssueYmd()
        //{
        //    string sIssueYmd = DrCertInfo["ISSUE_YMD"]?.ToString();
        //    string sResult = string.Empty;
        //    if(sIssueYmd.Length == 10)
        //    {
        //        sResult = sIssueYmd.Substring(0, 4) + "년 " + sIssueYmd.Substring(5, 2) + "월 " + sIssueYmd.Substring(8, 2) + "일";
        //    }


        //    XRLblIssueYmd.Text = sResult;
        //}
    }
}
