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
using DevExpress.XtraGrid.Views.Grid;

namespace AccAdm
{
    public partial class DailyScrabUnitCost : DevExpress.XtraEditors.XtraForm
    {
        public DailyScrabUnitCost()
        {
            InitializeComponent();
        }

        private void DailyScrabUnitCost_Load(object sender, EventArgs e)
        {
            DateEditDT.EditValue = DateTime.Now;
        }
        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sDate = DateEditDT.EditValue.ToString().Replace("-", "").Substring(0, 8);
            GetGridRetr(sDate);
            GetYKGsCostDay(sDate);
            GetYKGsCostMon(sDate);
            GetYKAverageDay(sDate);
            GetYKAverageMon(sDate);
            GetScrabDayEx(sDate);
            GetScrabDay(sDate);
            GetScrabMon(sDate);
            GetGsAmount(sDate);
            GetGsAmountTotal(sDate);
        }

        

        private void GetGridRetr(string sDate)
        {
          
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.Append("");
            strSql.Append("SELECT C.j_bnum      ");
            strSql.Append("	     ,C.Gubun1      ");
            strSql.Append("	     ,D.C_Company   ");
            strSql.Append("	     ,C.gumsubigo   ");
            strSql.Append("	     ,D.J_Region    ");
            strSql.Append("	     ,E.name        ");
            strSql.Append("      ,A.danjung    ");
            strSql.Append("      ,A.Halin      ");
            strSql.Append("      ,A.MiDanga    ");
            strSql.Append("      ,A.Danga      ");
            strSql.Append("      ,IFNULL(A.CKongKep, 0) / IFNULL(A.danjung, 0)AS CKongKep  ");
            strSql.Append("      ,IFNULL(A.CKongKep, 0) / IFNULL(A.danjung, 0) + IFNULL(A.Danga,0) AS MaIpDanga ");
            strSql.Append("      ,IFNULL(A.MiDanga,0) - IFNULL(A.Danga,0) - IFNULL(A.CKongKep, 0) / IFNULL(A.danjung, 0)AS Difference  ");
            strSql.Append("      ,A.iKongKep   ");
            strSql.Append("      ,IFNULL(A.CKongKep,0) + IFNULL(A.iKongKep,0) AS MaIp  ");
            strSql.Append("  FROM inlist A, jajae B, mesuring C, custom D, sawon E     ");
            strSql.Append(" WHERE A.j_serial = B.j_serial     ");
            strSql.Append("   AND A.J_ID = C.ipchulgo_maipid  ");
            strSql.Append("   AND A.J_ID1 = D.C_Serial        ");
            strSql.Append("   AND D.Damdang = E.S_Serial      ");
            strSql.Append("   AND A.KeraType = '매입'         ");
            strSql.Append("   AND A.J_Date ='" + sDate + "'       ");
            strSql.Append("   AND A.j_lotno = '1'             ");
            strSql.Append("   AND B.daegubun <> '슈레더'      ");
            strSql.Append("   AND B.Gubun1 <> '인센티브'      ");
            strSql.Append(" ORDER BY  E.name, binary(D.c_company) asc    ");
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0) GridRetr.DataSource = dt;


        }
        private void GetYKGsCostDay(string sDate)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.Append("");
            strSql.Append("SELECT sum(a.danjung) AS '1'  ");
            strSql.Append("      ,sum(a.CKongKep)   ");
            strSql.Append("      ,sum(a.KongKep)    ");
            strSql.Append("      ,sum(a.KongKep) -  sum(a.CKongKep)AS '2'" );
            strSql.Append("      ,(sum(a.KongKep) -  sum(a.CKongKep))/sum(a.danjung) AS '3'  ");
            strSql.Append(" FROM inlist a, jajae b ");
            strSql.Append("WHERE a.j_serial = b.j_serial ");
            strSql.Append("  AND a.KeraType = '매출'     ");
            strSql.Append("  AND a.J_Date = '"+sDate+"'   ");
            strSql.Append("  AND a.j_id1 = '3018'        ");
            strSql.Append("  AND a.j_lotno = '1'         ");
            strSql.Append("  AND b.daegubun <> '슈레더'  ");
            strSql.Append("  AND b.Gubun1 <> '생철A'     ");
            strSql.Append("  AND b.Gubun1 <> '생철B'     ");
            strSql.Append("  AND b.Gubun1 <> '생철 AH'   ");
            strSql.Append("  AND b.Gubun1 <> '생철 AL'   ");
            strSql.Append("  AND b.Gubun1 <> '중량A'     ");
            strSql.Append("  AND b.Gubun1 <> '중량 AL'   ");
            strSql.Append("  AND b.Gubun1 <> '중량 A - B' ");
            strSql.Append("  AND b.Gubun1 <> '중량 - ABL' ");
            strSql.Append("  AND b.Gubun1 <> '중량 AH'  ");
            strSql.Append("  AND b.Gubun1 <> '중량_ABH'  ");
            strSql.Append("  AND b.Gubun1 <> '중량B'     ");
            strSql.Append("  AND b.Gubun1 <> '인센티브'  ");
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            TxtYkGsDay1.Text = "";
            TxtYkGsDay2.Text = "";
            TxtYkGsDay3.Text = "";
            TxtYkGsDay1.Text = dt.Rows[0]["1"].ToString();
            TxtYkGsDay2.Text = dt.Rows[0]["2"].ToString();
            TxtYkGsDay3.Text = dt.Rows[0]["3"].ToString();
           

        }
        private void GetYKGsCostMon(string sDate)
        {
            string sFrom = sDate;
            sFrom =sFrom.Substring(0, 6)+"01";

            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.Append("");           
            strSql.Append(" SELECT sum(a.danjung)AS '1', sum(a.CKongKep), sum(a.KongKep),sum(a.KongKep)-sum(a.CKongKep)AS '2',(sum(a.KongKep)-sum(a.CKongKep))/sum(a.danjung) AS '3'    ");
            strSql.Append("   FROM inlist a, jajae b      ");
            strSql.Append("  WHERE a.j_serial = b.j_serial       ");
            strSql.Append("    AND a.KeraType = '매출'          ");
            strSql.Append("    AND a.J_Date >= '" + sFrom + "'         ");
            strSql.Append("    AND a.J_Date <= '" + sDate + "'        ");
            strSql.Append("    AND a.j_id1 = '3018'             ");
            strSql.Append("    AND a.j_lotno = '1'               ");
            strSql.Append("    AND b.daegubun <> '슈레더'       ");
            strSql.Append("    AND b.Gubun1 <> '생철A'           ");
            strSql.Append("    AND b.Gubun1 <> '생철B'           ");
            strSql.Append("    AND b.Gubun1 <> '생철 AH'        ");
            strSql.Append("    AND b.Gubun1 <> '생철 AL'         ");
            strSql.Append("    AND b.Gubun1 <> '중량A'          ");
            strSql.Append("    AND b.Gubun1 <> '중량 AL'         ");
            strSql.Append("    AND b.Gubun1 <> '중량 A-B'        ");
            strSql.Append("    AND b.Gubun1 <> '중량-ABL'        ");
            strSql.Append("    AND b.Gubun1 <> '중량 AH'        ");
            strSql.Append("    AND b.Gubun1 <> '중량_ABH'        ");
            strSql.Append("    AND b.Gubun1 <> '중량B'        ");
            strSql.Append("    AND b.Gubun1 <> '인센티브'        ");
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            TxtYkGsMon1.Text = "";
            TxtYkGsMon2.Text = "";
            TxtYkGsMon3.Text = "";
            TxtYkGsMon1.Text = dt.Rows[0]["1"].ToString();
            TxtYkGsMon2.Text = dt.Rows[0]["2"].ToString();
            TxtYkGsMon3.Text = dt.Rows[0]["3"].ToString();

        }
        private void GetYKAverageMon(string sDate)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.Append("");
            strSql.Append(" SELECT sum(a.danjung) AS '1', sum(a.CKongKep), sum(a.KongKep) ,sum(a.KongKep)-sum(a.CKongKep)AS '2',(sum(a.KongKep)-sum(a.CKongKep))/sum(a.danjung) AS '3'     ");
            strSql.Append("   FROM  inlist a, jajae b    ");
            strSql.Append("   WHERE a.j_serial = b.j_serial      ");
            strSql.Append("   AND   a.KeraType = '매출'          ");
            strSql.Append("   AND   a.J_Date = '"+sDate+"'        ");
            strSql.Append("   AND   a.j_id1 = '3018'             ");
            strSql.Append("   AND   b.Gubun1 <> '인센티브'       ");
            strSql.Append("   AND   a.j_lotno = '1'              ");
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            TxtYkaverageDay1.Text = "";
            TxtYkaverageDay2.Text = "";
            TxtYkaverageDay3.Text = "";
            TxtYkaverageDay1.Text = dt.Rows[0]["1"].ToString();
            TxtYkaverageDay2.Text = dt.Rows[0]["2"].ToString();
            TxtYkaverageDay3.Text = dt.Rows[0]["3"].ToString();
        }
        private void GetYKAverageDay(string sDate)
        {
            string sFrom = sDate;
            sFrom = sFrom.Substring(0, 6) + "01";

            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.Append("");
            strSql.Append(" SELECT sum(a.danjung) AS '1', sum(a.CKongKep), sum(a.KongKep) ,sum(a.KongKep)-sum(a.CKongKep)AS '2',(sum(a.KongKep)-sum(a.CKongKep))/sum(a.danjung) AS '3'     ");
            strSql.Append("   FROM  inlist a, jajae b    ");
            strSql.Append("   WHERE a.j_serial = b.j_serial      ");
            strSql.Append("   AND   a.KeraType = '매출'          ");
            strSql.Append("   AND   a.J_Date >= '" + sFrom + "'        ");
            strSql.Append("   AND   a.J_Date <= '" + sDate + "'        ");
            strSql.Append("   AND   a.j_id1 = '3018'             ");
            strSql.Append("   AND   b.Gubun1 <> '인센티브'       ");
            strSql.Append("   AND   a.j_lotno = '1'              ");
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            TxtYkaverageMon1.Text = "";
            TxtYkaverageMon2.Text = "";
            TxtYkaverageMon3.Text = "";
            TxtYkaverageMon1.Text = dt.Rows[0]["1"].ToString();
            TxtYkaverageMon2.Text = dt.Rows[0]["2"].ToString();
            TxtYkaverageMon3.Text = dt.Rows[0]["3"].ToString();
        }
        private void GetScrabDayEx(string sDate)
        {          
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.Append("");
            strSql.Append(" SELECT sum(a.danjung) as '1', sum(a.Halin) as '2', sum(a.danga), sum(a.CKongKep), sum(a.iKongKep) AS '7'      ");
            strSql.Append(" ,sum(a.iKongKep) /sum(a.danjung)AS '3'                     ");
            strSql.Append(" ,(sum(a.CKongKep)+sum(a.iKongKep))/sum(a.danjung) AS '4'   ");
            strSql.Append(" ,sum(a.CKongKep)+sum(a.iKongKep) AS '8'                    ");
            strSql.Append(" from inlist a, jajae b    ");
            strSql.Append(" WHERE a.j_serial = b.j_serial     ");
            strSql.Append(" AND   a.KeraType = '매입'         ");
            strSql.Append(" AND   a.J_Date = '"+sDate+"'       ");
            strSql.Append(" AND   a.j_lotno = '1'             ");
            strSql.Append(" AND   b.daegubun <> '슈레더'      ");
            strSql.Append(" AND   b.Gubun1 <> '인센티브'      ");
            strSql.Append(" AND   b.gubun1 <> '수입선반'      ");
            strSql.Append(" AND   b.daegubun <> '수입경량'    ");
            strSql.Append(" AND   b.daegubun <> '수입철사'    ");
            strSql.Append(" AND   b.daegubun <> '수입압축'    ");
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            TxtScrabDayEx1.Text = "";
            TxtScrabDayEx2.Text = "";
            TxtScrabDayEx3.Text = "";
            TxtScrabDayEx4.Text = "";
            TxtScrabDayEx5.Text = "";
            TxtScrabDayEx6.Text = "";
            TxtScrabDayEx7.Text = "";
            TxtScrabDayEx8.Text = "";
            TxtScrabDayEx1.Text = dt.Rows[0]["1"].ToString();
            TxtScrabDayEx2.Text = dt.Rows[0]["2"].ToString();
            TxtScrabDayEx3.Text = dt.Rows[0]["3"].ToString();
            TxtScrabDayEx4.Text = dt.Rows[0]["4"].ToString();
            //TxtScrabDayEx5.Text = dt.Rows[0]["5"].ToString();
            //TxtScrabDayEx6.Text = dt.Rows[0]["6"].ToString();
            TxtScrabDayEx7.Text = dt.Rows[0]["7"].ToString();
            TxtScrabDayEx8.Text = dt.Rows[0]["8"].ToString();
        }
        private void GetScrabDay(string sDate)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.Append("");
            strSql.Append(" SELECT sum(a.danjung) as '1', sum(a.Halin) as '2', sum(a.danga), sum(a.CKongKep), sum(a.iKongKep) AS '7'      ");
            strSql.Append(" ,sum(a.iKongKep) /sum(a.danjung)AS '3'                     ");
            strSql.Append(" ,(sum(a.CKongKep)+sum(a.iKongKep))/sum(a.danjung) AS '4'   ");
            strSql.Append(" ,sum(a.CKongKep)+sum(a.iKongKep) AS '8'                    ");
            strSql.Append(" from inlist a, jajae b    ");
            strSql.Append(" WHERE a.j_serial = b.j_serial     ");
            strSql.Append(" AND   a.KeraType = '매입'         ");
            strSql.Append(" AND   a.J_Date = '" + sDate + "'       ");
            strSql.Append(" AND   a.j_lotno = '1'             ");
            strSql.Append(" AND   b.daegubun <> '슈레더'      ");
            strSql.Append(" AND   b.Gubun1 <> '인센티브'      ");
      
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            TxtScrabDay1.Text = "";
            TxtScrabDay2.Text = "";
            TxtScrabDay3.Text = "";
            TxtScrabDay4.Text = "";
            TxtScrabDay5.Text = "";
            TxtScrabDay6.Text = "";
            TxtScrabDay7.Text = "";
            TxtScrabDay8.Text = "";
            TxtScrabDay1.Text = dt.Rows[0]["1"].ToString();
            TxtScrabDay2.Text = dt.Rows[0]["2"].ToString();
            TxtScrabDay3.Text = dt.Rows[0]["3"].ToString();
            TxtScrabDay4.Text = dt.Rows[0]["4"].ToString();
            //TxtScrabDay5.Text = dt.Rows[0]["5"].ToString();
            //TxtScrabDay6.Text = dt.Rows[0]["6"].ToString();
            TxtScrabDay7.Text = dt.Rows[0]["7"].ToString();
            TxtScrabDay8.Text = dt.Rows[0]["8"].ToString();
        }
        private void GetScrabMon(string sDate)
        {
            string sFrom = sDate;
            sFrom = sFrom.Substring(0, 6) + "01";
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.Append("");
            strSql.Append(" SELECT sum(a.danjung) as '1', sum(a.Halin) as '2', sum(a.danga), sum(a.CKongKep), sum(a.iKongKep) AS '7'      ");
            strSql.Append(" ,sum(a.iKongKep) /sum(a.danjung)AS '3'                     ");
            strSql.Append(" ,(sum(a.CKongKep)+sum(a.iKongKep))/sum(a.danjung) AS '4'   ");
            strSql.Append(" ,sum(a.CKongKep)+sum(a.iKongKep) AS '8'                    ");
            strSql.Append(" from inlist a, jajae b    ");
            strSql.Append(" WHERE a.j_serial = b.j_serial     ");
            strSql.Append(" AND   a.KeraType = '매입'         ");
            strSql.Append(" AND   a.J_Date >= '" + sFrom + "'       ");
            strSql.Append(" AND   a.J_Date <= '" + sDate + "'       ");
            strSql.Append(" AND   a.j_lotno = '1'             ");
            strSql.Append(" AND   b.daegubun <> '슈레더'      ");
            strSql.Append(" AND   b.Gubun1 <> '인센티브'      ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            TxtScrabMon1.Text = "";
            TxtScrabMon2.Text = "";
            TxtScrabMon3.Text = "";
            TxtScrabMon4.Text = "";
            TxtScrabMon5.Text = "";
            TxtScrabMon6.Text = "";
            TxtScrabMon7.Text = "";
            TxtScrabMon8.Text = "";
             TxtScrabMon1.Text = dt.Rows[0]["1"].ToString();
             TxtScrabMon2.Text = dt.Rows[0]["2"].ToString();
             TxtScrabMon3.Text = dt.Rows[0]["3"].ToString();
            TxtScrabMon4.Text = dt.Rows[0]["4"].ToString();
            //TxtScrabMon5.Text = dt.Rows[0]["5"].ToString();
            //TxtScrabMon6.Text = dt.Rows[0]["6"].ToString();
            TxtScrabMon7.Text = dt.Rows[0]["7"].ToString();
            TxtScrabMon8.Text = dt.Rows[0]["8"].ToString();
        }
        private void GetGsAmount(string sDate)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.Append("");
            strSql.Append(" SELECT sum(a.danjung) AS '1', sum(a.Halin) AS '2', sum(a.danga), sum(a.CKongKep), sum(a.iKongKep) AS '7'     ");
            strSql.Append(" ,sum(a.iKongKep)/sum(a.danjung) AS '3'          ");
            strSql.Append(" ,( sum(a.CKongKep)+ sum(a.iKongKep) )/sum(a.danjung) AS '4'     ");
            strSql.Append(" ,sum(a.CKongKep)+ sum(a.iKongKep) AS '8'        ");
            strSql.Append(" from inlist a, jajae b   ");
            strSql.Append(" WHERE a.j_serial = b.j_serial    ");
            strSql.Append(" AND   a.KeraType = '매입'        ");
            strSql.Append(" AND   a.J_Date = '" + sDate + "'      ");
            strSql.Append(" AND   a.j_lotno = 1              ");
            strSql.Append(" AND   b.daegubun <> '슈레더'     ");
            strSql.Append(" AND   b.Gubun1 <> '생철A'        ");
            strSql.Append(" AND   b.Gubun1 <> '생철B'        ");
            strSql.Append(" AND   b.Gubun1 <> '생철 BL'      ");
            strSql.Append(" AND   b.Gubun1 <> '생철 AH'      ");
            strSql.Append(" AND   b.Gubun1 <> '생철 AL'      ");
            strSql.Append(" AND   b.Gubun1 <> '중량A'        ");
            strSql.Append(" AND   b.Gubun1 <> '중량 AL'      ");
            strSql.Append(" AND   b.Gubun1 <> '중량 A-B'     ");
            strSql.Append(" AND   b.Gubun1 <> '중량-ABL'     ");
            strSql.Append(" AND   b.Gubun1 <> '중량 AH'      ");
            strSql.Append(" AND   b.Gubun1 <> '중량_ABH'     ");
            strSql.Append(" AND   b.Gubun1 <> '중량B'        ");
            strSql.Append(" AND   b.Gubun1 <> '기계작업철'   ");
            strSql.Append(" AND   b.gubun1 <>' 인센티브'     ");
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            TxtGsAmount1.Text = "";
            TxtGsAmount2.Text = "";
            TxtGsAmount3.Text = "";
            TxtGsAmount4.Text = "";
            TxtGsAmount5.Text = "";
            TxtGsAmount6.Text = "";
            TxtGsAmount7.Text = "";
            TxtGsAmount8.Text = "";
             TxtGsAmount1.Text = dt.Rows[0]["1"].ToString();
             TxtGsAmount2.Text = dt.Rows[0]["2"].ToString();
             TxtGsAmount3.Text = dt.Rows[0]["3"].ToString();
            TxtGsAmount4.Text = dt.Rows[0]["4"].ToString();
            //TxtGsAmount5.Text = dt.Rows[0]["5"].ToString();
            //TxtGsAmount6.Text = dt.Rows[0]["6"].ToString();
            TxtGsAmount7.Text = dt.Rows[0]["7"].ToString();
            TxtGsAmount8.Text = dt.Rows[0]["8"].ToString();
        }
        private void GetGsAmountTotal(string sDate)
        {
            string sFrom = sDate;
            sFrom = sFrom.Substring(0, 6) + "01";
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.Append("");
            strSql.Append(" SELECT sum(a.danjung) AS '1', sum(a.Halin) AS '2', sum(a.danga), sum(a.CKongKep), sum(a.iKongKep) AS '7'     ");
            strSql.Append(" ,sum(a.iKongKep)/sum(a.danjung) AS '3'          ");
            strSql.Append(" ,( sum(a.CKongKep)+ sum(a.iKongKep) )/sum(a.danjung) AS '4'     ");
            strSql.Append(" ,sum(a.CKongKep)+ sum(a.iKongKep) AS '8'        ");
            strSql.Append(" from inlist a, jajae b   ");
            strSql.Append(" WHERE a.j_serial = b.j_serial    ");
            strSql.Append(" AND   a.KeraType = '매입'        ");
            strSql.Append(" AND   a.J_Date >= '"+sFrom+"'      ");
            strSql.Append(" AND   a.J_Date <= '" + sDate + "'      ");
            strSql.Append(" AND   a.j_lotno = 1              ");
            strSql.Append(" AND   b.daegubun <> '슈레더'     ");
            strSql.Append(" AND   b.Gubun1 <> '생철A'        ");
            strSql.Append(" AND   b.Gubun1 <> '생철B'        ");
            strSql.Append(" AND   b.Gubun1 <> '생철 BL'      ");
            strSql.Append(" AND   b.Gubun1 <> '생철 AH'      ");
            strSql.Append(" AND   b.Gubun1 <> '생철 AL'      ");
            strSql.Append(" AND   b.Gubun1 <> '중량A'        ");
            strSql.Append(" AND   b.Gubun1 <> '중량 AL'      ");
            strSql.Append(" AND   b.Gubun1 <> '중량 A-B'     ");
            strSql.Append(" AND   b.Gubun1 <> '중량-ABL'     ");
            strSql.Append(" AND   b.Gubun1 <> '중량 AH'      ");
            strSql.Append(" AND   b.Gubun1 <> '중량_ABH'     ");
            strSql.Append(" AND   b.Gubun1 <> '중량B'        ");
            strSql.Append(" AND   b.Gubun1 <> '기계작업철'   ");
            strSql.Append(" AND   b.gubun1 <>' 인센티브'     ");
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            TxtGsAmountTotal1.Text = "";
            TxtGsAmountTotal2.Text = "";
            TxtGsAmountTotal3.Text = "";
            TxtGsAmountTotal4.Text = "";
            TxtGsAmountTotal5.Text = "";
            TxtGsAmountTotal6.Text = "";
            TxtGsAmountTotal7.Text = "";
            TxtGsAmountTotal8.Text = "";
            TxtGsAmountTotal1.Text = dt.Rows[0]["1"].ToString();
            TxtGsAmountTotal2.Text = dt.Rows[0]["2"].ToString();
            TxtGsAmountTotal3.Text = dt.Rows[0]["3"].ToString();
            TxtGsAmountTotal4.Text = dt.Rows[0]["4"].ToString();
            //TxtGsAmountTotal5.Text = dt.Rows[0]["5"].ToString();
            //double s = 9.9;
            //TxtGsAmountTotal5.Text = Math.Round(s).ToString();
            //TxtGsAmountTotal6.Text = dt.Rows[0]["6"].ToString();
            TxtGsAmountTotal7.Text = dt.Rows[0]["7"].ToString();
            TxtGsAmountTotal8.Text = dt.Rows[0]["8"].ToString();
        }
        private void GridViewRetr_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {

            if (e.Column.FieldName == "Difference")
            {
                string s = GridViewRetr.GetRowCellValue(e.RowHandle, "Difference").ToString() is null ? "" : GridViewRetr.GetRowCellValue(e.RowHandle, "Difference").ToString();
                if (double.Parse(s)<0)
                {
                    e.Appearance.ForeColor = Color.Red;
                }

            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void DailyScrabUnitCost_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) //종료
            {
                
            }
            else if (e.KeyCode == Keys.F1)//추가
            {

            }
            else if (e.KeyCode == Keys.F3)//저장
            {

            }
            else if (e.KeyCode == Keys.F4)//삭제
            {

            }
            else if (e.KeyCode == Keys.F5)//조회
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)//엑셀
            {

            }
            else if (e.KeyCode == Keys.F8)//출력
            {

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
    }
}