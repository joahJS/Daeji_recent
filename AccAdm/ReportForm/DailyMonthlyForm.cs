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
    public partial class DailyMonthlyForm : DevExpress.XtraEditors.XtraForm
    {
        public DailyMonthlyForm()
        {
            InitializeComponent();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            StringBuilder strSql = new StringBuilder();
            string ConvertStr;
            string sRetrMonth = DateEditRetr.EditValue.ToString();
            string subMonthFrom = sRetrMonth.Substring(0, 7)+"-01";
            int MonthTo =  Convert.ToInt16(sRetrMonth.Substring(5, 2))+1;
            if (MonthTo > 9)
            {
                ConvertStr = MonthTo.ToString()+"-01";
            }
            else
            {
                ConvertStr = "0"+ MonthTo.ToString()+"-01";
            }
            string subMonthTo = sRetrMonth.Substring(0, 5)+ConvertStr;


            strSql.AppendLine(" SELECT round(sum(G.WEIGHT)/1000) as WEIGHT,");
            strSql.AppendLine("       '슈레더' as '등급',A.J_DATE ");
            strSql.AppendLine(" FROM INLIST A ");
            strSql.AppendLine(" LEFT OUTER JOIN IPCHULGO B ");
            strSql.AppendLine("     ON B.J_ID = A.J_ID ");
            strSql.AppendLine(" LEFT OUTER JOIN JAJAE C");
            strSql.AppendLine("     ON C.J_SERIAL = A.J_SERIAL");
            strSql.AppendLine(" LEFT OUTER JOIN ACC_DEALER_CD D ");
            strSql.AppendLine("     ON D.DEALER_CD = A.J_ID1 ");
            strSql.AppendLine(" LEFT OUTER JOIN SAWON E ");
            strSql.AppendLine("     ON E.S_SERIAL = D.CHRG_ID  ");
            strSql.AppendLine(" LEFT OUTER JOIN COM_BASE_CD F "); 
            strSql.AppendLine("     ON F.COM_NM = C.GUBUN1 ");
            strSql.AppendLine("     AND F.CD_GB = '0001' ");
            strSql.AppendLine(" LEFT OUTER JOIN MESURING G ");
            strSql.AppendLine("     ON G.IPCHULGO_MACHULID = A.J_ID ");
            strSql.AppendLine(" WHERE A.J_DATE >= '"+ subMonthFrom + "'");
            strSql.AppendLine(" AND A.J_DATE < '"+ subMonthTo + "'"); 
            strSql.AppendLine(" AND A.KERATYPE = '매출'");
            strSql.AppendLine(" AND B.KERAGUBUN <> '직송'");
            strSql.AppendLine(" AND F.COM_SUB_CD1 IN ('B' )"); 
            strSql.AppendLine(" group by A.J_Date");
            strSql.AppendLine("   ORDER BY A.J_DATE, D.DEALER_NM ");

            DataTable dtShredder = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT round(sum(G.WEIGHT)/1000) AS WEIGHT, '중량B' as '등급' , A.J_DATE");
            strSql.AppendLine(" FROM INLIST A");
            strSql.AppendLine(" LEFT OUTER JOIN IPCHULGO B");
            strSql.AppendLine(" ON B.J_ID = A.J_ID");
            strSql.AppendLine(" LEFT OUTER JOIN JAJAE C");
            strSql.AppendLine(" ON C.J_SERIAL = A.J_SERIAL");
            strSql.AppendLine(" LEFT OUTER JOIN ACC_DEALER_CD D");
            strSql.AppendLine(" ON D.DEALER_CD = A.J_ID1");
            strSql.AppendLine(" LEFT OUTER JOIN SAWON E");
            strSql.AppendLine(" ON E.S_SERIAL = D.CHRG_ID");
            strSql.AppendLine(" LEFT OUTER JOIN COM_BASE_CD F");
            strSql.AppendLine(" ON F.COM_NM = C.GUBUN1");
            strSql.AppendLine(" AND F.CD_GB = '0001'");
            strSql.AppendLine(" LEFT OUTER JOIN MESURING G");
            strSql.AppendLine(" ON G.IPCHULGO_MACHULID = A.J_ID");
            strSql.AppendLine(" WHERE A.J_DATE >= '" + subMonthFrom + "'");
            strSql.AppendLine(" AND A.J_DATE < '" + subMonthTo + "'");
            strSql.AppendLine(" AND A.KERATYPE = '매출'");
            strSql.AppendLine(" AND B.KERAGUBUN <> '직송'");
            strSql.AppendLine(" AND F.COM_SUB_CD1 IN('A' )");
            strSql.AppendLine(" and G.Gubun1 = '중량B'");
            strSql.AppendLine(" group by A.J_Date");
            strSql.AppendLine(" ORDER BY A.J_DATE, D.DEALER_NM");

            DataTable dtWight = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            strSql.Clear();
            strSql.AppendLine("");
            strSql.AppendLine(" SELECT round(sum(G.WEIGHT)/ 1000) AS WEIGHT, '중량B' as '등급', A.J_DATE");
            strSql.AppendLine(" FROM INLIST A");
            strSql.AppendLine(" LEFT OUTER JOIN IPCHULGO B");
            strSql.AppendLine("   ON B.J_ID = A.J_ID");
            strSql.AppendLine(" LEFT OUTER JOIN JAJAE C");
            strSql.AppendLine("   ON C.J_SERIAL = A.J_SERIAL");
            strSql.AppendLine(" LEFT OUTER JOIN ACC_DEALER_CD D");
            strSql.AppendLine("   ON D.DEALER_CD = A.J_ID1");
            strSql.AppendLine(" LEFT OUTER JOIN SAWON E");
            strSql.AppendLine("   ON E.S_SERIAL = D.CHRG_ID");
            strSql.AppendLine(" LEFT OUTER JOIN COM_BASE_CD F");
            strSql.AppendLine("   ON F.COM_NM = C.GUBUN1");
            strSql.AppendLine("  AND F.CD_GB = '0001'");
            strSql.AppendLine(" LEFT OUTER JOIN MESURING G");
            strSql.AppendLine("   ON G.IPCHULGO_MACHULID = A.J_ID");
            strSql.AppendLine(" WHERE A.J_DATE >= '" + subMonthFrom + "'");
            strSql.AppendLine(" AND A.J_DATE < '" + subMonthTo + "'");
            strSql.AppendLine("  AND A.KERATYPE = '매출'");
            strSql.AppendLine("  AND B.KERAGUBUN <> '직송'");
            strSql.AppendLine("  AND F.COM_SUB_CD1 IN('A' )");
            strSql.AppendLine("  and G.Gubun1 = 'G/S A' in('경량B')");
            strSql.AppendLine("  group by A.J_Date");
            strSql.AppendLine("  ORDER BY A.J_DATE, D.DEALER_NM");

            DataTable dtOutGS = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());


            DataTable dt_sum = new DataTable();
            dt_sum.Columns.Add("STAGET", typeof(double));
            dt_sum.Columns.Add("SWEIGHT", typeof(double));
            dt_sum.Columns.Add("SMINER", typeof(double));

            dt_sum.Columns.Add("WTAGET", typeof(double));
            dt_sum.Columns.Add("WWEIGHT", typeof(double));
            dt_sum.Columns.Add("WMINER", typeof(double));

            dt_sum.Columns.Add("OTAGET", typeof(double));
            dt_sum.Columns.Add("OWEIGHT", typeof(double));
            dt_sum.Columns.Add("OMINER", typeof(double));

            int srow = 0;
            int wrow = 0;
            int orow = 0;

            for(int row = 0; row  < 31; row++)
            {
                dt_sum.Rows.Add();
                
                if (dtShredder.Rows[srow]["J_DATE"].ToString().Trim().Substring(8, 2) == (row + 1).ToString("00"))
                {
                    if (dtShredder.Rows[srow]["등급"].ToString().Trim() == "슈레더")
                    {
                        dt_sum.Rows[row]["STAGET"] = Math.Round(1000.0 / 31.0);
                        dt_sum.Rows[row]["SWEIGHT"] = dtShredder.Rows[srow]["WEIGHT"].ToString().Trim();
                        dt_sum.Rows[row]["SMINER"] = double.Parse(dtShredder.Rows[srow]["WEIGHT"].ToString().Trim()) - Math.Round(1000.0 / 31.0);
                        srow++;

                        if (srow == dtShredder.Rows.Count) srow = dtShredder.Rows.Count - 1;
                    }
                }
                else
                {
                    dt_sum.Rows[row]["STAGET"] = Math.Round(1000.0 / 31.0);
                    dt_sum.Rows[row]["SWEIGHT"] = 0;
                    dt_sum.Rows[row]["SMINER"] = Math.Round(1000.0 / 31.0 * -1);
                }
                //여기지

                if (dtWight.Rows[wrow]["J_DATE"].ToString().Trim().Substring(8, 2) == (row + 1).ToString("00"))
                {
                    if (dtWight.Rows[wrow]["등급"].ToString().Trim() == "중량B")
                    {
                        dt_sum.Rows[row]["WTAGET"] = Math.Round(1000.0 / 31.0);
                        dt_sum.Rows[row]["WWEIGHT"] = dtWight.Rows[wrow]["WEIGHT"].ToString().Trim();
                        dt_sum.Rows[row]["WMINER"] = double.Parse(dtWight.Rows[wrow]["WEIGHT"].ToString().Trim()) - Math.Round(1000.0 / 31.0);
                        wrow++;

                        if (wrow == dtWight.Rows.Count) wrow = dtWight.Rows.Count - 1;
                    }
                }
                else
                {
                    dt_sum.Rows[row]["WTAGET"] = Math.Round(1000.0 / 31.0);
                    dt_sum.Rows[row]["WWEIGHT"] = 0;
                    dt_sum.Rows[row]["WMINER"] = Math.Round(1000.0 / 31.0 * -1);
                }
                //끝
                if (dtOutGS.Rows[orow]["J_DATE"].ToString().Trim().Substring(8, 2) == (row + 1).ToString("00"))
                {
                    if (dtOutGS.Rows[orow]["등급"].ToString().Trim() == "G/S A")
                    {
                        dt_sum.Rows[row]["OTAGET"] = Math.Round(1000.0 / 31.0);
                        dt_sum.Rows[row]["OWEIGHT"] = dtOutGS.Rows[orow]["OEIGHT"].ToString().Trim();
                        dt_sum.Rows[row]["OMINER"] = double.Parse(dtOutGS.Rows[orow]["OWIGHT"].ToString().Trim()) - Math.Round(1000.0 / 31.0);
                        orow++;

                        if (orow == dtOutGS.Rows.Count) orow = dtOutGS.Rows.Count - 1;
                    }
                }
                else
                {
                    dt_sum.Rows[row]["OTAGET"] = Math.Round(1000.0 / 31.0);
                    dt_sum.Rows[row]["OWEIGHT"] = 0;
                    dt_sum.Rows[row]["OMINER"] = Math.Round(1000.0 / 31.0 * -1);
                }
            }

            ReportViewer fm = new ReportViewer(dt_sum, "DailyMonthlyForm");
            fm.ShowDialog();
        }



        private void DailyMonthlyForm_Load(object sender, EventArgs e)
        {
            DateEditRetr.EditValue = DateTime.Now.Date.ToString();
        }

        private void DailyMonthlyForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Dispose();
            }
            else if (e.KeyCode == Keys.F1)
            {
                
            }
            else if (e.KeyCode == Keys.F3)
            {
                
            }
            else if (e.KeyCode == Keys.F4)
            {

            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
            }
        }
    }
}