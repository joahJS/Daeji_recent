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

namespace AccAdm
{
    public partial class AccAutoSlipType : DevExpress.XtraEditors.XtraForm
    {
        public AccAutoSlipType()
        {
            InitializeComponent();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sSlipKind = LkupSlipKind.EditValue.ToString();

            StringBuilder strSql = new StringBuilder();

            //strSql.AppendLine(" ");
            //strSql.AppendLine(" SELECT  IFNULL(A.J_Check, '0') AS CLO  ");
            //strSql.AppendLine("       , A.SUN AS SUN   ");
            //strSql.AppendLine("       , DATE_FORMAT(A.J_DATE, '%Y-%m-%d') AS WEIGHT_YMD ");
            //strSql.AppendLine("       , (SELECT X.COM_CD FROM COM_BASE_CD X WHERE X.CD_GB = 'KERATYPE' AND X.COM_NM = A.KERATYPE) AS GB ");
            //strSql.AppendLine("       , (SELECT X.COM_CD FROM COM_BASE_CD X WHERE X.CD_GB = 'K_NAME' AND X.COM_NM = A.K_NAME) AS GEA ");
            //strSql.AppendLine("       , CONVERT(A.MAIPCHERID, CHAR) AS IN_DEALER_CD ");
            //strSql.AppendLine("       , CONVERT(A.J_ASSIGNID, CHAR) AS OUT_DEALER_CD ");
            //strSql.AppendLine("       , A.J_BNUM AS CARNO   ");
            //strSql.AppendLine("       , A.GUBUN1 AS GRADE ");
            //strSql.AppendLine("       , A.SECONDWEIGHT AS TOT ");
            //strSql.AppendLine("       , A.FIRSTWEIGHT AS EMPAMT ");
            //strSql.AppendLine("       , A.WEIGHT AS OWN ");
            //strSql.AppendLine("       , CASE WHEN A.KERATYPE = '출고' THEN A.OCHAGAM ELSE A.ICHAGAM END AS REDUCE ");
            //strSql.AppendLine("       , CASE WHEN A.KERATYPE = '출고' THEN A.OWEIGHT ELSE A.IWEIGHT END AS AMOUNT ");
            //strSql.AppendLine("       , A.CUSTOMWEIGHT AS CUSTOM ");
            //strSql.AppendLine("       , A.LOSSWEIGHT AS LOSS ");
            //strSql.AppendLine("       , A.IDANGA ");
            //strSql.AppendLine("       , A.ODANGA ");
            //strSql.AppendLine("       , A.IKONGKEP ");
            //strSql.AppendLine("       , A.OKONGKEP ");
            //strSql.AppendLine("       , A.TRANSPORTKUMAK AS COST ");
            //strSql.AppendLine("       , B.NAME AS INSPECTOR ");
            //strSql.AppendLine("       , A.J_STATE AS RED_RMK ");
            //strSql.AppendLine("       , A.GUMSUBIGO AS GUM_RMK ");
            //strSql.AppendLine("       , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.FIRSTTIME, '%H : %I') ELSE DATE_FORMAT(A.SECONDTIME, '%H : %I') END AS TIME1 ");
            //strSql.AppendLine("       , CASE WHEN A.KERATYPE = '입고' THEN DATE_FORMAT(A.SECONDTIME, '%H : %I') ELSE DATE_FORMAT(A.FIRSTTIME, '%H : %I') END AS TIME2 ");
            //strSql.AppendLine("       , DATE_FORMAT(A.AGREE_DATE, '%Y-%m-%d') AS CLO_YMD ");
            //strSql.AppendLine("       , A.JUNPYOID AS JUNPYOID  ");
            //strSql.AppendLine("    FROM MESURING A     ");
            //strSql.AppendLine("    LEFT OUTER JOIN SAWON B  ");
            //strSql.AppendLine("      ON A.GUMSU_SERIAL = B.S_SERIAL  ");
            //strSql.AppendLine("   WHERE A.J_DATE >= '" + sYmd + "' ");
            //strSql.AppendLine("     AND A.J_DATE <= '" + sEndYmd + "' ");
            //strSql.AppendLine("     AND IFNULL(A.J_Check, '0') LIKE CONCAT('%', '" + sCloGb + "', '%') ");
            //strSql.AppendLine("     AND A.KERATYPE LIKE CONCAT('%', '" + sKeraType + "', '%') ");
            //strSql.AppendLine("     AND A.GUBUN1 LIKE CONCAT('%', '" + sGubun + "', '%') ");
            //strSql.AppendLine("   ORDER BY A.J_DATE, A.JUNPYOID ");

            //DataTable dt = MySqlDb.GetDataTable(MySqlDb.dbCon, strSql.ToString());
            //GridRetr.DataSource = dt;
        }
    }
}