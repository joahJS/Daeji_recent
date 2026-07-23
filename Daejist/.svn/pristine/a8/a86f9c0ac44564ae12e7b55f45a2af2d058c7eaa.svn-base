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
    public partial class TeamPurcPerformance : DevExpress.XtraEditors.XtraForm
    {
        public TeamPurcPerformance()
        {
            InitializeComponent();
        }

        private void TeamPurcPerformance_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;
        }

        private void BtnPurcRetr_Click(object sender, EventArgs e)
        {
            string sYmdFrom = DateEditFrom.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            
            if(string.IsNullOrEmpty(sYmdFrom) || string.IsNullOrEmpty(sYmdTo))
            {
                XtraMessageBox.Show("조회기간을 설정하세요.");
                DateEditFrom.Focus();
                DateEditFrom.SelectAll();

                return;
            }


            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT SUM(A.DANJUNG), SUM(A.KONGKEP), SUM(A.CKONGKEP), SUM(A.HALIN)  ");
            strSql.AppendLine("   FROM INLIST A  ");
            strSql.AppendLine("   LEFT OUTER JOIN JAJAE B ");
            strSql.AppendLine("     ON A.J_SERIAL = B.J_SERIAL ");
            strSql.AppendLine("    AND B.DAEGUBUN = '고철A' ");
            strSql.AppendLine("    AND B.GUBUN1 <> '인센티브' ");
            strSql.AppendLine("  WHERE A.J_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            strSql.AppendLine("    AND A.KERATYPE = '매출' ");
            strSql.AppendLine("    AND A.J_LOTNO = '1' ");
            strSql.AppendLine("    AND A.J_ID1 = '3018' ");

            DataTable dtGsHeavy = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            string sDanjung = dtGsHeavy.Rows[0]["SUM(A.DANJUNG)"]?.ToString();
            string sHeavy = TxtHeavyWeight.EditValue?.ToString();

            double dMoney = 0;

            if (string.IsNullOrEmpty(sDanjung) || string.IsNullOrEmpty(sHeavy)) // 중량 스크랩 매출평균단가 구하기
            {
                XtraMessageBox.Show("기간내 스크랩 중량 매출이 없습니다. 적용단가를 입력하고 조회해 주세요.");
                return;
            }
            else if (!string.IsNullOrEmpty(sDanjung))
            {
                double dKongKep = dtGsHeavy.Rows[0]["SUM(A.KONGKEP)"] == null ? 0 : Convert.ToDouble(dtGsHeavy.Rows[0]["SUM(A.KONGKEP)"]);
                double dCKongKep = dtGsHeavy.Rows[0]["SUM(A.CKONGKEP)"] == null ? 0 : Convert.ToDouble(dtGsHeavy.Rows[0]["SUM(A.CKONGKEP)"]);

                dMoney = dKongKep - dCKongKep;
                TxtHvPurcUtPrc.EditValue = Math.Round(((dMoney - (dMoney * (Convert.ToDouble(TxtDiscount.EditValue) / 100) / 365* 91)) / Convert.ToDouble(sDanjung)), 1);
                
                TxtHvAplyUtPrc.EditValue = Math.Round((((dMoney + (Convert.ToDouble(sDanjung) * Convert.ToDouble(TxtIncentive.EditValue)))
                                    - ((dMoney + (Convert.ToDouble(sDanjung) * Convert.ToDouble(TxtIncentive.EditValue)))
                                    * (Convert.ToDouble(TxtDiscount.EditValue) / 100) / 365 * 91)) / Convert.ToDouble(sDanjung)), 1);
                
            }
            else
            {
                TxtHvAplyUtPrc.EditValue = TxtHeavyWeight.EditValue;
            }

            // 중량 스크랩 자료 뿌리기
            double dHvDanga = Convert.ToDouble(TxtHvAplyUtPrc.EditValue); //적용매출단가
            double dHvMajin = Convert.ToDouble(TxtHvMgnGoal.EditValue); //스크랩중량마진

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT *  ");
            strSql.AppendLine("   FROM ( ");
            strSql.AppendLine("          SELECT COALESCE(D.EMP_NM,'전체') AS EMP_NM ");
            strSql.AppendLine("               , SUM(A.DANJUNG) AS DANJUNG ");
            strSql.AppendLine("               , SUM(A.HALIN) AS HALIN  ");
            strSql.AppendLine("               , SUM(A.IKONGKEP) AS IKONGKEP  ");
            strSql.AppendLine("               , SUM(A.CKONGKEP) AS CKONGKEP ");
            strSql.AppendLine("               , 0.00 AS DDANGA  ");
            strSql.AppendLine("               , 0.00 AS HEAVYMAJIN  ");
            strSql.AppendLine("               , 0 AS MAECHUL  ");
            strSql.AppendLine("               , D.GRADE_CD ");
            strSql.AppendLine("            FROM INLIST A  ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE B  ");
            strSql.AppendLine("              ON B.J_SERIAL = A.J_SERIAL  ");
            strSql.AppendLine("             AND B.DAEGUBUN = '고철A'  ");
            strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD C  ");
            strSql.AppendLine("              ON C.DEALER_CD = A.J_ID1  ");
            strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS D  ");
            strSql.AppendLine("              ON D.EMP_ID = C.CHRG_ID  ");
            strSql.AppendLine("           WHERE A.KERATYPE = '매입'  ");
            strSql.AppendLine("             AND A.J_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("             AND A.J_DATE <= '" + sYmdTo + "' ");
            strSql.AppendLine("             AND A.J_LOTNO = '1' ");
            strSql.AppendLine("             AND D.EMP_NM <> '김홍철'  ");
            strSql.AppendLine("             AND D.EMP_NM <> '기타'  ");
            strSql.AppendLine("           GROUP BY D.EMP_NM WITH ROLLUP ");
            strSql.AppendLine("           )PD ");
            strSql.AppendLine("  ORDER BY EMP_NM = '전체', GRADE_CD, EMP_NM ");

            DataTable dtHeavy = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            double dTotMeaChul = 0;
            for (int i = 0; i < dtHeavy.Rows.Count; i++) //도착매입단가, 중량마진, 매출이익 내부 계산 후 그리드 바인딩
            {
                Cursor = Cursors.WaitCursor;

                double dIKongKep = dtHeavy.Rows[i]["IKONGKEP"] == null ? 0 : Convert.ToDouble(dtHeavy.Rows[i]["IKONGKEP"]);
                double dCKongKep = dtHeavy.Rows[i]["CKONGKEP"] == null ? 0 : Convert.ToDouble(dtHeavy.Rows[i]["CKONGKEP"]);
                double dDanjung = dtHeavy.Rows[i]["DANJUNG"] == null ? 0 : Convert.ToDouble(dtHeavy.Rows[i]["DANJUNG"]);
                //도착매입단가
                double ArrvPurcUtPrc = Math.Round(((dIKongKep + dCKongKep) / dDanjung), 1);
                dtHeavy.Rows[i]["DDANGA"] = ArrvPurcUtPrc;
                
                //스크랩 중량 마진
                double dWgtMargin = Math.Round((dHvDanga - ArrvPurcUtPrc - dHvMajin), 1);
                dtHeavy.Rows[i]["HEAVYMAJIN"] = dWgtMargin;
                
                //매출이익
                double dSaleProfit = Math.Round((dDanjung * dWgtMargin), 1);
                dtHeavy.Rows[i]["MAECHUL"] = dSaleProfit;

                string sEmpNm = dtHeavy.Rows[i]["EMP_NM"]?.ToString();
                if (!string.IsNullOrEmpty(sEmpNm) && sEmpNm.Equals("전체"))
                {
                    dtHeavy.Rows[i]["MAECHUL"] = dTotMeaChul;
                }
                else
                {
                    dTotMeaChul += dSaleProfit;
                }

                Cursor = Cursors.Default;
            }

            GridHeavy.DataSource = dtHeavy;
            
            // 경량 스크랩 매출평균단가 구하기

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT SUM(A.DANJUNG) ");
            strSql.AppendLine("      , SUM(A.KONGKEP) ");
            strSql.AppendLine("      , SUM(A.CKONGKEP) ");
            strSql.AppendLine("      , SUM(A.HALIN)  ");
            strSql.AppendLine("   FROM INLIST A ");
            strSql.AppendLine("   LEFT OUTER JOIN JAJAE B ");
            strSql.AppendLine("     ON B.J_SERIAL = A.J_SERIAL ");
            strSql.AppendLine("  WHERE A.KERATYPE = '매출' ");
            strSql.AppendLine("    AND A.J_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            strSql.AppendLine("    AND A.J_LOTNO = '1' ");
            strSql.AppendLine("    AND A.J_ID1 = '3018' ");
            strSql.AppendLine("    AND B.DAEGUBUN = '고철B' ");
            strSql.AppendLine("    AND B.GUBUN1 <> '인센티브' ");

            DataTable dtGsLight = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            string sLtDanjung = dtGsLight.Rows[0]["SUM(A.DANJUNG)"]?.ToString();
            string sLight = TxtLightWeight.EditValue?.ToString();

            double dLtMoney = 0;
            if (string.IsNullOrEmpty(sLtDanjung) && string.IsNullOrEmpty(sLight))
            {
                XtraMessageBox.Show("기간내 스크랩 경량 매출이 없습니다. 적용단가를 입력하고 조회해 주세요.");
                return;
            }
            else if (!string.IsNullOrEmpty(sLtDanjung))
            {
                double dKongKep = dtGsHeavy.Rows[0]["SUM(A.KONGKEP)"] == null ? 0 : Convert.ToDouble(dtGsHeavy.Rows[0]["SUM(A.KONGKEP)"]);
                double dCKongKep = dtGsHeavy.Rows[0]["SUM(A.CKONGKEP)"] == null ? 0 : Convert.ToDouble(dtGsHeavy.Rows[0]["SUM(A.CKONGKEP)"]);

                dLtMoney = (Convert.ToDouble(dKongKep) - Convert.ToDouble(dCKongKep));

                TxtLtPurcUtPrc.EditValue = Math.Round(((dLtMoney - (dLtMoney * (Convert.ToDouble(TxtDiscount.EditValue) / 100) / 365 * 91))
                                       / Convert.ToDouble(sLtDanjung)), 1);

                TxtLtAplyUtPrc.EditValue = Math.Round(((dLtMoney + (Convert.ToDouble(sLtDanjung) * Convert.ToDouble(TxtIncentive.EditValue)))
                                - ((dLtMoney + (Convert.ToDouble(sLtDanjung) * Convert.ToDouble(TxtIncentive.EditValue)))
                                * (Convert.ToDouble(TxtDiscount.EditValue) / 100) / 365 * 91)) / Convert.ToDouble(sLtDanjung), 1);
            }
            else
            {
                TxtLtAplyUtPrc.EditValue = TxtLightWeight.EditValue;
            }
            
            double dLtDanga = Convert.ToDouble(TxtLtAplyUtPrc.EditValue); //적용매출단가
            double dLtMajin = Convert.ToDouble(TxtLtMgnGoal.EditValue); //스크랩경량마진

            Cursor = Cursors.WaitCursor;

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT * ");
            strSql.AppendLine("   FROM ( ");
            strSql.AppendLine("          SELECT COALESCE(D.EMP_NM,'전체') AS EMP_NM ");
            strSql.AppendLine("               , COALESCE(B.GUBUN1,'합계') AS GUBUN1 ");
            strSql.AppendLine("               , SUM(A.DANJUNG) AS DANJUNG  ");
            strSql.AppendLine("               , SUM(A.HALIN) AS HALIN  ");
            strSql.AppendLine("               , SUM(A.IKONGKEP) AS IKONGKEP  ");
            strSql.AppendLine("               , SUM(A.CKONGKEP) AS CKONGKEP  ");
            strSql.AppendLine("               , 0.0 AS DDANGA  ");
            strSql.AppendLine("               , 0.0 AS LIGHTMAJIN  ");
            strSql.AppendLine("               , 0 AS MAECHUL ");
            strSql.AppendLine("               , D.GRADE_CD ");
            strSql.AppendLine("            FROM INLIST A ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE B ");
            strSql.AppendLine("              ON B.J_SERIAL = A.J_SERIAL ");
            strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD C ");
            strSql.AppendLine("              ON C.DEALER_CD = A.J_ID1 ");
            strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS D  ");
            strSql.AppendLine("              ON D.EMP_ID = C.CHRG_ID ");
            strSql.AppendLine("           WHERE A.KERATYPE = '매입' ");
            strSql.AppendLine("             AND A.J_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("             AND A.J_DATE <= '" + sYmdTo + "' ");
            strSql.AppendLine("             AND A.J_LOTNO = '1' ");
            strSql.AppendLine("             AND B.DAEGUBUN = '고철B' ");
            strSql.AppendLine("             AND D.EMP_NM <> '김홍철' ");
            strSql.AppendLine("           GROUP BY D.EMP_NM, B.GUBUN1 WITH ROLLUP  ");
            strSql.AppendLine("           ) D ");
            strSql.AppendLine("  ORDER BY EMP_NM = '전체', D.GRADE_CD, EMP_NM, GUBUN1 = '합계' ");

            DataTable dtLight = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            Cursor = Cursors.Default;

            double dSubMaechul = 0;
            double dTotMaechul = 0;
            for (int i = 0; i < dtLight.Rows.Count; i++) //데이터들 중 합계데이터를 쿼리에서 뽑아내지 못하여 C#안에서 처리
            {
                Cursor = Cursors.WaitCursor;

                string sEmpNm = dtLight.Rows[i]["EMP_NM"]?.ToString();
                string sGubun1 = dtLight.Rows[i]["GUBUN1"]?.ToString();

                if (sEmpNm.Equals("전체")) //마지막 ROW의 MAECHUL 셀 
                {
                    dtLight.Rows[i]["MAECHUL"] = dTotMaechul;
                }
                else if (sGubun1.Equals("합계")) //중간 소계 ROW의 MAECHUL 셀
                {
                    dTotMaechul += dSubMaechul;
                    dtLight.Rows[i]["MAECHUL"] = dSubMaechul;
                    dSubMaechul = 0;
                }
                else
                {
                    double dIKongkep = dtLight.Rows[i]["IKONGKEP"] == null ? 0 : Convert.ToDouble(dtLight.Rows[i]["IKONGKEP"]);
                    double dCKongkep = dtLight.Rows[i]["CKONGKEP"] == null ? 0 : Convert.ToDouble(dtLight.Rows[i]["CKONGKEP"]);
                    double dDanjung = dtLight.Rows[i]["DANJUNG"] == null ? 0 : Convert.ToDouble(dtLight.Rows[i]["DANJUNG"]);

                    double dResultDanga = Math.Round(((dIKongkep + dCKongkep) / dDanjung), 1);
                    dtLight.Rows[i]["DDANGA"] = dResultDanga;

                    double dResultMajin = Math.Round((dLtDanga - dResultDanga - dLtMajin), 1);
                    dtLight.Rows[i]["LIGHTMAJIN"] = dResultMajin;

                    double dResultMaechul = Math.Round((dDanjung * dResultMajin), 1);
                    dtLight.Rows[i]["MAECHUL"] = dResultMaechul;

                    dSubMaechul += dResultMaechul;
                }
                Cursor = Cursors.Default;
            }
            GridLight.DataSource = dtLight;

            Cursor = Cursors.WaitCursor;

            //직납 자료 
            double dDrInsentive = TxtIncentive.EditValue == null ? 0 : Convert.ToDouble(TxtIncentive.EditValue);
            double dDrBill = TxtDiscount.EditValue == null ? 0 : Convert.ToDouble(TxtDiscount.EditValue);

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT *  ");
            strSql.AppendLine("   FROM( ");
            strSql.AppendLine(" 		SELECT COALESCE(D.EMP_NM,'전체') AS EMP_NM ");
            strSql.AppendLine(" 		     , SUM(A.DANJUNG) AS DANJUNG ");
            strSql.AppendLine(" 		     , SUM(A.KONGKEP) AS KONGKEP ");
            strSql.AppendLine(" 		     , SUM(A.MIKONGKEP) AS MIKONGKEP ");
            strSql.AppendLine(" 		     , SUM(A.CKONGKEP) AS CKONGKEP ");
            strSql.AppendLine(" 		     , SUM(F.J_KONGKEP) AS J_KONGKEP");
            strSql.AppendLine(" 		     , 0 AS MEACHUL ");
            strSql.AppendLine(" 		     , 0 AS ARRV_PURC_PROFIT ");
            strSql.AppendLine(" 		     , 0 AS SALE_PROFIT  ");
            strSql.AppendLine(" 		     , D.GRADE_CD ");
            strSql.AppendLine(" 		  FROM INLIST A ");
            strSql.AppendLine(" 		  LEFT OUTER JOIN JAJAE B ");
            strSql.AppendLine(" 		    ON B.J_SERIAL = A.J_SERIAL ");
            strSql.AppendLine(" 		  LEFT OUTER JOIN ACC_DEALER_CD C ");
            strSql.AppendLine(" 		    ON C.DEALER_CD = A.J_ID1 ");
            strSql.AppendLine(" 		  LEFT OUTER JOIN HR_EMP_BASIS D ");
            strSql.AppendLine(" 		    ON D.EMP_ID = C.CHRG_ID ");
            strSql.AppendLine(" 		  LEFT OUTER JOIN IPCHULGO F ");
            strSql.AppendLine(" 		    ON F.J_ID - 1000 = A.J_ID ");
            strSql.AppendLine(" 		 WHERE A.KERATYPE  =   '매입'  ");
            strSql.AppendLine(" 		   AND A.J_DATE >= '" + sYmdFrom + "'  ");
            strSql.AppendLine(" 		   AND A.J_DATE <= '" + sYmdTo + "'  ");
            strSql.AppendLine(" 		   AND A.J_LOTNO <>   '1' ");
            strSql.AppendLine(" 		   AND D.EMP_NM <> '기타'  ");
            strSql.AppendLine(" 		   AND D.EMP_NM <> '김홍철' ");
            strSql.AppendLine(" 		 GROUP BY D.EMP_NM WITH ROLLUP ");
            strSql.AppendLine("  	  ) PD ");
            strSql.AppendLine("   ORDER BY EMP_NM = '전체', GRADE_CD, EMP_NM ");

            DataTable dtDr = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            for (int i = 0; i < dtDr.Rows.Count; i++)
            {
                long lJKongkep = dtDr.Rows[i]["J_KONGKEP"] == null ? 0 : Convert.ToInt64(dtDr.Rows[i]["J_KONGKEP"]);
                long IDanjung = dtDr.Rows[i]["DANJUNG"] == null ? 0 : Convert.ToInt64(dtDr.Rows[i]["DANJUNG"]);

                //매출금액(인센티브 포함)
                long Meachul = (lJKongkep + IDanjung * Convert.ToInt64(dDrInsentive))
                                    - ((lJKongkep + IDanjung * Convert.ToInt64(dDrInsentive))
                                    * (Convert.ToInt64(dDrBill) / 100) / 365 * 91);
                dtDr.Rows[i]["MEACHUL"] = Meachul;

                //도착도 매입금액
                long MiKongkep = dtDr.Rows[i]["MIKONGKEP"] == null ? 0 : Convert.ToInt64(dtDr.Rows[i]["MIKONGKEP"]);
                long CKongkep = dtDr.Rows[i]["CKONGKEP"] == null ? 0 : Convert.ToInt64(dtDr.Rows[i]["CKONGKEP"]);

                long ArrvPurcProfit = MiKongkep + CKongkep;
                dtDr.Rows[i]["ARRV_PURC_PROFIT"] = ArrvPurcProfit;

                //매출이익
                //double TmpMeachulProfit = Convert.ToDouble(dtjik.Rows[i]["sum(f.j_kongkep)"]) + TmpMeachul - TmpDdanga;
                long TmpMeachulProfit = Meachul - ArrvPurcProfit;
                dtDr.Rows[i]["SALE_PROFIT"] = TmpMeachulProfit;
            }

            GridDrSend.DataSource = dtDr;

            Cursor = Cursors.Default;

            // 슈레더 매출평균단가 구하기
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT SUM(A.DANJUNG) ");
            strSql.AppendLine("      , SUM(A.KONGKEP) ");
            strSql.AppendLine("      , SUM(A.CKONGKEP) ");
            strSql.AppendLine("      , SUM(A.HALIN) ");
            strSql.AppendLine("   FROM INLIST A ");
            strSql.AppendLine("   LEFT OUTER JOIN JAJAE B  ");
            strSql.AppendLine("     ON B.J_SERIAL = A.J_SERIAL ");
            strSql.AppendLine("  WHERE A.J_LOTNO = '1' ");
            strSql.AppendLine("    AND A.J_ID1 = '3018' ");
            strSql.AppendLine("    AND A.J_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("    AND A.J_DATE <= '" + sYmdTo + "' ");
            strSql.AppendLine("    AND B.DAEGUBUN = '슈레더' ");
            strSql.AppendLine("    AND B.GUBUN1 <> '인센티브' ");

            DataTable dtShreder = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            string sSrdDanjung = dtShreder.Rows[0]?["SUM(A.DANJUNG)"]?.ToString();
            string sMenuSd = TxtShreder.EditValue?.ToString();

            long Money = 0;
            if (string.IsNullOrEmpty(sSrdDanjung) && string.IsNullOrEmpty(sMenuSd))
            {
                XtraMessageBox.Show("기간내 슈레더 매출이 없습니다. 적용단가를 입력하고 조회해 주세요.");
                Cursor = Cursors.Default;
                return;
            }
            else if (!string.IsNullOrEmpty(sSrdDanjung))
            {
                string sKongkep = dtShreder.Rows[0]?["SUM(A.KONGKEP)"]?.ToString();
                string sCKongkep = dtShreder.Rows[0]?["SUM(A.CKONGKEP)"]?.ToString();
                string incentive = TxtIncentive.EditValue == null ? "0" : TxtIncentive.EditValue?.ToString();
                string sBill = TxtDiscount.EditValue == null ? "0" : TxtDiscount.EditValue?.ToString();

                Money = (Convert.ToInt64(sKongkep) - Convert.ToInt64(sCKongkep));
                TxtSrdSaleUtPrc.EditValue = ((Money + (Convert.ToInt64(sSrdDanjung) * Convert.ToInt64(sMenuSd)))
                                        - ((Money + (Convert.ToInt64(sSrdDanjung)) * Convert.ToInt64(incentive))
                                        * ((Convert.ToInt64(sBill) / 100) / 365 * 91)))
                                        / Convert.ToInt64(sSrdDanjung);
            }
            else
            {
                TxtSrdSaleUtPrc.EditValue = TxtShreder.EditValue;
            }

            long sdUtPrc = TxtSrdSaleUtPrc.EditValue == null ? 0 : Convert.ToInt64(TxtSrdSaleUtPrc.EditValue);
            long sdWaste = TxtSrdDiscount.EditValue == null ? 0 : Convert.ToInt64(TxtSrdDiscount.EditValue);

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT *  ");
            strSql.AppendLine("   FROM (  ");
            strSql.AppendLine("          SELECT COALESCE(D.EMP_NM,'전체') AS EMP_NM  ");
            strSql.AppendLine("               , COALESCE(B.GUBUN1,'합계') AS GUBUN1 ");
            strSql.AppendLine("               , SUM(A.DANJUNG) AS DANJUNG   ");
            strSql.AppendLine("               , SUM(A.HALIN) AS HALIN   ");
            strSql.AppendLine("               , SUM(A.IKONGKEP) AS IKONGKEP   ");
            strSql.AppendLine("               , SUM(A.CKONGKEP) AS CKONGKEP   ");
            strSql.AppendLine("               , 0.0 AS DANGA ");
            strSql.AppendLine("               , 0.0 AS DAMONT   ");
            strSql.AppendLine("               , 0 AS SDMEACHUL  ");
            strSql.AppendLine("               , D.GRADE_CD  ");
            strSql.AppendLine("            FROM INLIST A  ");
            strSql.AppendLine("            LEFT OUTER JOIN JAJAE B  ");
            strSql.AppendLine("              ON B.J_SERIAL = A.J_SERIAL  ");
            strSql.AppendLine("            LEFT OUTER JOIN ACC_DEALER_CD C  ");
            strSql.AppendLine("              ON C.DEALER_CD = A.J_ID1  ");
            strSql.AppendLine("            LEFT OUTER JOIN HR_EMP_BASIS D  ");
            strSql.AppendLine("              ON D.EMP_ID = C.CHRG_ID  ");
            strSql.AppendLine("           WHERE A.KERATYPE = '매입'  ");
            strSql.AppendLine("             AND A.J_DATE >= '" + sYmdFrom + "' ");
            strSql.AppendLine("             AND A.J_DATE <= '" + sYmdTo + "' ");
            strSql.AppendLine("             AND A.J_LOTNO = '1'  ");
            strSql.AppendLine("             AND B.DAEGUBUN = '슈레더'  ");
            strSql.AppendLine("             AND D.EMP_NM <> '김홍철'  ");
            strSql.AppendLine("           GROUP BY D.EMP_NM, B.GUBUN1 WITH ROLLUP   ");
            strSql.AppendLine("           ) D  ");
            strSql.AppendLine("  ORDER BY EMP_NM = '전체', D.GRADE_CD, EMP_NM, GUBUN1 = '합계'  ");

            DataTable dtShreder2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridShreder.DataSource = dtShreder2;


        }

        private void GridViewHeavy_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                if (e.RowHandle % 2 == 0)
                {
                    e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
                }

                string sEmpNm = GridViewHeavy.GetRowCellValue(e.RowHandle, "EMP_NM")?.ToString();
                if (!string.IsNullOrEmpty(sEmpNm) && sEmpNm.Equals("전체"))
                {
                    e.Appearance.BackColor = Color.PaleTurquoise;
                }
            }
        }
        
        private void GridViewLight_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if(e.RowHandle >= 0)
            {
                if (e.RowHandle % 2 == 0)
                {
                    e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
                }
                string sGubun = GridViewLight.GetRowCellValue(e.RowHandle, "GUBUN1")?.ToString();
                string sEmpNm = GridViewLight.GetRowCellValue(e.RowHandle, "EMP_NM")?.ToString();

                if (!string.IsNullOrEmpty(sGubun) && sGubun.Equals("합계"))
                {
                    e.Appearance.BackColor = Color.Moccasin;
                }

                if (!string.IsNullOrEmpty(sEmpNm) && sEmpNm.Equals("전체"))
                {
                    e.Appearance.BackColor = Color.PaleTurquoise;
                }
            }
        }

        private void GridViewDrSend_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                if (e.RowHandle % 2 == 0)
                {
                    e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
                }

                string sEmpNm = GridViewHeavy.GetRowCellValue(e.RowHandle, "EMP_NM")?.ToString();
                if (!string.IsNullOrEmpty(sEmpNm) && sEmpNm.Equals("전체"))
                {
                    e.Appearance.BackColor = Color.PaleTurquoise;
                }
            }
        }

        private void GridViewShreder_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle >= 0)
            {
                if (e.RowHandle % 2 == 0)
                {
                    e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
                }
                string sGubun = GridViewLight.GetRowCellValue(e.RowHandle, "GUBUN1")?.ToString();
                string sEmpNm = GridViewLight.GetRowCellValue(e.RowHandle, "EMP_NM")?.ToString();

                if (!string.IsNullOrEmpty(sGubun) && sGubun.Equals("합계"))
                {
                    e.Appearance.BackColor = Color.Moccasin;
                }

                if (!string.IsNullOrEmpty(sEmpNm) && sEmpNm.Equals("전체"))
                {
                    e.Appearance.BackColor = Color.PaleTurquoise;
                }
            }
        }
    }
}