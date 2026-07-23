using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ComLib;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid;
using DevExpress.XtraEditors;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class LstSaleResult : DevExpress.XtraEditors.XtraForm
    {
        double tmpDMoney;
        double tmpProd;
        GridColumnSummaryItem item1 = new GridColumnSummaryItem();
        GridColumnSummaryItem item2 = new GridColumnSummaryItem();
        GridColumnSummaryItem item3 = new GridColumnSummaryItem();
        GridColumnSummaryItem itemSd1 = new GridColumnSummaryItem();
        GridColumnSummaryItem itemSd2 = new GridColumnSummaryItem();
        GridColumnSummaryItem itemSd3 = new GridColumnSummaryItem();

        public LstSaleResult()
        {
            DBConn.DbConn();
            InitializeComponent();
            InitValue();
        }

        public DataRow rowUserInfo { get; set; }
        private void LstSaleResult_Load(object sender, EventArgs e)
        {
            
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "AccDealerCdDev");
        }

        private void BtnSerch_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;

            }

            Cursor = Cursors.WaitCursor;
            StringBuilder strSql = new StringBuilder();
            string sYmdFrom = DeditStart.EditValue.ToString();
            string sYmdTo = DeditEnd.EditValue.ToString();

            // 중량 스크랩 매출평균단가 구하기
            strSql.Clear();

            strSql.AppendLine("SELECT   sum(a.danjung), sum(a.kongkep), sum(a.ckongkep), sum(a.halin) ");
            strSql.AppendLine("FROM     inlist a, jajae b ");
            strSql.AppendLine("WHERE    a.j_serial  = b.j_serial ");
            strSql.AppendLine("AND      a.keratype  =   '매출' ");
            strSql.AppendLine("AND      a.j_date    >=  '" + sYmdFrom + "' ");
            strSql.AppendLine("AND      a.j_date    <=  '" + sYmdTo + "' ");
            strSql.AppendLine("AND      a.j_lotno   =   '1' ");
            strSql.AppendLine("AND      a.j_id1     =   '3018' ");
            strSql.AppendLine("AND      b.daegubun  =   '고철A' ");
            strSql.AppendLine("AND      b.Gubun1    <>  '인센티브' ");

            DataTable tempGsHeavy = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (String.IsNullOrEmpty(tempGsHeavy.Rows[0]["sum(a.danjung)"].ToString()) && String.IsNullOrEmpty(TeditHeavy.EditValue.ToString()))
            {
                MessageBox.Show("기간내 스크랩 중량 매출이 없습니다. 적용단가를 입력하고 조회해 주세요.");
                Cursor = Cursors.Default;
                return;
            }
            else if(!String.IsNullOrEmpty(tempGsHeavy.Rows[0]["sum(a.danjung)"].ToString()))
            {

                tmpDMoney = (Convert.ToDouble(tempGsHeavy.Rows[0]["sum(a.kongkep)"]) - Convert.ToDouble(tempGsHeavy.Rows[0]["sum(a.ckongkep)"]));
                TeditMeachul.EditValue = (tmpDMoney - (tmpDMoney * (Convert.ToDouble(TeditBill.EditValue) / 100) / 365 * 91))
                                       / Convert.ToDouble(tempGsHeavy.Rows[0]["sum(a.danjung)"]);
                TeditDanga.EditValue = ((tmpDMoney + (Convert.ToDouble(tempGsHeavy.Rows[0]["sum(a.danjung)"]) * Convert.ToDouble(TeditInsentiv.EditValue)))
                                    - ((tmpDMoney + (Convert.ToDouble(tempGsHeavy.Rows[0]["sum(a.danjung)"]) * Convert.ToDouble(TeditInsentiv.EditValue)))
                                    * (Convert.ToDouble(TeditBill.EditValue) / 100) / 365 * 91)) / Convert.ToDouble(tempGsHeavy.Rows[0]["sum(a.danjung)"]);
                
            }
            else
            {
                TeditDanga.EditValue = TeditHeavy.EditValue;
            }
            strSql.Clear();

            // 중량 스크랩 자료 뿌리기
            double Danga = Convert.ToDouble(TeditDanga.EditValue); //적용매출단가
            double Majin = Convert.ToDouble(TeditMajin.EditValue); //스크랩중량마진

            strSql.AppendLine("SELECT   e.name, sum(a.danjung), sum(a.halin), sum(a.ikongkep), sum(a.ckongkep), ");
            strSql.AppendLine("         0.00 as Ddanga, 0.00 as HeavyMajin, 0 as Meachul");
            strSql.AppendLine("FROM     inlist a, jajae b, custom d, sawon e ");
            strSql.AppendLine("WHERE    a.j_serial  = b.j_serial ");
            strSql.AppendLine("AND      a.J_ID1      =  d.C_Serial ");
            strSql.AppendLine("AND      d.Damdang   = e.S_Serial ");
            strSql.AppendLine("AND      a.keratype  =   '매입' ");
            strSql.AppendLine("AND      a.j_date    >=  '" + sYmdFrom + "' ");
            strSql.AppendLine("AND      a.j_date    <=  '" + sYmdTo + "' ");
            strSql.AppendLine("AND      a.j_lotno   =   '1' ");
            strSql.AppendLine("AND      b.daegubun  =   '고철A' ");
            strSql.AppendLine("AND      e.name <> '김홍철' ");
            strSql.AppendLine("AND      e.name <> '기타' ");
            strSql.AppendLine("group by e.name ");
            strSql.AppendLine("order by binary(e.jikep) asc");

            DataTable dtHeavy = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            gridControl1.DataSource = dtHeavy;
            for (int i = 0; i < dtHeavy.Rows.Count; i++)
            {
                //도착매입단가
                double TmpDanga = (Convert.ToDouble(dtHeavy.Rows[i]["sum(a.ikongkep)"]) + Convert.ToDouble(dtHeavy.Rows[i]["sum(a.ckongkep)"]))
                                    / Convert.ToDouble(dtHeavy.Rows[i]["sum(a.danjung)"]);

                gridView1.SetRowCellValue(i, "Ddanga", TmpDanga);
                //스크랩 중량 마진
                double TmpMajin = Danga - TmpDanga - Majin;
                gridView1.SetRowCellValue(i, "HeavyMajin", TmpMajin);
                //매출이익
                double TmpMeachul = Convert.ToDouble(dtHeavy.Rows[i]["sum(a.danjung)"]) * TmpMajin;
                gridView1.SetRowCellValue(i, "Meachul", TmpMeachul);
            }
            gridView1.RefreshData();
            strSql.Clear();

            // 경량 스크랩 매출평균단가 구하기
            strSql.AppendLine("SELECT   sum(a.danjung), sum(a.kongkep), sum(a.ckongkep), sum(a.halin) ");
            strSql.AppendLine("FROM     inlist a, jajae b ");
            strSql.AppendLine("WHERE    a.j_serial  = b.j_serial ");
            strSql.AppendLine("AND      a.keratype  =   '매출' ");
            strSql.AppendLine("AND      a.j_date    >=  '" + sYmdFrom + "' ");
            strSql.AppendLine("AND      a.j_date    <=  '" + sYmdTo + "' ");
            strSql.AppendLine("AND      a.j_lotno   =   '1' ");
            strSql.AppendLine("AND      a.j_id1     =   '3018' ");
            strSql.AppendLine("AND      b.daegubun  =   '고철B' ");
            strSql.AppendLine("AND      b.Gubun1    <>  '인센티브' ");

            DataTable tempGsLight = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (String.IsNullOrEmpty(tempGsLight.Rows[0]["sum(a.danjung)"].ToString()) && String.IsNullOrEmpty(TeditLight.EditValue.ToString()))
            {
                MessageBox.Show("기간내 스크랩 경량 매출이 없습니다. 적용단가를 입력하고 조회해 주세요.");
                Cursor = Cursors.Default;
                return;
            }
            else if (!String.IsNullOrEmpty(tempGsLight.Rows[0]["sum(a.danjung)"].ToString()))
            {

                tmpDMoney = (Convert.ToDouble(tempGsLight.Rows[0]["sum(a.kongkep)"]) - Convert.ToDouble(tempGsLight.Rows[0]["sum(a.ckongkep)"]));
                TeditMeachul1.EditValue = (tmpDMoney - (tmpDMoney * (Convert.ToDouble(TeditBill.EditValue) / 100) / 365 * 91))
                                       / Convert.ToDouble(tempGsLight.Rows[0]["sum(a.danjung)"]);
                TeditDanga1.EditValue = ((tmpDMoney + (Convert.ToDouble(tempGsLight.Rows[0]["sum(a.danjung)"]) * Convert.ToDouble(TeditInsentiv.EditValue)))
                                    - ((tmpDMoney + (Convert.ToDouble(tempGsLight.Rows[0]["sum(a.danjung)"]) * Convert.ToDouble(TeditInsentiv.EditValue)))
                                    * (Convert.ToDouble(TeditBill.EditValue) / 100) / 365 * 91)) / Convert.ToDouble(tempGsLight.Rows[0]["sum(a.danjung)"]);

            }
            else
            {
                TeditDanga1.EditValue = TeditLight.EditValue;
            }

            strSql.Clear();

            double Danga1 = Convert.ToDouble(TeditDanga1.EditValue); //적용매출단가
            double Majin1 = Convert.ToDouble(TeditMajin1.EditValue); //스크랩경량마진

            // 경량 스크랩 자료 뿌리기 - 손상영
            strSql.AppendLine("SELECT   e.name, sum(a.danjung), sum(a.halin), sum(a.ikongkep), sum(a.ckongkep), b.gubun1, ");
            strSql.AppendLine("         0.0 as Ddanga, 0.0 as LightMajin, 0 as Meachul");
            strSql.AppendLine("FROM     inlist a, jajae b, custom d, sawon e ");
            strSql.AppendLine("WHERE    a.j_serial  = b.j_serial ");
            strSql.AppendLine("AND      a.J_ID1      =  d.C_Serial ");
            strSql.AppendLine("AND      d.Damdang   = e.S_Serial ");
            strSql.AppendLine("AND      a.keratype  =   '매입' ");
            strSql.AppendLine("AND      a.j_date    >=  '" + sYmdFrom + "' ");
            strSql.AppendLine("AND      a.j_date    <=  '" + sYmdTo + "' ");
            strSql.AppendLine("AND      a.j_lotno   =   '1' ");
            strSql.AppendLine("AND      b.daegubun  =   '고철B' ");
            strSql.AppendLine("AND      e.name = '손상영' ");
            strSql.AppendLine("group by b.gubun1 ");
            strSql.AppendLine("order by binary(e.jikep) asc");

            DataTable dtLight1 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            gridControl3.DataSource = dtLight1;
            for (int i = 0; i < dtLight1.Rows.Count; i++)
            {
                //도착매입단가
                double TmpDanga = (Convert.ToDouble(dtLight1.Rows[i]["sum(a.ikongkep)"]) + Convert.ToDouble(dtLight1.Rows[i]["sum(a.ckongkep)"]))
                                    / Convert.ToDouble(dtLight1.Rows[i]["sum(a.danjung)"]);
                gridView3.SetRowCellValue(i, "Ddanga", TmpDanga);
                //스크랩 경량 마진
                double TmpMajin = Danga1 - TmpDanga - Majin1;
                gridView3.SetRowCellValue(i, "LightMajin", TmpMajin);
                //매출이익
                double TmpMeachul = Convert.ToDouble(dtLight1.Rows[i]["sum(a.danjung)"]) * TmpMajin;
                gridView3.SetRowCellValue(i, "Meachul", TmpMeachul);
            }
            gridView3.RefreshData();
            strSql.Clear();

            // 경량 스크랩 자료 뿌리기 - 오상훈
            strSql.AppendLine("SELECT   e.name, sum(a.danjung), sum(a.halin), sum(a.ikongkep), sum(a.ckongkep), b.gubun1, ");
            strSql.AppendLine("         0.00 as Ddanga, 0.0 as LightMajin, 0 as Meachul");
            strSql.AppendLine("FROM     inlist a, jajae b, custom d, sawon e ");
            strSql.AppendLine("WHERE    a.j_serial  = b.j_serial ");
            strSql.AppendLine("AND      a.J_ID1      =  d.C_Serial ");
            strSql.AppendLine("AND      d.Damdang   = e.S_Serial ");
            strSql.AppendLine("AND      a.keratype  =   '매입' ");
            strSql.AppendLine("AND      a.j_date    >=  '" + sYmdFrom + "' ");
            strSql.AppendLine("AND      a.j_date    <=  '" + sYmdTo + "' ");
            strSql.AppendLine("AND      a.j_lotno   =   '1' ");
            strSql.AppendLine("AND      b.daegubun  =   '고철B' ");
            strSql.AppendLine("AND      e.name = '오상훈' ");
            strSql.AppendLine("group by b.gubun1 ");
            strSql.AppendLine("order by binary(e.jikep) asc");

            DataTable dtLight2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            gridControl4.DataSource = dtLight2;
            for (int i = 0; i < dtLight2.Rows.Count; i++)
            {
                //도착매입단가
                double TmpDanga = (Convert.ToDouble(dtLight2.Rows[i]["sum(a.ikongkep)"]) + Convert.ToDouble(dtLight2.Rows[i]["sum(a.ckongkep)"]))
                                    / Convert.ToDouble(dtLight2.Rows[i]["sum(a.danjung)"]);
                gridView4.SetRowCellValue(i, "Ddanga", TmpDanga);
                //스크랩 경량 마진
                double TmpMajin = Danga1 - TmpDanga - Majin1;
                gridView4.SetRowCellValue(i, "LightMajin", TmpMajin);
                //매출이익
                double TmpMeachul = Convert.ToDouble(dtLight2.Rows[i]["sum(a.danjung)"]) * TmpMajin;
                gridView4.SetRowCellValue(i, "Meachul", TmpMeachul);
            }
            gridView4.RefreshData();
            strSql.Clear();

            // 경량 스크랩 자료 뿌리기 - 김명철
            strSql.AppendLine("SELECT   e.name, sum(a.danjung), sum(a.halin), sum(a.ikongkep), sum(a.ckongkep), b.gubun1, ");
            strSql.AppendLine("         0.0 as Ddanga, 0.0 as LightMajin, 0 as Meachul");
            strSql.AppendLine("FROM     inlist a, jajae b, custom d, sawon e ");
            strSql.AppendLine("WHERE    a.j_serial  = b.j_serial ");
            strSql.AppendLine("AND      a.J_ID1      =  d.C_Serial ");
            strSql.AppendLine("AND      d.Damdang   = e.S_Serial ");
            strSql.AppendLine("AND      a.keratype  =   '매입' ");
            strSql.AppendLine("AND      a.j_date    >=  '" + sYmdFrom + "' ");
            strSql.AppendLine("AND      a.j_date    <=  '" + sYmdTo + "' ");
            strSql.AppendLine("AND      a.j_lotno   =   '1' ");
            strSql.AppendLine("AND      b.daegubun  =   '고철B' ");
            strSql.AppendLine("AND      e.name = '김명철' ");
            strSql.AppendLine("group by b.gubun1 ");
            strSql.AppendLine("order by binary(e.jikep) asc");

            DataTable dtLight3 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            gridControl5.DataSource = dtLight3;
            for (int i = 0; i < dtLight3.Rows.Count; i++)
            {
                //도착매입단가
                double TmpDanga = (Convert.ToDouble(dtLight3.Rows[i]["sum(a.ikongkep)"]) + Convert.ToDouble(dtLight3.Rows[i]["sum(a.ckongkep)"]))
                                    / Convert.ToDouble(dtLight3.Rows[i]["sum(a.danjung)"]);
                gridView5.SetRowCellValue(i, "Ddanga", TmpDanga);
                //스크랩 경량 마진
                double TmpMajin = Danga1 - TmpDanga - Majin1;
                gridView5.SetRowCellValue(i, "LightMajin", TmpMajin);
                //매출이익
                double TmpMeachul = Convert.ToDouble(dtLight3.Rows[i]["sum(a.danjung)"]) * TmpMajin;
                gridView5.SetRowCellValue(i, "Meachul", TmpMeachul);
            }
            gridView5.RefreshData();
            strSql.Clear();

            // 직납 자료 뿌리기 
            double Insentive = Convert.ToDouble(TeditInsentiv.EditValue);
            double bill = Convert.ToDouble(TeditBill.EditValue);

            //매입금액:mikongkep  매출금액:j_kongkep
            strSql.AppendLine("SELECT   e.name, sum(a.danjung), sum(a.kongkep), sum(a.mikongkep), sum(a.ckongkep), sum(f.j_kongkep), ");
            strSql.AppendLine("         0 as Meachul, 0 as Ddanga, 0 as MeachulProfit "); 
            strSql.AppendLine("FROM     inlist a, jajae b, custom d, sawon e, ipchulgo f ");
            strSql.AppendLine("WHERE    a.j_serial  = b.j_serial ");
            strSql.AppendLine("AND      a.J_ID1      =  d.C_Serial ");
            strSql.AppendLine("AND      d.Damdang   = e.S_Serial ");
            strSql.AppendLine("AND      a.j_id      = f.j_id - 1000 ");
            strSql.AppendLine("AND      a.keratype  =   '매입' ");
            strSql.AppendLine("AND      a.j_date    >=  '" + sYmdFrom + "' ");
            strSql.AppendLine("AND      a.j_date    <=  '" + sYmdTo + "' ");
            strSql.AppendLine("AND      a.j_lotno   <>   '1' ");
            strSql.AppendLine("AND      e.name <> '기타' ");
            strSql.AppendLine("group by e.name ");
            strSql.AppendLine("order by binary(e.jikep) asc");

            DataTable dtjik = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            gridControl2.DataSource = dtjik;
            for (int i = 0; i < dtjik.Rows.Count; i++)
            {
                //매출금액(인센티브 포함)
                double TmpMeachul = (Convert.ToDouble(dtjik.Rows[i]["sum(f.j_kongkep)"]) + (Convert.ToDouble(dtjik.Rows[i]["sum(a.danjung)"]) * Insentive))
                                    - ((Convert.ToDouble(dtjik.Rows[i]["sum(f.j_kongkep)"]) + (Convert.ToDouble(dtjik.Rows[i]["sum(a.danjung)"]) * Insentive))
                                    * (bill / 100) / 365 * 91);
                gridView2.SetRowCellValue(i, "Meachul", TmpMeachul);
                //도착도 매입금액
                double TmpDdanga = Convert.ToDouble(dtjik.Rows[i]["sum(a.mikongkep)"]) + Convert.ToDouble(dtjik.Rows[i]["sum(a.ckongkep)"]);
                gridView2.SetRowCellValue(i, "Ddanga", TmpDdanga);
                //매출이익
                //double TmpMeachulProfit = Convert.ToDouble(dtjik.Rows[i]["sum(f.j_kongkep)"]) + TmpMeachul - TmpDdanga;
                double TmpMeachulProfit = TmpMeachul - TmpDdanga;
                gridView2.SetRowCellValue(i, "MeachulProfit", TmpMeachulProfit);
            }
            gridView2.RefreshData();
            strSql.Clear();

            // 슈레더 매출평균단가 구하기
            strSql.AppendLine("SELECT   sum(a.danjung), sum(a.kongkep), sum(a.ckongkep), sum(a.halin) ");
            strSql.AppendLine("FROM     inlist a, jajae b ");
            strSql.AppendLine("WHERE    a.j_serial  = b.j_serial ");
            strSql.AppendLine("AND      a.keratype  =   '매출' ");
            strSql.AppendLine("AND      a.j_date    >=  '" + sYmdFrom + "' ");
            strSql.AppendLine("AND      a.j_date    <=  '" + sYmdTo + "' ");
            strSql.AppendLine("AND      a.j_lotno   =   '1' ");
            strSql.AppendLine("AND      a.j_id1     =   '3018' ");
            strSql.AppendLine("AND      b.daegubun  =   '슈레더' ");
            strSql.AppendLine("AND      b.Gubun1    <>  '인센티브' ");

            DataTable tempSd = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (String.IsNullOrEmpty(tempSd.Rows[0]["sum(a.danjung)"].ToString()) && String.IsNullOrEmpty(TeditSd.EditValue.ToString()))
            {
                MessageBox.Show("기간내 슈레더 매출이 없습니다. 적용단가를 입력하고 조회해 주세요.");
                Cursor = Cursors.Default;
                return;
            }
            else if (!String.IsNullOrEmpty(tempSd.Rows[0]["sum(a.danjung)"].ToString()))
            {
                tmpDMoney = (Convert.ToDouble(tempSd.Rows[0]["sum(a.kongkep)"]) - Convert.ToDouble(tempSd.Rows[0]["sum(a.ckongkep)"]));
                TeditMeachul2.EditValue = ((tmpDMoney + (Convert.ToDouble(tempSd.Rows[0]["sum(a.danjung)"]) * Convert.ToDouble(TeditInsentiv.EditValue)))
                                        - ((tmpDMoney + (Convert.ToDouble(tempSd.Rows[0]["sum(a.danjung)"]) * Convert.ToDouble(TeditInsentiv.EditValue) ))
                                        * ((Convert.ToDouble(TeditBill.EditValue) / 100) / 365 * 91))) 
                                        / Convert.ToDouble(tempSd.Rows[0]["sum(a.danjung)"]);
            }
            else
            {
                TeditMeachul2.EditValue = TeditSd.EditValue;
            }

            double SdDanga = Convert.ToDouble(TeditMeachul2.EditValue); //매출평균단가
            double Waste = Convert.ToDouble(TeditWaste.EditValue);    //폐기물처리비

            strSql.Clear();

            // 슈레더 자료 뿌리기 - 손상영
            strSql.AppendLine("SELECT   e.name, sum(a.danjung), sum(a.halin), sum(a.ikongkep), sum(a.ckongkep), b.gubun1, 0 as SdMeachul,");
            strSql.AppendLine("         0.0 as Danga, 0 as DAmont ");
            strSql.AppendLine("FROM     inlist a, jajae b, custom d, sawon e ");
            strSql.AppendLine("WHERE    a.j_serial  = b.j_serial ");
            strSql.AppendLine("AND      a.J_ID1      =  d.C_Serial ");
            strSql.AppendLine("AND      d.Damdang   = e.S_Serial ");
            strSql.AppendLine("AND      a.keratype  =   '매입' ");
            strSql.AppendLine("AND      a.j_date    >=  '" + sYmdFrom + "' ");
            strSql.AppendLine("AND      a.j_date    <=  '" + sYmdTo + "' ");
            strSql.AppendLine("AND      a.j_lotno   =   '1' ");
            strSql.AppendLine("AND      b.daegubun  =   '슈레더' ");
            strSql.AppendLine("AND      e.name = '손상영' ");
            strSql.AppendLine("group by b.gubun1 ");
            strSql.AppendLine("order by b.gubun1 desc");

            DataTable dtSD1 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            double Recovery = 0;
            gridControl6.DataSource = dtSD1;

            for (int i = 0; i < dtSD1.Rows.Count; i++)
            {
                Recovery = get_recovery(dtSD1.Rows[i]["gubun1"].ToString());
                //도착매입금액
                double damont = Convert.ToDouble(dtSD1.Rows[i]["sum(a.ikongkep)"]) + Convert.ToDouble(dtSD1.Rows[i]["sum(a.ckongkep)"]);
                gridView6.SetRowCellValue(i, "DAmont", damont);
                //매입평균단가
                double danga = damont / Convert.ToDouble(dtSD1.Rows[i]["sum(a.danjung)"]);
                gridView6.SetRowCellValue(i, "Danga", danga);
                //매출이익
                double sdmeachul = 0;
                if (dtSD1.Rows[i]["gubun1"].ToString().Contains("슈레더"))  //슈레더 등급 매출이익 계산식
                {
                    sdmeachul = (Convert.ToDouble(dtSD1.Rows[i]["sum(a.danjung)"]) * Recovery * SdDanga) - damont;
                }
                else  // 슈레더 등급 이외 매출이익 계산식
                {
                    sdmeachul = (Convert.ToDouble(dtSD1.Rows[i]["sum(a.danjung)"]) * Recovery * SdDanga) - damont
                              - (Convert.ToDouble(dtSD1.Rows[i]["sum(a.danjung)"]) * tmpProd)
                              - (Convert.ToDouble(dtSD1.Rows[i]["sum(a.danjung)"]) * (1 - Recovery) * Waste);
                }
                gridView6.SetRowCellValue(i, "SdMeachul", sdmeachul);
            }
            gridView6.RefreshData();
            strSql.Clear();

            // 슈레더 자료 뿌리기 - 오상훈
            strSql.AppendLine("SELECT   e.name, sum(a.danjung), sum(a.halin), sum(a.ikongkep), sum(a.ckongkep), b.gubun1, 0 as SdMeachul,");
            strSql.AppendLine("         0.0 as Danga, 0 as DAmont ");
            strSql.AppendLine("FROM     inlist a, jajae b, custom d, sawon e ");
            strSql.AppendLine("WHERE    a.j_serial  = b.j_serial ");
            strSql.AppendLine("AND      a.J_ID1      =  d.C_Serial ");
            strSql.AppendLine("AND      d.Damdang   = e.S_Serial ");
            strSql.AppendLine("AND      a.keratype  =   '매입' ");
            strSql.AppendLine("AND      a.j_date    >=  '" + sYmdFrom + "' ");
            strSql.AppendLine("AND      a.j_date    <=  '" + sYmdTo + "' ");
            strSql.AppendLine("AND      a.j_lotno   =   '1' ");
            strSql.AppendLine("AND      b.daegubun  =   '슈레더' ");
            strSql.AppendLine("AND      e.name = '오상훈' ");
            strSql.AppendLine("group by b.gubun1 ");
            strSql.AppendLine("order by b.gubun1 desc");

            DataTable dtSD2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            Recovery = 0;
            gridControl7.DataSource = dtSD2;

            for (int i = 0; i < dtSD2.Rows.Count; i++)
            {
                Recovery = get_recovery(dtSD2.Rows[i]["gubun1"].ToString());
                //도착매입금액
                double damont = Convert.ToDouble(dtSD2.Rows[i]["sum(a.ikongkep)"]) + Convert.ToDouble(dtSD2.Rows[i]["sum(a.ckongkep)"]);
                gridView7.SetRowCellValue(i, "DAmont", damont);
                //매입평균단가
                double danga = damont / Convert.ToDouble(dtSD2.Rows[i]["sum(a.danjung)"]);
                gridView7.SetRowCellValue(i, "Danga", danga);
                //매출이익
                double sdmeachul = 0;
                if (dtSD2.Rows[i]["gubun1"].ToString().Contains("슈레더"))  //슈레더 등급 매출이익 계산식
                {
                    sdmeachul = (Convert.ToDouble(dtSD2.Rows[i]["sum(a.danjung)"]) * Recovery * SdDanga) - damont;
                }
                else  // 슈레더 등급 이외 매출이익 계산식
                {
                    sdmeachul = (Convert.ToDouble(dtSD2.Rows[i]["sum(a.danjung)"]) * Recovery * SdDanga) - damont
                              - (Convert.ToDouble(dtSD2.Rows[i]["sum(a.danjung)"]) * tmpProd)
                              - (Convert.ToDouble(dtSD2.Rows[i]["sum(a.danjung)"]) * (1 - Recovery) * Waste);
                }
                gridView7.SetRowCellValue(i, "SdMeachul", sdmeachul);
                
            }
            gridView7.RefreshData();

            strSql.Clear();
            // 슈레더 자료 뿌리기 - 김명철
            strSql.AppendLine("SELECT   e.name, sum(a.danjung), sum(a.halin), sum(a.ikongkep), sum(a.ckongkep), b.gubun1, 0 as SdMeachul,");
            strSql.AppendLine("         0.0 as Danga, 0 as DAmont ");
            strSql.AppendLine("FROM     inlist a, jajae b, custom d, sawon e ");
            strSql.AppendLine("WHERE    a.j_serial  = b.j_serial ");
            strSql.AppendLine("AND      a.J_ID1      =  d.C_Serial ");
            strSql.AppendLine("AND      d.Damdang   = e.S_Serial ");
            strSql.AppendLine("AND      a.keratype  =   '매입' ");
            strSql.AppendLine("AND      a.j_date    >=  '" + sYmdFrom + "' ");
            strSql.AppendLine("AND      a.j_date    <=  '" + sYmdTo + "' ");
            strSql.AppendLine("AND      a.j_lotno   =   '1' ");
            strSql.AppendLine("AND      b.daegubun  =   '슈레더' ");
            strSql.AppendLine("AND      e.name = '김명철' ");
            strSql.AppendLine("group by b.gubun1 ");
            strSql.AppendLine("order by b.gubun1 desc");

            DataTable dtSD3 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            Recovery = 0;
            gridControl8.DataSource = dtSD3;

            for (int i = 0; i < dtSD3.Rows.Count; i++)
            {
                Recovery = get_recovery(dtSD3.Rows[i]["gubun1"].ToString());
                //도착매입금액
                double damont = Convert.ToDouble(dtSD3.Rows[i]["sum(a.ikongkep)"]) + Convert.ToDouble(dtSD3.Rows[i]["sum(a.ckongkep)"]);
                gridView8.SetRowCellValue(i, "DAmont", damont);
                //매입평균단가
                double danga = damont / Convert.ToDouble(dtSD3.Rows[i]["sum(a.danjung)"]);
                gridView8.SetRowCellValue(i, "Danga", danga);
                //매출이익
                double sdmeachul = 0;
                if (dtSD3.Rows[i]["gubun1"].ToString().Contains("슈레더"))  //슈레더 등급 매출이익 계산식
                {
                    sdmeachul = (Convert.ToDouble(dtSD3.Rows[i]["sum(a.danjung)"]) * Recovery * SdDanga) - damont;
                }
                else  // 슈레더 등급 이외 매출이익 계산식
                {
                    sdmeachul = (Convert.ToDouble(dtSD3.Rows[i]["sum(a.danjung)"]) * Recovery * SdDanga) - damont
                              - (Convert.ToDouble(dtSD3.Rows[i]["sum(a.danjung)"]) * tmpProd)
                              - (Convert.ToDouble(dtSD3.Rows[i]["sum(a.danjung)"]) * (1 - Recovery) * Waste);
                }
                gridView8.SetRowCellValue(i, "SdMeachul", sdmeachul);
            }
            gridView8.RefreshData();
            strSql.Clear();

            Cursor = Cursors.Default;
            simpleButton1.Visible = true;

            //스크랩 경량 총합계 계산
            double SumDanjung   = Convert.ToDouble(gridView3.Columns["sum(a.danjung)"].SummaryItem.SummaryValue.ToString())
                                + Convert.ToDouble(gridView4.Columns["sum(a.danjung)"].SummaryItem.SummaryValue.ToString())
                                + Convert.ToDouble(gridView5.Columns["sum(a.danjung)"].SummaryItem.SummaryValue.ToString());
            double SumHalin     = Convert.ToDouble(gridView3.Columns["sum(a.halin)"].SummaryItem.SummaryValue.ToString())
                                + Convert.ToDouble(gridView4.Columns["sum(a.halin)"].SummaryItem.SummaryValue.ToString())
                                + Convert.ToDouble(gridView5.Columns["sum(a.halin)"].SummaryItem.SummaryValue.ToString());
            double SumMeachul   = Convert.ToDouble(gridView3.Columns["Meachul"].SummaryItem.SummaryValue.ToString())
                                + Convert.ToDouble(gridView4.Columns["Meachul"].SummaryItem.SummaryValue.ToString())
                                + Convert.ToDouble(gridView5.Columns["Meachul"].SummaryItem.SummaryValue.ToString());

            item1.SetSummary(DevExpress.Data.SummaryItemType.Custom, string.Format("{0:0,0}", SumDanjung));
            item2.SetSummary(DevExpress.Data.SummaryItemType.Custom, string.Format("{0:0,0}", SumHalin));
            item3.SetSummary(DevExpress.Data.SummaryItemType.Custom, string.Format("{0:0,0}", SumMeachul));

            gridView5.RefreshData();

            //슈레더 총합계 계산
            double SumSdDanjung = Convert.ToDouble(gridView6.Columns["sum(a.danjung)"].SummaryItem.SummaryValue.ToString())
                                + Convert.ToDouble(gridView7.Columns["sum(a.danjung)"].SummaryItem.SummaryValue.ToString())
                                + Convert.ToDouble(gridView8.Columns["sum(a.danjung)"].SummaryItem.SummaryValue.ToString());
            double SumDAmont    = Convert.ToDouble(gridView6.Columns["DAmont"].SummaryItem.SummaryValue.ToString())
                                + Convert.ToDouble(gridView7.Columns["DAmont"].SummaryItem.SummaryValue.ToString())
                                + Convert.ToDouble(gridView8.Columns["DAmont"].SummaryItem.SummaryValue.ToString());
            double SumSdMeachul = Convert.ToDouble(gridView6.Columns["SdMeachul"].SummaryItem.SummaryValue.ToString())
                                + Convert.ToDouble(gridView7.Columns["SdMeachul"].SummaryItem.SummaryValue.ToString())
                                + Convert.ToDouble(gridView8.Columns["SdMeachul"].SummaryItem.SummaryValue.ToString());
            itemSd1.SetSummary(DevExpress.Data.SummaryItemType.Custom, string.Format("{0:0,0}", SumSdDanjung));
            itemSd2.SetSummary(DevExpress.Data.SummaryItemType.Custom, string.Format("{0:0,0}", SumDAmont));
            itemSd3.SetSummary(DevExpress.Data.SummaryItemType.Custom, string.Format("{0:0,0}", SumSdMeachul));

            gridView8.RefreshData();
        }

        public void InitValue()
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();

            strSql.AppendLine("SELECT   * ");
            strSql.AppendLine("FROM     meaipsiljuk ");
            //strSql.AppendLine("WHERE    daegubun = '기초자료' ");

            DataTable dtInit = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            TeditInsentiv.EditValue = 0;
            TeditBill.EditValue = 0;
            TeditBillInterest.EditValue = 0;
            TeditWaste.EditValue = 0;
            TeditMajin.EditValue = 0;
            TeditMajin1.EditValue = 0;

            for (int i = 0; i < dtInit.Rows.Count; i++)
            {
                switch (dtInit.Rows[i]["gubun"].ToString())
                {
                    case "인센티브":
                        TeditInsentiv.EditValue = Convert.ToDouble(dtInit.Rows[i]["datavalue"].ToString());
                        break;
                    case "어음할인":
                        TeditBill.EditValue = Convert.ToDouble(dtInit.Rows[i]["datavalue"].ToString());
                        break;
                    case "어음이자":
                        TeditBillInterest.EditValue = Convert.ToDouble(dtInit.Rows[i]["datavalue"].ToString());
                        break;
                    case "폐기물":
                        TeditWaste.EditValue = Convert.ToDouble(dtInit.Rows[i]["datavalue"].ToString());
                        break;
                    case "중량마진":
                        TeditMajin.EditValue = Convert.ToDouble(dtInit.Rows[i]["datavalue"].ToString());
                        break;
                    case "경량마진":
                        TeditMajin1.EditValue = Convert.ToDouble(dtInit.Rows[i]["datavalue"].ToString());
                        break;
                    case "생산단가":
                        tmpProd = Convert.ToDouble(dtInit.Rows[i]["datavalue"].ToString());
                        break;
                }
            }
            //총합계 foot 포멧세팅 후 생성
            item1.SetSummary(DevExpress.Data.SummaryItemType.Custom, string.Format("{0:0,0}", ""));
            gridView5.Columns["sum(a.danjung)"].Summary.Add(item1);
            item2.SetSummary(DevExpress.Data.SummaryItemType.Custom, string.Format("{0:0,0}", ""));
            gridView5.Columns["sum(a.halin)"].Summary.Add(item2);
            item3.SetSummary(DevExpress.Data.SummaryItemType.Custom, string.Format("{0:0,0}", ""));
            gridView5.Columns["Meachul"].Summary.Add(item3);

            itemSd1.SetSummary(DevExpress.Data.SummaryItemType.Custom, string.Format("{0:0,0}", ""));
            gridView8.Columns["sum(a.danjung)"].Summary.Add(itemSd1);
            itemSd2.SetSummary(DevExpress.Data.SummaryItemType.Custom, string.Format("{0:0,0}", ""));
            gridView8.Columns["DAmont"].Summary.Add(itemSd2);
            itemSd3.SetSummary(DevExpress.Data.SummaryItemType.Custom, string.Format("{0:0,0}", ""));
            gridView8.Columns["SdMeachul"].Summary.Add(itemSd3);

            //기본 검색일자 설정
            DeditStart.EditValue = DateTime.Today.AddDays(-1);
            DeditEnd.EditValue = DateTime.Today.AddDays(0);
        }

        //슈레더 회수율 검색 결과 리턴
        public double get_recovery(String get_gubun)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();

            strSql.AppendLine("SELECT   recovery ");
            strSql.AppendLine("FROM     meaipsiljuk ");
            strSql.AppendLine("WHERE    gubun = '" + get_gubun + "' ");

            SqlCommand cmd = new SqlCommand(strSql.ToString(), DBConn.dbCon);
            SqlDataReader rdr = cmd.ExecuteReader();

            strSql.Clear();
            rdr.Read();

            double d = Convert.ToDouble(rdr[0]);
            rdr.Close();

            return d;
        }

        //기초자료 탭 
        public SqlDataAdapter Adpt;
        public DataTable dtInitMain;
        DataTable dtInitSdMain;
        private void xtraTabControl1_Click(object sender, EventArgs e)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();

            strSql.AppendLine("SELECT   gubun,datavalue ");
            strSql.AppendLine("FROM     meaipsiljuk ");
            strSql.AppendLine("Where    daegubun   <>  '슈레더' ");

            SqlCommand cmd = new SqlCommand(strSql.ToString(), DBConn.dbCon);
            dtInitMain = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            gridControl9.DataSource = dtInitMain;
            dataNavigator1.DataSource = dtInitMain;

            strSql.Clear();
           
            strSql.AppendLine("SELECT   gubun,recovery ");
            strSql.AppendLine("FROM     meaipsiljuk ");
            strSql.AppendLine("Where    daegubun   =  '슈레더' ");

            dtInitSdMain = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            gridControl10.DataSource = dtInitSdMain;

        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;

            }
            SqlCommand cmd = new SqlCommand();
            StringBuilder strSql = new StringBuilder();
            string tmpGubun;
            double tmpValue;
            strSql.Clear();
            for (int i = 0; i < dtInitMain.Rows.Count; i++)
            {
                tmpGubun = dtInitMain.Rows[i]["gubun"].ToString();
                tmpValue = Convert.ToDouble(dtInitMain.Rows[i]["datavalue"].ToString());
                strSql.AppendLine("Update meaipsiljuk set datavalue = '" + tmpValue + "' ");
                strSql.AppendLine("Where  gubun = '" + tmpGubun + "' ");
                cmd = new SqlCommand(strSql.ToString(), DBConn.dbCon);
                cmd.ExecuteNonQuery();
                strSql.Clear();
            }
            for (int i = 0; i < dtInitSdMain.Rows.Count; i++)
            {
                tmpGubun = dtInitSdMain.Rows[i]["gubun"].ToString();
                tmpValue = Convert.ToDouble(dtInitSdMain.Rows[i]["recovery"].ToString());
                strSql.AppendLine("Update meaipsiljuk set recovery = '" + tmpValue + "' ");
                strSql.AppendLine("Where  gubun = '" + tmpGubun + "' ");
                cmd = new SqlCommand(strSql.ToString(), DBConn.dbCon);
                cmd.ExecuteNonQuery();
                strSql.Clear();
            }
            cmd.Dispose();
            InitValue();
        }

        //엑셀 다운로드
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;

            }
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                excelApp = new Excel.Application();
                wb = excelApp.Workbooks.Open(@"d:\구매팀실적.xlsx", 0, true);
                ws = wb.Sheets["구매팀실적"];

                ws.Cells[3, 3] = TeditInsentiv.EditValue;
                ws.Cells[3, 4] = Convert.ToDouble(TeditBill.EditValue) / 100;
                ws.Cells[3, 5] = TeditBillInterest.EditValue;
                ws.Cells[3, 6] = DeditStart.EditValue;
                ws.Cells[3, 7] = DeditEnd.EditValue;
                ws.Cells[5, 5] = TeditMeachul.EditValue;
                ws.Cells[5, 6] = TeditDanga.EditValue;
                ws.Cells[5, 7] = TeditMajin.EditValue;
                ws.Cells[13, 5] = TeditMeachul1.EditValue;
                ws.Cells[13, 6] = TeditDanga1.EditValue;
                ws.Cells[13, 7] = TeditMajin1.EditValue;
                ws.Cells[12, 11] = TeditMeachul2.EditValue;
                ws.Cells[12, 13] = TeditWaste.EditValue;

                for (int i = 0; i < gridView1.RowCount; i++) //스크랩 중량
                {
                    ws.Cells[i + 8, 1] = gridView1.GetRowCellValue(i, "name");
                    ws.Cells[i + 8, 3] = gridView1.GetRowCellValue(i, "sum(a.danjung)");
                    ws.Cells[i + 8, 4] = gridView1.GetRowCellValue(i, "sum(a.halin)");
                    ws.Cells[i + 8, 5] = gridView1.GetRowCellValue(i, "Ddanga");
                    ws.Cells[i + 8, 6] = gridView1.GetRowCellValue(i, "HeavyMajin");
                    ws.Cells[i + 8, 7] = gridView1.GetRowCellValue(i, "Meachul");
                }
                for (int i = 0; i < gridView3.RowCount; i++) //스크랩 경량 - 손상영
                {
                    ws.Cells[i + 16, 2] = gridView3.GetRowCellValue(i, "gubun1");
                    ws.Cells[i + 16, 3] = gridView3.GetRowCellValue(i, "sum(a.danjung)");
                    ws.Cells[i + 16, 4] = gridView3.GetRowCellValue(i, "sum(a.halin)");
                    ws.Cells[i + 16, 5] = gridView3.GetRowCellValue(i, "Ddanga");
                    ws.Cells[i + 16, 6] = gridView3.GetRowCellValue(i, "LightMajin");
                    ws.Cells[i + 16, 7] = gridView3.GetRowCellValue(i, "Meachul");
                }
                for (int i = 0; i < gridView4.RowCount; i++) //스크랩 경량 - 오상훈
                {
                    ws.Cells[i + 25, 2] = gridView4.GetRowCellValue(i, "gubun1");
                    ws.Cells[i + 25, 3] = gridView4.GetRowCellValue(i, "sum(a.danjung)");
                    ws.Cells[i + 25, 4] = gridView4.GetRowCellValue(i, "sum(a.halin)");
                    ws.Cells[i + 25, 5] = gridView4.GetRowCellValue(i, "Ddanga");
                    ws.Cells[i + 25, 6] = gridView4.GetRowCellValue(i, "LightMajin");
                    ws.Cells[i + 25, 7] = gridView4.GetRowCellValue(i, "Meachul");
                }
                for (int i = 0; i < gridView5.RowCount; i++) //스크랩 경량 - 김명철
                {
                    ws.Cells[i + 34, 2] = gridView5.GetRowCellValue(i, "gubun1");
                    ws.Cells[i + 34, 3] = gridView5.GetRowCellValue(i, "sum(a.danjung)");
                    ws.Cells[i + 34, 4] = gridView5.GetRowCellValue(i, "sum(a.halin)");
                    ws.Cells[i + 34, 5] = gridView5.GetRowCellValue(i, "Ddanga");
                    ws.Cells[i + 34, 6] = gridView5.GetRowCellValue(i, "LightMajin");
                    ws.Cells[i + 34, 7] = gridView5.GetRowCellValue(i, "Meachul");
                }
                for (int i = 0; i < gridView2.RowCount; i++) //직송
                {
                    ws.Cells[i + 8, 8]  = gridView2.GetRowCellValue(i, "name");
                    ws.Cells[i + 8, 9]  = gridView2.GetRowCellValue(i, "sum(a.danjung)");
                    ws.Cells[i + 8, 10] = gridView2.GetRowCellValue(i, "Meachul");
                    ws.Cells[i + 8, 11] = gridView2.GetRowCellValue(i, "Ddanga");
                    ws.Cells[i + 8, 12] = gridView2.GetRowCellValue(i, "MeachulProfit");
                }
                for (int i = 0; i < gridView6.RowCount; i++) //슈레더 - 손상영
                {
                    ws.Cells[i + 16, 9]  = gridView6.GetRowCellValue(i, "gubun1");
                    ws.Cells[i + 16, 10] = gridView6.GetRowCellValue(i, "sum(a.danjung)");
                    ws.Cells[i + 16, 11] = gridView6.GetRowCellValue(i, "Danga");
                    ws.Cells[i + 16, 12] = gridView6.GetRowCellValue(i, "DAmont");
                    ws.Cells[i + 16, 13] = gridView6.GetRowCellValue(i, "SdMeachul");
                }
                for (int i = 0; i < gridView7.RowCount; i++) //슈레더 - 오상훈
                {
                    ws.Cells[i + 25, 9]  = gridView7.GetRowCellValue(i, "gubun1");
                    ws.Cells[i + 25, 10] = gridView7.GetRowCellValue(i, "sum(a.danjung)");
                    ws.Cells[i + 25, 11] = gridView7.GetRowCellValue(i, "Danga");
                    ws.Cells[i + 25, 12] = gridView7.GetRowCellValue(i, "DAmont");
                    ws.Cells[i + 25, 13] = gridView7.GetRowCellValue(i, "SdMeachul");
                }
                for (int i = 0; i < gridView8.RowCount; i++) //슈레더 - 김명철
                {
                    ws.Cells[i + 34, 9]  = gridView8.GetRowCellValue(i, "gubun1");
                    ws.Cells[i + 34, 10] = gridView8.GetRowCellValue(i, "sum(a.danjung)");
                    ws.Cells[i + 34, 11] = gridView8.GetRowCellValue(i, "Danga");
                    ws.Cells[i + 34, 12] = gridView8.GetRowCellValue(i, "DAmont");
                    ws.Cells[i + 34, 13] = gridView8.GetRowCellValue(i, "SdMeachul");
                }

                excelApp.Visible = true;
            }
            finally
            {

            }
        }

        private void simpleButton5_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }
        }

        private void simpleButton6_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }
        }

        private void LstSaleResult_KeyDown(object sender, KeyEventArgs e)
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
                BtnSerch.PerformClick();
            }
            else if (e.KeyCode == Keys.F8)
            {
                simpleButton1.PerformClick();
            }
        }
    }
}
