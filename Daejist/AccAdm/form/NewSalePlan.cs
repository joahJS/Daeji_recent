using ComLib;
using DevExpress.XtraEditors;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

/*
 * 수정일자: 2022-09-14
 * 수정자  : 정은영
 * ID      : #0001
 * 수정내용: (현업요청)
 *           1. 직송 -> 경량C 로 변경
 *           2. 경량,경량C,중량으로 그리드 밴드 순서변경
 *           3. 경량C 합계수정
 *           4. 경량,중량,경량c 합계 야드매입목표 합계로 자동 작성 추가
 *           5. 엑셀양식수정
 * 
 * 
 */

namespace AccAdm
{
    public partial class NewSalePlan : DevExpress.XtraEditors.XtraForm
    {
        StringBuilder strSql = new StringBuilder();
        string sYYMM;
        int countMajin1, countMajin2, countMajin3, countMajin4, countMajin5, countMajinTot;
        int sumMajin1, sumMajin2, sumMajin3, sumMajin4, sumMajin5, sumMajinTot;
        double TotMajinCount, TotMajinSum = 0;
        public NewSalePlan()
        {
            DBConn.DbConn();
            InitializeComponent();
        }
        public DataRow rowUserInfo { get; set; }
        private void NewSalePlan_Load(object sender, EventArgs e)
        {
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            DEYYMM.EditValue = Convert.ToDateTime(DateTime.Today.ToShortDateString());
            sYYMM = DEYYMM.EditValue.ToString().Replace("-", "").Substring(0, 6);

            DataTable TmpMeaip = new DataTable("TmpMeaip");
            DataColumn cGubun = new DataColumn("gubun", typeof(string));
            DataColumn clight1 = new DataColumn("light1", typeof(double));
            DataColumn clight2 = new DataColumn("light2", typeof(double));
            DataColumn clight3 = new DataColumn("light3", typeof(double));
            DataColumn clight4 = new DataColumn("light4", typeof(double));
            //DataColumn clight5 = new DataColumn("light5", typeof(double));
            DataColumn clightTot = new DataColumn("lightTot", typeof(double));

            DataColumn cweight1 = new DataColumn("weight1", typeof(double));
            DataColumn cweight2 = new DataColumn("weight2", typeof(double));
            DataColumn cweight3 = new DataColumn("weight3", typeof(double));
            DataColumn cweight4 = new DataColumn("weight4", typeof(double));
            //DataColumn cweight5 = new DataColumn("weight5", typeof(double));
            DataColumn cweightTot = new DataColumn("weightTot", typeof(double));

            DataColumn cyk1 = new DataColumn("yk1", typeof(double));
            DataColumn cyk2 = new DataColumn("yk2", typeof(double));
            DataColumn cyk3 = new DataColumn("yk3", typeof(double));
            DataColumn cyk4 = new DataColumn("yk4", typeof(double));
            //DataColumn cyk5 = new DataColumn("yk5", typeof(double));
            DataColumn cykTot = new DataColumn("ykTot", typeof(double));

            //DataColumn cmajin1 = new DataColumn("majin1", typeof(double));
            //DataColumn cmajin2 = new DataColumn("majin2", typeof(double));
            //DataColumn cmajin3 = new DataColumn("majin3", typeof(double));
            //DataColumn cmajin4 = new DataColumn("majin4", typeof(double));
            //DataColumn cmajin5 = new DataColumn("majin5", typeof(double));
            //DataColumn cmajinTot = new DataColumn("majinTot", typeof(double));

            DataColumn cjik1 = new DataColumn("jik1", typeof(double));
            DataColumn cjik2 = new DataColumn("jik2", typeof(double));
            DataColumn cjik3 = new DataColumn("jik3", typeof(double));
            DataColumn cjik4 = new DataColumn("jik4", typeof(double));
            //DataColumn cjik5 = new DataColumn("jik5", typeof(double));
            DataColumn cjikTot = new DataColumn("jikTot", typeof(double));

            DataColumn cban1 = new DataColumn("Ban1", typeof(double));
            DataColumn cban2 = new DataColumn("Ban2", typeof(double));
            DataColumn cban3 = new DataColumn("Ban3", typeof(double));
            DataColumn cban4 = new DataColumn("Ban4", typeof(double));
            //DataColumn cban5 = new DataColumn("ban5", typeof(double));
            DataColumn cbanTot = new DataColumn("BanTot", typeof(double));

            DataColumn cEchapi1 = new DataColumn("Echapi1", typeof(double));
            DataColumn cEchapi2 = new DataColumn("Echapi2", typeof(double));
            DataColumn cEchapi3 = new DataColumn("Echapi3", typeof(double));
            DataColumn cEchapi4 = new DataColumn("Echapi4", typeof(double));
            //DataColumn Echapi5 = new DataColumn("Echapi5", typeof(double));
            DataColumn cEchapiTot = new DataColumn("EchapiTot", typeof(double));

            DataColumn cchapi1 = new DataColumn("chapi1", typeof(double));
            DataColumn cchapi2 = new DataColumn("chapi2", typeof(double));
            DataColumn cchapi3 = new DataColumn("chapi3", typeof(double));
            DataColumn cchapi4 = new DataColumn("chapi4", typeof(double));
            //DataColumn cchapi5 = new DataColumn("chapi5", typeof(double));
            DataColumn cchapiTot = new DataColumn("chapiTot", typeof(double));
            TmpMeaip.Columns.Add(cGubun);
            TmpMeaip.Columns.Add(clight1); TmpMeaip.Columns.Add(clight2); TmpMeaip.Columns.Add(clight3); TmpMeaip.Columns.Add(clight4);  TmpMeaip.Columns.Add(clightTot);
            TmpMeaip.Columns.Add(cweight1); TmpMeaip.Columns.Add(cweight2); TmpMeaip.Columns.Add(cweight3); TmpMeaip.Columns.Add(cweight4);  TmpMeaip.Columns.Add(cweightTot);
            TmpMeaip.Columns.Add(cyk1); TmpMeaip.Columns.Add(cyk2); TmpMeaip.Columns.Add(cyk3); TmpMeaip.Columns.Add(cyk4);  TmpMeaip.Columns.Add(cykTot);
            //TmpMeaip.Columns.Add(cmajin1); TmpMeaip.Columns.Add(cmajin2); TmpMeaip.Columns.Add(cmajin3); TmpMeaip.Columns.Add(cmajin4);  TmpMeaip.Columns.Add(cmajinTot);
            TmpMeaip.Columns.Add(cjik1); TmpMeaip.Columns.Add(cjik2); TmpMeaip.Columns.Add(cjik3); TmpMeaip.Columns.Add(cjik4);  TmpMeaip.Columns.Add(cjikTot);
            TmpMeaip.Columns.Add(cban1); TmpMeaip.Columns.Add(cban2); TmpMeaip.Columns.Add(cban3); TmpMeaip.Columns.Add(cban4); TmpMeaip.Columns.Add(cbanTot);
            TmpMeaip.Columns.Add(cEchapi1); TmpMeaip.Columns.Add(cEchapi2); TmpMeaip.Columns.Add(cEchapi3); TmpMeaip.Columns.Add(cEchapi4); TmpMeaip.Columns.Add(cEchapiTot);
            TmpMeaip.Columns.Add(cchapi1); TmpMeaip.Columns.Add(cchapi2); TmpMeaip.Columns.Add(cchapi3); TmpMeaip.Columns.Add(cchapi4);  TmpMeaip.Columns.Add(cchapiTot);
            //타이틀 생성
            TmpMeaip.Rows.Add("1주차", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            TmpMeaip.Rows.Add("2주차", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            TmpMeaip.Rows.Add("3주차", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            TmpMeaip.Rows.Add("4주차", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            TmpMeaip.Rows.Add("5주차", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            TmpMeaip.Rows.Add("6주차", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            gridControl1.DataSource = TmpMeaip;

            //매출목표주간 테이블 생성======================================
            DataTable tmpMaechul = new DataTable("tmpMaechul");
            DataColumn McGubun = new DataColumn("gubun", typeof(string));
            DataColumn Mc1jucha = new DataColumn("jucha1", typeof(double));
            DataColumn Mc1amount = new DataColumn("amount1", typeof(double));
            DataColumn Mc2jucha = new DataColumn("jucha2", typeof(double));
            DataColumn Mc2amount = new DataColumn("amount2", typeof(double));
            DataColumn Mc3jucha = new DataColumn("jucha3", typeof(double));
            DataColumn Mc3amount = new DataColumn("amount3", typeof(double));
            DataColumn Mc4jucha = new DataColumn("jucha4", typeof(double));
            DataColumn Mc4amount = new DataColumn("amount4", typeof(double));
            DataColumn Mc5jucha = new DataColumn("jucha5", typeof(double));
            DataColumn Mc5amount = new DataColumn("amount5", typeof(double));
            DataColumn Mc6jucha = new DataColumn("jucha6", typeof(double));
            DataColumn Mc6amount = new DataColumn("amount6", typeof(double));
            DataColumn McjuchaTot = new DataColumn("juchaTot", typeof(double));
            DataColumn McTotamount = new DataColumn("amountTot", typeof(double));

            tmpMaechul.Columns.Add(McGubun);
            tmpMaechul.Columns.Add(Mc1jucha); tmpMaechul.Columns.Add(Mc1amount); tmpMaechul.Columns.Add(Mc2jucha); tmpMaechul.Columns.Add(Mc2amount);
            tmpMaechul.Columns.Add(Mc3jucha); tmpMaechul.Columns.Add(Mc3amount); tmpMaechul.Columns.Add(Mc4jucha); tmpMaechul.Columns.Add(Mc4amount);
            tmpMaechul.Columns.Add(Mc5jucha); tmpMaechul.Columns.Add(Mc5amount); tmpMaechul.Columns.Add(Mc6jucha); tmpMaechul.Columns.Add(Mc6amount);
            tmpMaechul.Columns.Add(McjuchaTot); tmpMaechul.Columns.Add(McTotamount);

            tmpMaechul.Rows.Add("야드매출목표(G/S)", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            tmpMaechul.Rows.Add("야드매출목표(중량)", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            tmpMaechul.Rows.Add("야드매출목표(S/D)", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            tmpMaechul.Rows.Add("직납매출목표", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            tmpMaechul.Rows.Add("수입직납목표", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
            gridControl2.DataSource = tmpMaechul;

            ////영업사원 명단조회
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine("  SELECT EMP_NM              ");
            strSql.AppendLine("  FROM HR_EMP_BASIS          ");
            strSql.AppendLine("  WHERE DEPT_CD = 3000       "); //영업팀 3000번으로 변경
            strSql.AppendLine("  AND EMPL_GB = 'Y'          ");
            strSql.AppendLine("  ORDER BY GRADE_CD, EMP_ID  ");

            DataTable dtName = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if(dtName != null && dtName.Rows.Count > 0)
            {
                gridBand3.Caption = dtName.Rows[0]["EMP_NM"]?.ToString();

                if (dtName.Rows.Count > 1)
                {
                    gridBand4.Caption = dtName.Rows[1]["EMP_NM"]?.ToString();
                    if (dtName.Rows.Count > 2)
                    {
                        gridBand5.Caption = dtName.Rows[2]["EMP_NM"]?.ToString();
                        if (dtName.Rows.Count > 3)
                        {
                            gridBand6.Caption = dtName.Rows[3]["EMP_NM"]?.ToString();
                            if (dtName.Rows.Count > 4)
                            {
                                //gridBand59.Caption = dtName.Rows[4]["EMP_NM"]?.ToString();
                               
                            }
                        }
                    }
                }
            }
            else
            {
                gridBand3.Caption = string.Empty;
                gridBand4.Caption = string.Empty;
                gridBand5.Caption = string.Empty;
                gridBand6.Caption = string.Empty;
                
            }
            
            btSave.Visible = false;

        }

        private void btSerch_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            btSave.Visible = false;
            advBandedGridView1.OptionsBehavior.Editable = false;
            advBandedGridView2.OptionsBehavior.Editable = false;

            sYYMM = DEYYMM.EditValue.ToString().Replace("-", "").Substring(0, 6);
            string Tmp_Name;
            //if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            //{
            //    XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
            //    return;
            //}
            initGrid();
            ////영업사원 명단조회
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine("  SELECT a.NAME             ");
            strSql.AppendLine("  FROM SALEMAEIP a, HR_EMP_BASIS b   ");
            strSql.AppendLine("  WHERE a.name = b.emp_nm            ");
            strSql.AppendLine("  AND   a.yymm = '" + sYYMM + "'              ");
            strSql.AppendLine("  GROUP BY A.Name, B.GRADE_CD, B.EMP_ID");
            strSql.AppendLine("  ORDER BY B.GRADE_CD, B.EMP_ID        ");

            DataTable dtName = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            gridBand3.Caption = "";
            gridBand4.Caption = "";
            gridBand5.Caption = "";
            gridBand6.Caption = "";
            //gridBand59.Caption = "";
            
            if (dtName == null || dtName.Rows.Count <= 0)
            {
                ////영업사원 명단조회
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine("  SELECT EMP_NM AS NAME             ");
                strSql.AppendLine("  FROM HR_EMP_BASIS          ");
                strSql.AppendLine("  WHERE DEPT_CD = 3000       "); //영업팀 3000번으로 변경
                strSql.AppendLine("  AND EMPL_GB = 'Y'          ");
                strSql.AppendLine("  ORDER BY GRADE_CD, EMP_ID  ");

                dtName = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            }

            if (dtName.Rows.Count > 0) gridBand3.Caption = dtName.Rows[0]["NAME"].ToString();
            if (dtName.Rows.Count > 1) gridBand4.Caption = dtName.Rows[1]["NAME"].ToString();
            if (dtName.Rows.Count > 2) gridBand5.Caption = dtName.Rows[2]["NAME"].ToString();
            if (dtName.Rows.Count > 3) gridBand6.Caption = dtName.Rows[3]["NAME"].ToString();
            
            
            Tmp_Name = gridBand3.Caption;

            //월 근무일수 조회 =====================================================================================

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT yymm, WorkDays ");
            strSql.AppendLine("  FROM MonWork ");
            strSql.AppendLine("  WHERE YYMM = '" + sYYMM + "' ");
            
            DataTable dt_Wday = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            te_MonWork.Text = Convert.ToString( 0 );
            if (dt_Wday.Rows.Count >= 1) te_MonWork.Text = dt_Wday.Rows[0][1].ToString();


            //매입계획 조회 =====================================================================================

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT YYMM, NAME, GUBUN, JUCHA, I_LIGHT, I_WEIGHT, I_YK, I_JIK, I_BAN, I_ECHAPI, I_CHAPI ");
            strSql.AppendLine("  FROM SALEMAEIP ");
            strSql.AppendLine("  WHERE YYMM = '" + sYYMM + "' ");
            //strSql.AppendLine("  AND NAME = '손상영' ");
            strSql.AppendLine("  AND NAME = '" + Tmp_Name + "' ");
            strSql.AppendLine("  ORDER BY NAME, JUCHA");

            DataTable dt1 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                advBandedGridView1.SetRowCellValue(i, "light1", Convert.ToInt16(dt1.Rows[i]["i_light"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "weight1", Convert.ToInt16(dt1.Rows[i]["i_weight"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "yk1", Convert.ToInt16(dt1.Rows[i]["i_yk"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "jik1", Convert.ToInt16(dt1.Rows[i]["i_jik"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "Ban1", Convert.ToInt16(dt1.Rows[i]["i_ban"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "Echapi1", Convert.ToInt16(dt1.Rows[i]["i_Echapi"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "chapi1", Convert.ToInt16(dt1.Rows[i]["i_chapi"].ToString()));
            }
            
            Tmp_Name = gridBand4.Caption;
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT YYMM, NAME, GUBUN, JUCHA, I_LIGHT, I_WEIGHT, I_YK, I_JIK, I_BAN, I_ECHAPI, I_CHAPI ");
            strSql.AppendLine("  FROM SALEMAEIP ");
            strSql.AppendLine("  WHERE YYMM = '" + sYYMM + "' ");
            //strSql.AppendLine("  AND NAME = '오상훈' ");
            strSql.AppendLine("  AND NAME = '" + Tmp_Name + "' ");
            strSql.AppendLine("  ORDER BY NAME, JUCHA");

            DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            for (int i = 0; i < dt2.Rows.Count; i++)
            {
                advBandedGridView1.SetRowCellValue(i, "light2", Convert.ToInt16(dt2.Rows[i]["i_light"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "weight2", Convert.ToInt16(dt2.Rows[i]["i_weight"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "yk2", Convert.ToInt16(dt2.Rows[i]["i_yk"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "jik2", Convert.ToInt16(dt2.Rows[i]["i_jik"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "Ban2", Convert.ToInt16(dt2.Rows[i]["i_ban"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "Echapi2", Convert.ToInt16(dt2.Rows[i]["i_Echapi"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "chapi2", Convert.ToInt16(dt2.Rows[i]["i_chapi"].ToString()));
                
            }
            
            Tmp_Name = gridBand5.Caption;
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT YYMM, NAME, GUBUN, JUCHA, I_LIGHT, I_WEIGHT, I_YK, I_JIK, I_BAN, I_ECHAPI, I_CHAPI ");
            strSql.AppendLine("  FROM SALEMAEIP ");
            strSql.AppendLine("  WHERE YYMM = '" + sYYMM + "' ");
            //strSql.AppendLine("  AND NAME = '김명철' ");
            strSql.AppendLine("  AND NAME = '" + Tmp_Name + "' ");
            strSql.AppendLine("  ORDER BY NAME, JUCHA");

            DataTable dt3 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            for (int i = 0; i < dt3.Rows.Count; i++)
            {
                double kk = Convert.ToInt16(dt3.Rows[i]["i_light"].ToString());
                advBandedGridView1.SetRowCellValue(i, "light3", Convert.ToInt16(dt3.Rows[i]["i_light"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "weight3", Convert.ToInt16(dt3.Rows[i]["i_weight"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "yk3", Convert.ToInt16(dt3.Rows[i]["i_yk"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "jik3", Convert.ToInt16(dt3.Rows[i]["i_jik"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "Ban3", Convert.ToInt16(dt3.Rows[i]["i_ban"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "Echapi3", Convert.ToInt16(dt3.Rows[i]["i_Echapi"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "chapi3", Convert.ToInt16(dt3.Rows[i]["i_chapi"].ToString()));
                
            }
            
            Tmp_Name = gridBand6.Caption;
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT YYMM, NAME, GUBUN, JUCHA, I_LIGHT, I_WEIGHT, I_YK, I_JIK, I_BAN, I_ECHAPI, I_CHAPI ");
            strSql.AppendLine("  FROM SALEMAEIP ");
            strSql.AppendLine("  WHERE YYMM = '" + sYYMM + "' ");
            //strSql.AppendLine("  AND NAME = '기타' ");
            strSql.AppendLine("  AND NAME = '" + Tmp_Name + "' ");
            strSql.AppendLine("  ORDER BY NAME, JUCHA");

            DataTable dt4 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            for (int i = 0; i < dt4.Rows.Count; i++)
            {
                advBandedGridView1.SetRowCellValue(i, "light4", Convert.ToInt16(dt4.Rows[i]["i_light"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "weight4", Convert.ToInt16(dt4.Rows[i]["i_weight"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "yk4", Convert.ToInt16(dt4.Rows[i]["i_yk"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "jik4", Convert.ToInt16(dt4.Rows[i]["i_jik"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "Ban4", Convert.ToInt16(dt4.Rows[i]["i_ban"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "Echapi4", Convert.ToInt16(dt4.Rows[i]["i_Echapi"].ToString()));
                advBandedGridView1.SetRowCellValue(i, "chapi4", Convert.ToInt16(dt4.Rows[i]["i_chapi"].ToString()));
                
            }



            //매출계획 조회 =====================================================================================
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT YYMM, JUCHA, o_gs, o_weight, o_sd, o_yk, o_suip ");
            strSql.AppendLine("  FROM SALEMAECHUL                                      ");
            strSql.AppendLine("  WHERE YYMM = '" + sYYMM + "'                          ");
            strSql.AppendLine("  ORDER BY JUCHA                                        ");

            DataTable dt15 = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            string Tmpjucha = "";
            for (int i = 0; i < dt15.Rows.Count; i++)
            {
                if (i == 0) Tmpjucha = "o_gs";
                if (i == 1) Tmpjucha = "o_weight";
                if (i == 2) Tmpjucha = "o_sd";
                if (i == 3) Tmpjucha = "o_yk";
                if (i == 4) Tmpjucha = "o_suip";
                advBandedGridView2.SetRowCellValue(i, "jucha1", Convert.ToDouble(dt15.Rows[0][Tmpjucha].ToString()));
                advBandedGridView2.SetRowCellValue(i, "jucha2", Convert.ToDouble(dt15.Rows[1][Tmpjucha].ToString()));
                advBandedGridView2.SetRowCellValue(i, "jucha3", Convert.ToDouble(dt15.Rows[2][Tmpjucha].ToString()));
                advBandedGridView2.SetRowCellValue(i, "jucha4", Convert.ToDouble(dt15.Rows[3][Tmpjucha].ToString()));
                advBandedGridView2.SetRowCellValue(i, "jucha5", Convert.ToDouble(dt15.Rows[4][Tmpjucha].ToString()));
                advBandedGridView2.SetRowCellValue(i, "jucha6", Convert.ToDouble(dt15.Rows[5][Tmpjucha].ToString()));
            }

            SumInfo();
            Cursor = Cursors.Default;
            btSave.Visible = false;
        }
              
        private void btClear_Click(object sender, EventArgs e)
        {
            initGrid();
        }

        private void DEYYMM_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                btSerch.PerformClick();
            }
        }

        private void btSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            sYYMM = DEYYMM.EditValue.ToString().Replace("-", "").Substring(0, 6);
            string InsertSql, DelSql;
            Cursor = Cursors.WaitCursor;
            
            //근무일수 저장
            DelSql = "DELETE FROM MonWork WHERE YYMM = '" + sYYMM + "'  ";
            SqlCommand mwDel = new SqlCommand(DelSql, DBConn.dbCon);
            mwDel.ExecuteNonQuery();

            InsertSql = "INSERT INTO MonWork(yymm, WorkDays) VALUES(@yymm, @mwDays) ";
            SqlCommand mwInsert = new SqlCommand(InsertSql, DBConn.dbCon);
            mwInsert.Parameters.Add("@yymm", SqlDbType.VarChar).Value = sYYMM;
            mwInsert.Parameters.Add("@mwDays", SqlDbType.Int).Value = Convert.ToInt16(te_MonWork.Text);

            mwInsert.ExecuteNonQuery();
            
            //매입계획 저장================================================================================================
            DelSql = "DELETE FROM SALEMAEIP WHERE YYMM = '" + sYYMM + "'  ";
            SqlCommand cmdDel = new SqlCommand(DelSql, DBConn.dbCon);
            cmdDel.ExecuteNonQuery();

            InsertSql = "INSERT INTO SALEMAEIP(yymm, Name, jucha, gubun, i_light, i_weight, i_yk, i_jik, i_ban, i_Echapi, i_chapi) VALUES(@yymm, @name, @jucha, @gubun, @i_light, @i_weight, @i_yk, @i_jik, @i_ban, @i_Echapi, @i_chapi) ";
            SqlCommand cmdInsert = new SqlCommand(InsertSql, DBConn.dbCon);

            //cmdInsert.Parameters.Add("@serial", SqlDbType.Int);
            cmdInsert.Parameters.Add("@yymm", SqlDbType.VarChar);
            cmdInsert.Parameters.Add("@name", SqlDbType.VarChar);
            cmdInsert.Parameters.Add("@jucha", SqlDbType.VarChar);
            cmdInsert.Parameters.Add("@gubun", SqlDbType.VarChar);
            cmdInsert.Parameters.Add("@i_light", SqlDbType.Int);
            cmdInsert.Parameters.Add("@i_weight", SqlDbType.Int);
            cmdInsert.Parameters.Add("@i_yk", SqlDbType.Int);
            cmdInsert.Parameters.Add("@i_jik", SqlDbType.Int);
            cmdInsert.Parameters.Add("@i_ban", SqlDbType.Int);
            cmdInsert.Parameters.Add("@i_Echapi", SqlDbType.Int);
            cmdInsert.Parameters.Add("@i_chapi", SqlDbType.Int);
            for (int i = 0; i < 6; i++)
            {   //손상영
                cmdInsert.Parameters["@yymm"].Value = sYYMM;
                cmdInsert.Parameters["@name"].Value = gridBand3.Caption;
                cmdInsert.Parameters["@jucha"].Value = i + 1;
                cmdInsert.Parameters["@gubun"].Value = "매입";
                cmdInsert.Parameters["@i_light"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "light1"));
                cmdInsert.Parameters["@i_weight"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "weight1"));
                cmdInsert.Parameters["@i_yk"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "yk1"));
                cmdInsert.Parameters["@i_jik"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "jik1"));
                cmdInsert.Parameters["@i_ban"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Ban1"));
                cmdInsert.Parameters["@i_Echapi"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Echapi1"));
                cmdInsert.Parameters["@i_chapi"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "chapi1"));

                cmdInsert.ExecuteNonQuery();

                //오상훈
                cmdInsert.Parameters["@yymm"].Value = sYYMM;
                cmdInsert.Parameters["@name"].Value = gridBand4.Caption;
                cmdInsert.Parameters["@jucha"].Value = i + 1;
                cmdInsert.Parameters["@gubun"].Value = "매입";
                cmdInsert.Parameters["@i_light"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "light2"));
                cmdInsert.Parameters["@i_weight"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "weight2"));
                cmdInsert.Parameters["@i_yk"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "yk2"));
                cmdInsert.Parameters["@i_jik"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "jik2"));
                cmdInsert.Parameters["@i_ban"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Ban2"));
                cmdInsert.Parameters["@i_Echapi"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Echapi2"));
                cmdInsert.Parameters["@i_chapi"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "chapi2"));

                cmdInsert.ExecuteNonQuery();

                //김명철
                cmdInsert.Parameters["@yymm"].Value = sYYMM;
                cmdInsert.Parameters["@name"].Value = gridBand5.Caption;
                cmdInsert.Parameters["@jucha"].Value = i + 1;
                cmdInsert.Parameters["@gubun"].Value = "매입";
                cmdInsert.Parameters["@i_light"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "light3"));
                cmdInsert.Parameters["@i_weight"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "weight3"));
                cmdInsert.Parameters["@i_yk"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "yk3"));
                cmdInsert.Parameters["@i_jik"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "jik3"));
                cmdInsert.Parameters["@i_ban"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Ban3"));
                cmdInsert.Parameters["@i_Echapi"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Echapi3"));
                cmdInsert.Parameters["@i_chapi"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "chapi3"));

                cmdInsert.ExecuteNonQuery();

                //기타
                cmdInsert.Parameters["@yymm"].Value = sYYMM;
                cmdInsert.Parameters["@name"].Value = gridBand6.Caption;
                cmdInsert.Parameters["@jucha"].Value = i + 1;
                cmdInsert.Parameters["@gubun"].Value = "매입";
                cmdInsert.Parameters["@i_light"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "light4"));
                cmdInsert.Parameters["@i_weight"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "weight4"));
                cmdInsert.Parameters["@i_yk"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "yk4"));
                cmdInsert.Parameters["@i_jik"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "jik4"));
                cmdInsert.Parameters["@i_ban"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Ban4"));
                cmdInsert.Parameters["@i_Echapi"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Echapi4"));
                cmdInsert.Parameters["@i_chapi"].Value = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "chapi4"));

                cmdInsert.ExecuteNonQuery();

               


                //매출계획 저장================================================================================================
                string InsertSqlMaechul, DelSqlMaechul, Tmpjucha1 = "";
                DelSqlMaechul = "DELETE FROM SALEMAECHUL WHERE YYMM = '" + sYYMM + "'  ";
                SqlCommand cmdDelMaechul = new SqlCommand(DelSqlMaechul, DBConn.dbCon);
                cmdDelMaechul.ExecuteNonQuery();

                InsertSqlMaechul = "INSERT INTO SALEMAECHUL(yymm, jucha, o_gs, o_weight, o_sd, o_yk, o_suip) VALUES(@yymm, @jucha, @o_gs, @o_weight, @o_sd, @o_yk, @o_suip) ";
                SqlCommand cmdInsertMaechul = new SqlCommand(InsertSqlMaechul, DBConn.dbCon);

                //cmdInsertMaechul.Parameters.Add("@serial", SqlDbType.Int);
                cmdInsertMaechul.Parameters.Add("@yymm", SqlDbType.VarChar);
                cmdInsertMaechul.Parameters.Add("@jucha", SqlDbType.VarChar);
                cmdInsertMaechul.Parameters.Add("@o_gs", SqlDbType.Int);
                cmdInsertMaechul.Parameters.Add("@o_weight", SqlDbType.Int);
                cmdInsertMaechul.Parameters.Add("@o_sd", SqlDbType.Int);
                cmdInsertMaechul.Parameters.Add("@o_yk", SqlDbType.Int);
                cmdInsertMaechul.Parameters.Add("@o_suip", SqlDbType.Int);

                for (int j = 0; j < 6; j++)
                {   //GS매출
                    if (j == 0) Tmpjucha1 = "jucha1";
                    if (j == 1) Tmpjucha1 = "jucha2";
                    if (j == 2) Tmpjucha1 = "jucha3";
                    if (j == 3) Tmpjucha1 = "jucha4";
                    if (j == 4) Tmpjucha1 = "jucha5";
                    if (j == 5) Tmpjucha1 = "jucha6";
                    cmdInsertMaechul.Parameters["@yymm"].Value = sYYMM;
                    cmdInsertMaechul.Parameters["@jucha"].Value = j + 1;
                    cmdInsertMaechul.Parameters["@o_gs"].Value = Convert.ToInt16(advBandedGridView2.GetRowCellValue(0, Tmpjucha1));
                    cmdInsertMaechul.Parameters["@o_weight"].Value = Convert.ToInt16(advBandedGridView2.GetRowCellValue(1, Tmpjucha1));
                    cmdInsertMaechul.Parameters["@o_sd"].Value = Convert.ToInt16(advBandedGridView2.GetRowCellValue(2, Tmpjucha1));
                    cmdInsertMaechul.Parameters["@o_yk"].Value = Convert.ToInt16(advBandedGridView2.GetRowCellValue(3, Tmpjucha1));
                    cmdInsertMaechul.Parameters["@o_suip"].Value = Convert.ToInt16(advBandedGridView2.GetRowCellValue(4, Tmpjucha1));

                    cmdInsertMaechul.ExecuteNonQuery();
                }
            }
            btSave.Visible = false;
            te_MonWork.ReadOnly = true;
            advBandedGridView1.OptionsBehavior.Editable = false;
            advBandedGridView2.OptionsBehavior.Editable = false;
            Cursor = Cursors.Default;
        }
        public void SumInfo()
        {
            //매입===================================================================================
            double SumLight, SumWeight, SumYk, SumJik, SumChapi, SumBan, SumEChapi;
            advBandedGridView1.CloseEditor();
            for (int i = 0; i < 6; i++)
            {
                SumLight = 0;
                SumWeight = 0;
                SumYk = 0;
                SumJik = 0;
                SumBan = 0;
                SumEChapi = 0;
                SumChapi = 0;
                SumLight = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "light1").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "light2").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "light3").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "light4").ToString());
                //+ Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "light5").ToString());

                SumWeight = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "weight1").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "weight2").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "weight3").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "weight4").ToString());
                //+ Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "weight5").ToString());

                SumYk = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "yk1").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "yk2").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "yk3").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "yk4").ToString());
                //+ Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "yk5").ToString());

                SumBan = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Ban1").ToString())
                         + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Ban2").ToString())
                         + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Ban3").ToString())
                         + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Ban4").ToString());
                //+ Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "yk5").ToString());

                SumEChapi = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Echapi1").ToString())
                         + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Echapi2").ToString())
                         + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Echapi3").ToString())
                         + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "Echapi4").ToString());
                //+ Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "yk5").ToString());

                SumJik = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "jik1").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "jik2").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "jik3").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "jik4").ToString());
                //+ Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "jik5").ToString());

               

                SumChapi = Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "chapi1").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "chapi2").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "chapi3").ToString())
                          + Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "chapi4").ToString());
                          //+ Convert.ToInt16(advBandedGridView1.GetRowCellValue(i, "chapi5").ToString());
                          
                advBandedGridView1.SetRowCellValue(i, "lightTot", SumLight);
                advBandedGridView1.SetRowCellValue(i, "weightTot", SumWeight);
                advBandedGridView1.SetRowCellValue(i, "ykTot", SumYk);
                advBandedGridView1.SetRowCellValue(i, "jikTot", SumJik);
               
                advBandedGridView1.SetRowCellValue(i, "BanTot", SumBan);
                advBandedGridView1.SetRowCellValue(i, "EchapiTot", SumEChapi);
                advBandedGridView1.SetRowCellValue(i, "chapiTot", SumChapi);
                
                //#0001
                teMeaipSum.EditValue = Convert.ToInt16(advBandedGridView1.Columns["lightTot"].SummaryItem.SummaryValue.ToString())
                                     + Convert.ToInt16(advBandedGridView1.Columns["ykTot"].SummaryItem.SummaryValue.ToString())
                                     + Convert.ToInt16(advBandedGridView1.Columns["weightTot"].SummaryItem.SummaryValue.ToString())
                                     + Convert.ToInt16(advBandedGridView1.Columns["BanTot"].SummaryItem.SummaryValue.ToString())
                                     + Convert.ToInt16(advBandedGridView1.Columns["EchapiTot"].SummaryItem.SummaryValue.ToString())
                                     + Convert.ToInt16(advBandedGridView1.Columns["chapiTot"].SummaryItem.SummaryValue.ToString());
                
            }
            

            //매출===================================================================================
            double Sumhap;
            advBandedGridView2.CloseEditor();
            int kg = 0;
            for (int i = 0; i < 5; i++)
            {
                if (i == 0) kg = 24;  //GS
                if (i == 1) kg = 25;  //중량
                if (i == 2) kg = 27;  //SD
                if (i == 3) kg = 25;  //직납
                if (i == 4) kg = 25;  //수입

                Sumhap = 0;
                Sumhap = Convert.ToInt16(advBandedGridView2.GetRowCellValue(i, "jucha1").ToString())
                       + Convert.ToInt16(advBandedGridView2.GetRowCellValue(i, "jucha2").ToString())
                       + Convert.ToInt16(advBandedGridView2.GetRowCellValue(i, "jucha3").ToString())
                       + Convert.ToInt16(advBandedGridView2.GetRowCellValue(i, "jucha4").ToString())
                       + Convert.ToInt16(advBandedGridView2.GetRowCellValue(i, "jucha5").ToString())
                       + Convert.ToInt16(advBandedGridView2.GetRowCellValue(i, "jucha6").ToString());

                advBandedGridView2.SetRowCellValue(i, "amount1", Convert.ToInt16(advBandedGridView2.GetRowCellValue(i, "jucha1").ToString()) / kg);
                advBandedGridView2.SetRowCellValue(i, "amount2", Convert.ToInt16(advBandedGridView2.GetRowCellValue(i, "jucha2").ToString()) / kg);
                advBandedGridView2.SetRowCellValue(i, "amount3", Convert.ToInt16(advBandedGridView2.GetRowCellValue(i, "jucha3").ToString()) / kg);
                advBandedGridView2.SetRowCellValue(i, "amount4", Convert.ToInt16(advBandedGridView2.GetRowCellValue(i, "jucha4").ToString()) / kg);
                advBandedGridView2.SetRowCellValue(i, "amount5", Convert.ToInt16(advBandedGridView2.GetRowCellValue(i, "jucha5").ToString()) / kg);
                advBandedGridView2.SetRowCellValue(i, "amount6", Convert.ToInt16(advBandedGridView2.GetRowCellValue(i, "jucha6").ToString()) / kg);
                advBandedGridView2.SetRowCellValue(i, "amountTot", Sumhap / kg);
                advBandedGridView2.SetRowCellValue(i, "juchaTot", Sumhap);
                teMeachulSum.EditValue = Convert.ToInt16(advBandedGridView2.GetRowCellValue(0, "juchaTot").ToString())
                                       + Convert.ToInt16(advBandedGridView2.GetRowCellValue(1, "juchaTot").ToString())
                                       + Convert.ToInt16(advBandedGridView2.GetRowCellValue(2, "juchaTot").ToString());
            }
        }
        private void advBandedGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Down || e.KeyCode == Keys.Up || e.KeyCode == Keys.Left || e.KeyCode == Keys.Right)
            {
                SumInfo();
            }
            if (e.KeyCode == Keys.Enter) advBandedGridView1.MoveNext();
            advBandedGridView1.RefreshData();
        }
        private void advBandedGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Down || e.KeyCode == Keys.Up || e.KeyCode == Keys.Left || e.KeyCode == Keys.Right)
            {
                SumInfo();
            }
            if (e.KeyCode == Keys.Enter) advBandedGridView2.MoveNext();
            advBandedGridView2.RefreshData();
        }

        private void advBandedGridView1_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            advBandedGridView1.RefreshData();
        }

        private void advBandedGridView2_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            advBandedGridView2.RefreshData();
        }

        private void gridControl1_Click(object sender, EventArgs e)
        {
            SumInfo();
            advBandedGridView1.RefreshData();
        }

        private void gridControl2_Click(object sender, EventArgs e)
        {
            SumInfo();
            advBandedGridView2.RefreshData();
        }

        private void btEdit_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }
            btSave.Visible = true;
            te_MonWork.ReadOnly = false;
            advBandedGridView1.OptionsBehavior.Editable = true;
            advBandedGridView2.OptionsBehavior.Editable = true;
        }

        void initGrid()
        {
            teMeachulSum.EditValue = "";
            teMeaipSum.EditValue = "";
            //매입계획테이블 초기화
            for (int k = 0; k < 5; k++)
            {
                advBandedGridView1.SetRowCellValue(k, "light1", 0);
                advBandedGridView1.SetRowCellValue(k, "weight1", 0);
                advBandedGridView1.SetRowCellValue(k, "yk1", 0);
                advBandedGridView1.SetRowCellValue(k, "jik1", 0);
                advBandedGridView1.SetRowCellValue(k, "Ban1", 0);
                advBandedGridView1.SetRowCellValue(k, "Echapi1", 0);
                advBandedGridView1.SetRowCellValue(k, "chapi1", 0);
                advBandedGridView1.SetRowCellValue(k, "light2", 0);
                advBandedGridView1.SetRowCellValue(k, "weight2", 0);
                advBandedGridView1.SetRowCellValue(k, "yk2", 0);
                advBandedGridView1.SetRowCellValue(k, "jik2", 0);
                advBandedGridView1.SetRowCellValue(k, "Ban2", 0);
                advBandedGridView1.SetRowCellValue(k, "Echapi2", 0);
                advBandedGridView1.SetRowCellValue(k, "chapi2", 0);
                advBandedGridView1.SetRowCellValue(k, "light3", 0);
                advBandedGridView1.SetRowCellValue(k, "weight3", 0);
                advBandedGridView1.SetRowCellValue(k, "yk3", 0);
                advBandedGridView1.SetRowCellValue(k, "jik3", 0);
                advBandedGridView1.SetRowCellValue(k, "Ban3", 0);
                advBandedGridView1.SetRowCellValue(k, "Echapi3", 0);
                advBandedGridView1.SetRowCellValue(k, "chapi3", 0);
                advBandedGridView1.SetRowCellValue(k, "light4", 0);
                advBandedGridView1.SetRowCellValue(k, "weight4", 0);
                advBandedGridView1.SetRowCellValue(k, "yk4", 0);
                advBandedGridView1.SetRowCellValue(k, "jik4", 0);
                advBandedGridView1.SetRowCellValue(k, "ban4", 0);
                advBandedGridView1.SetRowCellValue(k, "Echapi4", 0);
                advBandedGridView1.SetRowCellValue(k, "chapi4", 0);
                //advBandedGridView1.SetRowCellValue(k, "light5", 0);
                //advBandedGridView1.SetRowCellValue(k, "weight5", 0);
                //advBandedGridView1.SetRowCellValue(k, "yk5", 0);
                //advBandedGridView1.SetRowCellValue(k, "jik5", 0);
                //advBandedGridView1.SetRowCellValue(k, "chapi5", 0);
                advBandedGridView1.SetRowCellValue(k, "lightTot", 0);
                advBandedGridView1.SetRowCellValue(k, "weightTot", 0);
                advBandedGridView1.SetRowCellValue(k, "ykTot", 0);
                advBandedGridView1.SetRowCellValue(k, "jikTot", 0);
                advBandedGridView1.SetRowCellValue(k, "BanTot", 0);
                advBandedGridView1.SetRowCellValue(k, "EchapiTot", 0);
                advBandedGridView1.SetRowCellValue(k, "chapiTot", 0);

            }
            //매출계획테이블 초기화
            for (int k = 0; k < 5; k++)
            {
                advBandedGridView2.SetRowCellValue(k, "jucha1", 0);
                advBandedGridView2.SetRowCellValue(k, "amount1", 0);
                advBandedGridView2.SetRowCellValue(k, "jucha2", 0);
                advBandedGridView2.SetRowCellValue(k, "amount2", 0);
                advBandedGridView2.SetRowCellValue(k, "jucha3", 0);
                advBandedGridView2.SetRowCellValue(k, "amount3", 0);
                advBandedGridView2.SetRowCellValue(k, "jucha4", 0);
                advBandedGridView2.SetRowCellValue(k, "amount4", 0);
                advBandedGridView2.SetRowCellValue(k, "jucha5", 0);
                advBandedGridView2.SetRowCellValue(k, "amount5", 0);
                advBandedGridView2.SetRowCellValue(k, "jucha6", 0);
                advBandedGridView2.SetRowCellValue(k, "amount6", 0);
                advBandedGridView2.SetRowCellValue(k, "juchaTot", 0);
                advBandedGridView2.SetRowCellValue(k, "amountTot", 0);
            }
        }
        private void simpleButton3_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            //}
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {

                FileInfo_1 fileInfo = new FileInfo_1("3");

                Cursor = Cursors.WaitCursor;
                string[] sPath = fileInfo.CheckFileInfo();
                Cursor = Cursors.Default;

                if (sPath != null)
                {
                    excelApp = new Excel.Application();
                    wb = excelApp.Workbooks.Open(sPath[0], 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);

                    //wb = excelApp.Workbooks.Open(@"c:\temp\영업실적목표.xlsx", 0, true);
                    ws = wb.Sheets["월 주간 목표"];

                    ws.Cells[1, 1] = DEYYMM.EditValue.ToString().Replace("-", "").Substring(4, 2);
                    ws.Cells[4, 2] = gridBand3.Caption;
                    ws.Cells[4, 8] = gridBand4.Caption;
                    ws.Cells[4, 14] = gridBand5.Caption;
                    ws.Cells[4, 20] = gridBand6.Caption;
                    //ws.Cells[4, 26] = gridBand59.Caption;
                    ws.Cells[4, 32] = gridBand7.Caption;

                    //#0001
                    for (int i = 0; i < advBandedGridView1.RowCount; i++) //매입목표
                    {
                        ws.Cells[i + 7, 2] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn1);
                        ws.Cells[i + 7, 3] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn3);
                        ws.Cells[i + 7, 4] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn2);
                        ws.Cells[i + 7, 5] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn4);
                        ws.Cells[i + 7, 6] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn47);
                        ws.Cells[i + 7, 7] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn5);

                        ws.Cells[i + 7, 8] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn6);
                        ws.Cells[i + 7, 9] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn8);
                        ws.Cells[i + 7, 10] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn7);
                        ws.Cells[i + 7, 11] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn9);
                        ws.Cells[i + 7, 12] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn48);
                        ws.Cells[i + 7, 13] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn10);

                        ws.Cells[i + 7, 14] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn11);
                        ws.Cells[i + 7, 15] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn13);
                        ws.Cells[i + 7, 16] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn12);
                        ws.Cells[i + 7, 17] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn14);
                        ws.Cells[i + 7, 18] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn49);
                        ws.Cells[i + 7, 19] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn15);

                        ws.Cells[i + 7, 20] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn16);
                        ws.Cells[i + 7, 21] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn18);
                        ws.Cells[i + 7, 22] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn17);
                        ws.Cells[i + 7, 23] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn19);
                        ws.Cells[i + 7, 24] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn50);
                        ws.Cells[i + 7, 25] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn20);

                        ws.Cells[i + 7, 26] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn40);
                        ws.Cells[i + 7, 27] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn42);
                        ws.Cells[i + 7, 28] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn41);
                        ws.Cells[i + 7, 29] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn43);
                        //ws.Cells[i + 7, 30] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn51);
                        ws.Cells[i + 7, 31] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn44);

                        ws.Cells[i + 7, 32] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn21);
                        ws.Cells[i + 7, 33] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn23);
                        ws.Cells[i + 7, 34] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn22);
                        ws.Cells[i + 7, 35] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn24);
                        ws.Cells[i + 7, 36] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn52);
                        ws.Cells[i + 7, 37] = advBandedGridView1.GetRowCellValue(i, bandedGridColumn25);
                    }


                    for (int i = 0; i < advBandedGridView2.RowCount; i++) //매출목표
                    {
                        ws.Cells[i + 20, 5] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn28);
                        ws.Cells[i + 20, 6] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn29);

                        ws.Cells[i + 20, 9] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn30);
                        ws.Cells[i + 20, 10] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn31);

                        ws.Cells[i + 20, 13] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn32);
                        ws.Cells[i + 20, 14] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn33);

                        ws.Cells[i + 20, 17] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn34);
                        ws.Cells[i + 20, 18] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn35);

                        ws.Cells[i + 20, 21] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn36);
                        ws.Cells[i + 20, 22] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn37);

                        ws.Cells[i + 20, 25] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn45);
                        ws.Cells[i + 20, 26] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn46);

                        ws.Cells[i + 20, 33] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn38);
                        ws.Cells[i + 20, 34] = advBandedGridView2.GetRowCellValue(i, bandedGridColumn39);

                    }


                    excelApp.Visible = true;
                }
            }
            finally
            {

            }
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = DEYYMM.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevMonth(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                DEYYMM.EditValue = sPrevDate;

                btSerch.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = DEYYMM.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextMonth(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                DEYYMM.EditValue = sNextDate;

                btSerch.PerformClick();
            }
        }
    }
}