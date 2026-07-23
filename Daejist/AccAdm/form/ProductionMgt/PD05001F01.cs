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
using System.Net;
using System.IO;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
/*
* 작성일자 : 2021-03-25
* 작성자 : 고혜성
*        -> 기존 PD02001F01 파일을 복사한 파일이며 검수내역관리 재구성
*           
* ---------------------HISTORY-----------------------
* 
*/
namespace AccAdm
{
    public partial class PD05001F01 : DevExpress.XtraEditors.XtraForm
    {
        public PD05001F01()
        {
            InitializeComponent();
        }

        public GridView[] arrGrdView;
        private void PD02001F01_Load(object sender, EventArgs e)
        {
            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            arrGrdView = new GridView[] { GridViewRetr, GridViewRetr2 };
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

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout2.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl2.RestoreLayoutFromXml(sFile);
            }

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout3.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl3.RestoreLayoutFromXml(sFile);
            }
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
                string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);

                if (string.IsNullOrEmpty(sYmdFrom))
                {
                    Cursor = Cursors.Default;
                    XtraMessageBox.Show("등록일자를 바르게 입력하세요.");
                    DateEditFrom.SelectAll();
                    DateEditFrom.Focus();
                    return;
                }
                else if (string.IsNullOrEmpty(sYmdTo))
                {
                    Cursor = Cursors.Default;
                    XtraMessageBox.Show("등록일자를 바르게 입력하세요.");
                    DateEditTo.SelectAll();
                    DateEditTo.Focus();
                    return;
                }

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("DATE_F", sYmdFrom);
                dicParams.Add("DATE_T", sYmdTo);
                dicParams.Add("FIND_IDX", CboFindSbj.SelectedIndex.ToString());
                dicParams.Add("FIND_WORD", TxtFindWord.EditValue?.ToString().Trim());
                dicParams.Add("ITNL_YN", RdgbItnlYn.EditValue?.ToString());

                GridControl grdList = new GridControl();
                GridView grdView = new GridView();
                DataTable dt = new DataTable();
                if (TabControl.SelectedTabPage == TabPageGumsu)
                {
                    grdList = GridRetr;
                    grdView = GridViewRetr;
                    dt = GetInfo(dicParams);
                }
                else
                {
                    grdList = GridRetr2;
                    grdView = GridViewRetr2;
                    dt = GetSummary(dicParams);
                }


                grdList.DataSource = dt;
                if (grdView.RowCount > 0)
                {
                    grdView.Focus();
                }
                else if (grdView.RowCount == 0)
                {
                    if (string.IsNullOrEmpty(TxtFindWord.EditValue?.ToString()))
                    {
                        TxtFindWord.SelectAll();
                        TxtFindWord.Focus();
                    }
                    else
                    {
                        DateEditFrom.SelectAll();
                        DateEditFrom.Focus();
                    }
                }

                Cursor = Cursors.Default;
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }


        private DataTable GetInfo(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendFormat("\r\n ");
            strSql.AppendFormat("\r\n SELECT A1.JUNPYOID ");
            strSql.AppendFormat("\r\n      , A2.ISPT_NO ");
            strSql.AppendFormat("\r\n      , A1.J_DATE ");
            strSql.AppendFormat("\r\n      , DATE_FORMAT(A2.ENT_DT, '%Y-%m-%d') AS ENT_DT ");
            strSql.AppendFormat("\r\n      , B1.USRNM AS CUSER ");
            strSql.AppendFormat("\r\n      , A1.J_BNUM ");
            strSql.AppendFormat("\r\n      , A1.GUBUN1 ");
            strSql.AppendFormat("\r\n      , A2.DEALER_CD ");
            strSql.AppendFormat("\r\n      , A2.DEALER_NM ");
            strSql.AppendFormat("\r\n      , A2.WEIGHT ");
            strSql.AppendFormat("\r\n      , A2.CHAGAM ");
            strSql.AppendFormat("\r\n      , A1.J_STATE ");
            strSql.AppendFormat("\r\n      , CASE WHEN A1.KERATYPE = '입고' THEN A1.ICHAGAM ELSE A1.OCHAGAM END AS WGT_ADMT ");
            strSql.AppendFormat("\r\n      , A2.ITNL_YN ");
            strSql.AppendFormat("\r\n      , A2.ISPT_OPN ");
            strSql.AppendFormat("\r\n      , A2.IMG_CNT ");
            strSql.AppendFormat("\r\n   FROM MESURING A1 ");
            strSql.AppendFormat("\r\n   LEFT JOIN MESURE_ISPT_INFO A2 ");
            strSql.AppendFormat("\r\n     ON A1.JUNPYOID = A2.JUNPYOID ");
            strSql.AppendFormat("\r\n   LEFT JOIN ZUSRLST B1 ");
            strSql.AppendFormat("\r\n     ON A2.ENT_ID = B1.USRCD ");
            strSql.AppendFormat("\r\n  WHERE A2.JUNPYOID IS NOT NULL ");
            strSql.AppendFormat("\r\n    AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n    AND (('{0}' = '' AND 1 = 1) ", dicParams["FIND_WORD"]);
            strSql.AppendFormat("\r\n         OR ");
            strSql.AppendFormat("\r\n         ('{0}' = '0' AND A2.DEALER_NM LIKE '%{1}%') ", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            strSql.AppendFormat("\r\n         OR ");
            strSql.AppendFormat("\r\n         ('{0}' = '1' AND A1.GUBUN1 LIKE '%{1}%') ", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            strSql.AppendFormat("\r\n         OR ");
            strSql.AppendFormat("\r\n         ('{0}' = '2' AND A1.GUMSUBIGO LIKE '%{1}%') ", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            strSql.AppendFormat("\r\n         OR ");
            strSql.AppendFormat("\r\n         ('{0}' = '3' AND A1.J_STATE LIKE '%{1}%') ", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            strSql.AppendFormat("\r\n         OR ");
            strSql.AppendFormat("\r\n         ('{0}' = '4' AND A2.ISPT_OPN LIKE '%{1}%') ", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            strSql.AppendFormat("\r\n         OR ");
            strSql.AppendFormat("\r\n         ('{0}' = '5' AND B1.USRNM LIKE '%{1}%') ", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            strSql.AppendFormat("\r\n         OR ");
            strSql.AppendFormat("\r\n         ('{0}' = '6' AND A1.J_BNUM LIKE '%{1}%') )", dicParams["FIND_IDX"], dicParams["FIND_WORD"]);
            strSql.AppendFormat("\r\n    AND (('{0}' = 'ALL' AND 1 = 1) ", dicParams["ITNL_YN"]);
            strSql.AppendFormat("\r\n         OR ");
            strSql.AppendFormat("\r\n         ('{0}' = 'Y' AND A2.ITNL_YN = 'Y') ", dicParams["ITNL_YN"]);
            strSql.AppendFormat("\r\n         OR ");
            strSql.AppendFormat("\r\n         ('{0}' <> 'Y' AND A2.ITNL_YN <> 'Y') )", dicParams["ITNL_YN"]);
            strSql.AppendFormat("\r\n  ORDER BY A1.J_DATE, A1.SUN ");

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private DataTable GetSummary(Dictionary<string, string> dicParams)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendFormat("\r\n ");
            strSql.AppendFormat("\r\n SELECT * ");
            strSql.AppendFormat("\r\n   FROM ( ");
            strSql.AppendFormat("\r\n          SELECT Y1.OPN_GB ");
            strSql.AppendFormat("\r\n               , Y1.GB ");
            strSql.AppendFormat("\r\n               , Y1.MAIPCHERID ");
            strSql.AppendFormat("\r\n               , Y1.DEALER_NM ");
            strSql.AppendFormat("\r\n               , Y1.WGT ");
            strSql.AppendFormat("\r\n               , Y2.SUM_CHAGAM ");
            strSql.AppendFormat("\r\n               , Y1.IN_CNT ");
            strSql.AppendFormat("\r\n               , Y1.CHAGAM_CNT ");
            strSql.AppendFormat("\r\n               , Y1.CHAGAM_WGT ");
            strSql.AppendFormat("\r\n               , IFNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT ");
            strSql.AppendFormat("\r\n               , Y1.ITNL_YN ");
            strSql.AppendFormat("\r\n               , Y3.OPN_DATE ");
            strSql.AppendFormat("\r\n               , Y3.OPN_RMK ");
            strSql.AppendFormat("\r\n               , '{0}' AS DATE_T ", dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n            FROM ( ");
            strSql.AppendFormat("\r\n                   SELECT '1' AS OPN_GB ");
            strSql.AppendFormat("\r\n                        , '스크랩' AS GB ");
            strSql.AppendFormat("\r\n                        , A1.MAIPCHERID ");
            strSql.AppendFormat("\r\n                        , A2.DEALER_NM ");
            strSql.AppendFormat("\r\n                        , SUM(A1.IWEIGHT) AS WGT ");
            strSql.AppendFormat("\r\n                        , IFNULL(A3.CHAGAM, IFNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM ");
            strSql.AppendFormat("\r\n                        , COUNT(*) AS IN_CNT ");
            strSql.AppendFormat("\r\n                        , IFNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT ");
            strSql.AppendFormat("\r\n                        , SUM(A1.ICHAGAM) AS CHAGAM_WGT ");
            strSql.AppendFormat("\r\n                        , CASE WHEN IFNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN ");
            strSql.AppendFormat("\r\n                     FROM MESURING A1 ");
            strSql.AppendFormat("\r\n                     LEFT JOIN ACC_DEALER_CD A2 ");
            strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A2.DEALER_CD ");
            strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.DEALER_CD ");
            strSql.AppendFormat("\r\n                                      , SUM(X1.CHAGAM) AS CHAGAM ");
            strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT ");
            strSql.AppendFormat("\r\n                                   FROM MESURE_ISPT_INFO X1 ");
            strSql.AppendFormat("\r\n                                  GROUP BY X1.DEALER_CD ) A3 ");
            strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A3.DEALER_CD ");
            strSql.AppendFormat("\r\n                      AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.MAIPCHERID ");
            strSql.AppendFormat("\r\n                                      , COUNT(*) AS CHAGAM_CNT ");
            strSql.AppendFormat("\r\n                                   FROM MESURING X1 ");
            strSql.AppendFormat("\r\n                                  WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n                                    AND X1.ICHAGAM > 0 ");
            strSql.AppendFormat("\r\n                                  GROUP BY X1.MAIPCHERID ) A4 ");
            strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A4.MAIPCHERID ");
            strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.JUNPYOID, X1.CHAGAM ");
            strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            strSql.AppendFormat("\r\n                                  WHERE X1.CHAGAM > 0 ");
            strSql.AppendFormat("\r\n                                  LIMIT 1 ) A5 ");
            strSql.AppendFormat("\r\n                       ON A1.JUNPYOID = A5.JUNPYOID ");
            strSql.AppendFormat("\r\n                     LEFT JOIN JAJAE B1 ");
            strSql.AppendFormat("\r\n                       ON A1.J_SERIAL = B1.J_SERIAL ");
            strSql.AppendFormat("\r\n                    WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n                      AND A1.KERATYPE = '입고' ");
            strSql.AppendFormat("\r\n                      AND B1.DAEGUBUN IN ('고철A', '고철B') ");
            strSql.AppendFormat("\r\n                    GROUP BY A1.MAIPCHERID ");
            strSql.AppendFormat("\r\n                 ) Y1 ");
            strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            strSql.AppendFormat("\r\n                            , SUM(IFNULL(X3.CHAGAM, IFNULL(X2.CHAGAM, IFNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            strSql.AppendFormat("\r\n                         LEFT JOIN ( SELECT Y1.JUNPYOID ");
            strSql.AppendFormat("\r\n                                          , Y1.CHAGAM ");
            strSql.AppendFormat("\r\n                                       FROM MESURING_SEQ Y1 ");
            strSql.AppendFormat("\r\n                                      WHERE Y1.CHAGAM > 0 ");
            strSql.AppendFormat("\r\n                                      LIMIT 1 ) X2 ");
            strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X2.JUNPYOID ");
            strSql.AppendFormat("\r\n                         LEFT JOIN MESURE_ISPT_INFO X3 ");
            strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X3.JUNPYOID ");
            strSql.AppendFormat("\r\n                         LEFT JOIN JAJAE X4 ");
            strSql.AppendFormat("\r\n                           ON X1.J_SERIAL = X4.J_SERIAL ");
            strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            strSql.AppendFormat("\r\n                          AND X4.DAEGUBUN IN ('고철A', '고철B') ");
            strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y2 ");
            strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y2.MAIPCHERID ");
            strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.DEALER_CD ");
            strSql.AppendFormat("\r\n                             , X1.OPN_GB ");
            strSql.AppendFormat("\r\n                             , X1.OPN_DATE ");
            strSql.AppendFormat("\r\n                             , X1.OPN_RMK ");
            strSql.AppendFormat("\r\n                          FROM MESURE_OPN_HISTORY X1 ");
            strSql.AppendFormat("\r\n                         WHERE X1.OPN_DATE = '{0}' ) Y3  ", dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y3.DEALER_CD  ");
            strSql.AppendFormat("\r\n             AND Y1.OPN_GB = Y3.OPN_GB ");
            strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            strSql.AppendFormat("\r\n                            , COUNT(*) AS RETURN_CNT ");
            strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            strSql.AppendFormat("\r\n                          AND X1.J_SERIAL = 4049042 #스크랩반품 ");
            strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y4 ");
            strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y4.MAIPCHERID ");
            strSql.AppendFormat("\r\n           UNION ALL ");
            strSql.AppendFormat("\r\n          SELECT Y1.OPN_GB ");
            strSql.AppendFormat("\r\n               , Y1.GB ");
            strSql.AppendFormat("\r\n               , Y1.MAIPCHERID ");
            strSql.AppendFormat("\r\n               , Y1.DEALER_NM ");
            strSql.AppendFormat("\r\n               , Y1.WGT ");
            strSql.AppendFormat("\r\n               , Y2.SUM_CHAGAM ");
            strSql.AppendFormat("\r\n               , Y1.IN_CNT ");
            strSql.AppendFormat("\r\n               , Y1.CHAGAM_CNT ");
            strSql.AppendFormat("\r\n               , Y1.CHAGAM_WGT ");
            strSql.AppendFormat("\r\n               , IFNULL(Y4.RETURN_CNT, 0) AS RETURN_CNT ");
            strSql.AppendFormat("\r\n               , Y1.ITNL_YN ");
            strSql.AppendFormat("\r\n               , Y3.OPN_DATE ");
            strSql.AppendFormat("\r\n               , Y3.OPN_RMK ");
            strSql.AppendFormat("\r\n               , '{0}' AS DATE_T ", dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n            FROM ( ");
            strSql.AppendFormat("\r\n                   SELECT '2' AS OPN_GB ");
            strSql.AppendFormat("\r\n                        , '슈레더' AS GB ");
            strSql.AppendFormat("\r\n                        , A1.MAIPCHERID ");
            strSql.AppendFormat("\r\n                        , A2.DEALER_NM ");
            strSql.AppendFormat("\r\n                        , SUM(A1.IWEIGHT) AS WGT ");
            strSql.AppendFormat("\r\n                        , IFNULL(A3.CHAGAM, IFNULL(SUM(A1.ICHAGAM), 0)) AS CHAGAM ");
            strSql.AppendFormat("\r\n                        , COUNT(*) AS IN_CNT ");
            strSql.AppendFormat("\r\n                        , IFNULL(A4.CHAGAM_CNT, 0) AS CHAGAM_CNT ");
            strSql.AppendFormat("\r\n                        , SUM(A1.ICHAGAM) AS CHAGAM_WGT ");
            strSql.AppendFormat("\r\n                        , CASE WHEN IFNULL(A3.ITNL_CNT, 0) > 0 THEN '유' ELSE '무' END AS ITNL_YN ");
            strSql.AppendFormat("\r\n                     FROM MESURING A1 ");
            strSql.AppendFormat("\r\n                     LEFT JOIN ACC_DEALER_CD A2 ");
            strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A2.DEALER_CD ");
            strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.DEALER_CD ");
            strSql.AppendFormat("\r\n                                      , SUM(X1.CHAGAM) AS CHAGAM ");
            strSql.AppendFormat("\r\n                                      , COUNT(CASE WHEN X1.ITNL_YN = 'Y' THEN 1 END) AS ITNL_CNT ");
            strSql.AppendFormat("\r\n                                   FROM MESURE_ISPT_INFO X1 ");
            strSql.AppendFormat("\r\n                                  GROUP BY X1.DEALER_CD ) A3 ");
            strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A3.DEALER_CD ");
            strSql.AppendFormat("\r\n                      AND A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.MAIPCHERID ");
            strSql.AppendFormat("\r\n                                      , COUNT(*) AS CHAGAM_CNT ");
            strSql.AppendFormat("\r\n                                   FROM MESURING X1 ");
            strSql.AppendFormat("\r\n                                  WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n                                    AND X1.ICHAGAM > 0 ");
            strSql.AppendFormat("\r\n                                  GROUP BY X1.MAIPCHERID ) A4 ");
            strSql.AppendFormat("\r\n                       ON A1.MAIPCHERID = A4.MAIPCHERID ");
            strSql.AppendFormat("\r\n                     LEFT JOIN ( SELECT X1.JUNPYOID, X1.CHAGAM ");
            strSql.AppendFormat("\r\n                                   FROM MESURING_SEQ X1 ");
            strSql.AppendFormat("\r\n                                  WHERE X1.CHAGAM > 0 ");
            strSql.AppendFormat("\r\n                                  LIMIT 1 ) A5 ");
            strSql.AppendFormat("\r\n                       ON A1.JUNPYOID = A5.JUNPYOID ");
            strSql.AppendFormat("\r\n                     LEFT JOIN JAJAE B1 ");
            strSql.AppendFormat("\r\n                       ON A1.J_SERIAL = B1.J_SERIAL ");
            strSql.AppendFormat("\r\n                    WHERE A1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n                      AND A1.KERATYPE = '입고' ");
            strSql.AppendFormat("\r\n                      AND B1.DAEGUBUN IN ('슈레더') ");
            strSql.AppendFormat("\r\n                    GROUP BY A1.MAIPCHERID ");
            strSql.AppendFormat("\r\n                 ) Y1 ");
            strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            strSql.AppendFormat("\r\n                            , SUM(IFNULL(X3.CHAGAM, IFNULL(X2.CHAGAM, IFNULL(X1.ICHAGAM, 0)))) AS SUM_CHAGAM ");
            strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            strSql.AppendFormat("\r\n                         LEFT JOIN ( SELECT Y1.JUNPYOID ");
            strSql.AppendFormat("\r\n                                          , Y1.CHAGAM ");
            strSql.AppendFormat("\r\n                                       FROM MESURING_SEQ Y1 ");
            strSql.AppendFormat("\r\n                                      WHERE Y1.CHAGAM > 0 ");
            strSql.AppendFormat("\r\n                                      LIMIT 1 ) X2 ");
            strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X2.JUNPYOID ");
            strSql.AppendFormat("\r\n                         LEFT JOIN MESURE_ISPT_INFO X3 ");
            strSql.AppendFormat("\r\n                           ON X1.JUNPYOID = X3.JUNPYOID ");
            strSql.AppendFormat("\r\n                         LEFT JOIN JAJAE X4 ");
            strSql.AppendFormat("\r\n                           ON X1.J_SERIAL = X4.J_SERIAL ");
            strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            strSql.AppendFormat("\r\n                          AND X4.DAEGUBUN IN ('슈레더') ");
            strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y2 ");
            strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y2.MAIPCHERID ");
            strSql.AppendFormat("\r\n            LEFT JOIN ( SELECT X1.DEALER_CD ");
            strSql.AppendFormat("\r\n                             , X1.OPN_GB ");
            strSql.AppendFormat("\r\n                             , X1.OPN_DATE ");
            strSql.AppendFormat("\r\n                             , X1.OPN_RMK ");
            strSql.AppendFormat("\r\n                          FROM MESURE_OPN_HISTORY X1 ");
            strSql.AppendFormat("\r\n                         WHERE X1.OPN_DATE = '{0}' ) Y3  ", dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y3.DEALER_CD  ");
            strSql.AppendFormat("\r\n             AND Y1.OPN_GB = Y3.OPN_GB ");
            strSql.AppendFormat("\r\n            LEFT JOIN (SELECT X1.MAIPCHERID ");
            strSql.AppendFormat("\r\n                            , COUNT(*) AS RETURN_CNT ");
            strSql.AppendFormat("\r\n                         FROM MESURING X1 ");
            strSql.AppendFormat("\r\n                        WHERE X1.J_DATE BETWEEN '{0}' AND '{1}' ", dicParams["DATE_F"], dicParams["DATE_T"]);
            strSql.AppendFormat("\r\n                          AND X1.KERATYPE = '입고' ");
            strSql.AppendFormat("\r\n                          AND X1.J_SERIAL = 5050042 #슈레더반품 ");
            strSql.AppendFormat("\r\n                        GROUP BY X1.MAIPCHERID ) Y4 ");
            strSql.AppendFormat("\r\n              ON Y1.MAIPCHERID = Y4.MAIPCHERID ");
            strSql.AppendFormat("\r\n        ) Z1 ");
            strSql.AppendFormat("\r\n  ORDER BY Z1.OPN_GB ASC, REPLACE(Z1.DEALER_NM, '(주)', '') ASC ");

            return ComnEtcFunc.GetDataTable(DBConn.dbCon, strSql.ToString(), dicParams);
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            ComnEtcFunc.ExportExcelFile(string.Format("{0}_", this.Text), GridRetr);
        }

        public string RST_DT;
        private void BtnReport_Click(object sender, EventArgs e)
        {
            //PD02001F02 frm = new PD02001F02();
            //frm.P_PD02001F01 = this;
            //if(frm.ShowDialog() != DialogResult.OK)
            //{
            //    return;
            //}

            //DateTime dt = DateTime.Parse(RST_DT);
            //int iWeekCnt = GetWeekCnt(dt);

            //DateTime FirstDateOfWeek = GetFirstDateOfWeek(dt.Year, iWeekCnt);
            //DateTime LastDateOfWeek = FirstDateOfWeek.AddDays(6);
            //string sWeek = string.Format("{0}   ~   {1}   {2}월 주차{3}", FirstDateOfWeek.ToString("yyyy-MM-dd"), LastDateOfWeek.ToString("yyyy-MM-dd"), FirstDateOfWeek.Month, CalcWeekNumberFromDate(dt.Year, dt.Month, dt.Day).ToString());
            
            //Dictionary<string, string> dicParams = new Dictionary<string, string>();
            //dicParams.Add("DATE", sWeek);

            
        }

        // 내가 만든 것.
        private static int GetWeekCnt(DateTime dt)
        {
            int week = Enum.GetValues(typeof(DayOfWeek)).Length;
            int dayOffset = (int)dt.AddDays(-(dt.Day - 1)).DayOfWeek;
            int weekCnt = (dt.Day + dayOffset) / week;
            weekCnt += ((dt.Day + dayOffset) % week) > 0 ? 1 : 0;
            return weekCnt;
        }
        
        // 이건 해당 년도에 주차 수를 구한다.
        public int GetWeeksOfYear(DateTime date)
        {
            System.Globalization.CultureInfo cult_info = System.Globalization.CultureInfo.CreateSpecificCulture("ko");
            System.Globalization.Calendar cal = cult_info.Calendar;
            int weekNo = cal.GetWeekOfYear(date, cult_info.DateTimeFormat.CalendarWeekRule, cult_info.DateTimeFormat.FirstDayOfWeek);
            int week1day = cal.GetWeekOfYear(date.AddDays(-(date.Day + 1)), cult_info.DateTimeFormat.CalendarWeekRule, cult_info.DateTimeFormat.FirstDayOfWeek);
            return weekNo - week1day + 1;
        }

        /// <summary>
        /// 특정 주차 시작일 구하기
        /// </summary>
        /// <param name="year">연도</param>
        /// <param name="week">주차</param>
        /// <returns>특정 주차 시작일</returns>
        public DateTime GetFirstDateOfWeek(int year, int week)
        {
            DateTime firstDateOfYear = new DateTime(year, 1, 1);
            DateTime firstDateOfFirstWeek = firstDateOfYear.AddDays(7 - (int)(firstDateOfYear.DayOfWeek) + 1);
            return firstDateOfFirstWeek.AddDays(7 * (week - 1));
        }

        private static int CalcWeekNumberFromDate(int year, int month, int day)
        {
            int iWeekNumber = 0;
            DateTime dTime = new DateTime(year, month, day);
            string FirstOfMonth = DateTime.Now.ToString("yyyy-MM-01");
            DateTime dt = new DateTime(); DateTime.TryParse(FirstOfMonth, out dt);
            System.Globalization.CultureInfo myCl = new System.Globalization.CultureInfo("ko-KR");
            System.Globalization.Calendar myCal = myCl.Calendar;
            System.Globalization.CalendarWeekRule myCWR = myCl.DateTimeFormat.CalendarWeekRule;
            DayOfWeek myFirstDOW = myCl.DateTimeFormat.FirstDayOfWeek;
            
            //오늘 몇주차
            int WeekOfToday = myCal.GetWeekOfYear(dTime, myCWR, myFirstDOW);
            
            //오늘이 있는달 첫날짜 주차
            int WeekOfFirstday = myCal.GetWeekOfYear(dt, myCWR, myFirstDOW);
            int WeekOfMonth = WeekOfToday - WeekOfFirstday;
            iWeekNumber = WeekOfMonth + 1;

            return iWeekNumber;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void PD02001F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnExcel.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void RdgbItnlYn_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void CboFindSbj_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if(e.FocusedRowHandle < 0)
            {
                if(GridImg.DataSource != null)
                {
                    DataTable dt = (DataTable)GridImg.DataSource;
                    dt.Rows.Clear();
                    GridImg.DataSource = dt;
                }
                return;
            }

            //불필요한 FTP 서버의 접속을 줄이기 위하여 이미지카운트가 0 인경우 아래 로직 수행되지 않도록 구현
            string sImgCnt = GridViewRetr.GetFocusedRowCellValue(GridCol1ImgCnt)?.ToString();
            int iImgCnt = string.IsNullOrEmpty(sImgCnt) ? 0 : Convert.ToInt32(sImgCnt);

            if (iImgCnt == 0)
            {
                //이미지 그리드 초기화
                if (GridImg.DataSource != null)
                {
                    DataTable dt = (DataTable)GridImg.DataSource;
                    dt.Rows.Clear();
                    GridImg.DataSource = dt;
                }
                return;
            }

            string sENT_DT = GridViewRetr.GetFocusedRowCellValue(GridCol1EntDt)?.ToString().Replace("-", "");
            string sJUNPYOID = GridViewRetr.GetFocusedRowCellValue(GridCol1JunpyoID)?.ToString();

            GetImagesFromFTP(sENT_DT, sJUNPYOID);

        }

        private void GetImagesFromFTP(string sEntDt, string sJunpyoID)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                //string sInitDir = string.Format(@"ftp://192.168.0.202/Gumsu_Images/{0}/{1}/{2}/{3}", sEntDt.Substring(0, 4), sEntDt.Substring(4, 2), sEntDt.Substring(6, 2), sJunpyoID);
                string sInitDir = string.Format(@"ftp://"+ComnEtcFunc.FTP_IP+"/Gumsu_Images/{0}/{1}/{2}/{3}", sEntDt.Substring(0, 4), sEntDt.Substring(4, 2), sEntDt.Substring(6, 2), sJunpyoID);

                FtpWebRequest req1 = (FtpWebRequest)WebRequest.Create(sInitDir);
                string user = ComnEtcFunc.FTP_USER;
                string pw = ComnEtcFunc.FTP_PW;
                req1.Credentials = new NetworkCredential(user, pw);
                req1.Method = WebRequestMethods.Ftp.ListDirectory;

                string[] filesInDirectory = null;
                Dictionary<string, Image> dicImages = new Dictionary<string, Image>();
                using (FtpWebResponse req1Res = (FtpWebResponse)req1.GetResponse())
                {
                    StreamReader reader1 = new StreamReader(req1Res.GetResponseStream());
                    string strData = reader1.ReadToEnd();
                    //폴더 내 파일이름
                    filesInDirectory = strData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    reader1.Close();

                    foreach (string filePath in filesInDirectory)
                    {
                        string[] filesCopy = filePath.Split('\\');
                        dicImages.Add(filesCopy[filesCopy.Length - 1], ComnEtcFunc.DownloadFTPFile(string.Format(@"{0}\{1}", sInitDir, filePath), user, pw));
                    }
                }

                DataTable dt = new DataTable();
                dt.TableName = "Table1";
                dt.Columns.Add("IMAGE", typeof(byte[]));
                dt.Columns.Add("FILE_NAME");

                foreach (KeyValuePair<string, Image> item in dicImages)
                {
                    DataRow row = dt.NewRow();
                    row["FILE_NAME"] = item.Key;
                    row["IMAGE"] = ComnEtcFunc.ImageToByteArray(item.Value);
                    dt.Rows.Add(row);
                }

                GridImg.DataSource = dt;

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                return;
            }
        }

        private void TabControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            if(e.Page == TabPageGumsu)
            {
                LayoutFindIdx.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutFindWord.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutItnl.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LayoutBtnSave.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
            else
            {
                LayoutFindIdx.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutFindWord.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutItnl.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LayoutBtnSave.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }

            BtnRetr.PerformClick();
        }

        private void PD02001F01_TextChanged(object sender, EventArgs e)
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
                layoutControl2.SaveLayoutToXml(path + @"\" + this.Name + "_Layout2.xaml");
                layoutControl3.SaveLayoutToXml(path + @"\" + this.Name + "_Layout3.xaml");

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (TabControl.SelectedTabPage != TabPageSummary)
                return;

            try
            {
                Cursor = Cursors.WaitCursor;

                DataTable dtPrv = (DataTable)GridRetr2.DataSource;
                DataTable dt = dtPrv.GetChanges(DataRowState.Modified);
                
                if(dt == null)
                {
                    Cursor = Cursors.Default;
                    XtraMessageBox.Show("리스트에 총평을 입력하세요.");
                    return;
                }

                string sDealerCd = string.Empty;
                string sOpnGb = string.Empty;
                //bool bYn = false;
                //foreach(DataRow row in dt.Rows)
                //{
                //    sDealerCd = row["MAIPCHERID"].ToString();
                //    sOpnGb = row["OPN_GB"].ToString();
                //    string sRMK = row["OPN_RMK"].ToString();

                //    if (string.IsNullOrEmpty(sRMK))
                //    {
                //        bYn = true;
                //        Cursor = Cursors.Default;
                //        XtraMessageBox.Show("총평을 입력하세요.");
                //    }
                //}

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                int iCnt = 0;
                foreach(DataRow row in dt.Rows)
                {
                    string sDEALER_CD = row["MAIPCHERID"].ToString();
                    string sOPN_GB = row["OPN_GB"].ToString();
                    string sOPN_DATE = row["OPN_DATE"].ToString();
                    string sDATE_T = row["DATE_T"].ToString();
                    string sOPN_RMK = row["OPN_RMK"].ToString();

                    sOPN_DATE = string.IsNullOrEmpty(sOPN_DATE) ? sDATE_T : sOPN_DATE;

                    if (iCnt++ == 0)
                    {
                        sDealerCd = sDEALER_CD;
                        sOpnGb = sOPN_GB;
                    }
                    
                    strSql.Clear();
                    strSql.AppendFormat(" ");
                    strSql.AppendFormat(" INSERT INTO MESURE_OPN_HISTORY ");
                    strSql.AppendFormat("           ( DEALER_CD ");
                    strSql.AppendFormat("           , OPN_GB ");
                    strSql.AppendFormat("           , OPN_DATE ");
                    strSql.AppendFormat("           , OPN_RMK ");
                    strSql.AppendFormat("           , ENT_ID ");
                    strSql.AppendFormat("           , ENT_DT ) ");
                    strSql.AppendFormat("     VALUES( {0} ", sDEALER_CD); //DEALER_CD
                    strSql.AppendFormat("           , '{0}' ", sOPN_GB); //OPN_GB
                    strSql.AppendFormat("           , '{0}' ", sOPN_DATE); //OPN_DATE
                    strSql.AppendFormat("           , '{0}' ", sOPN_RMK); //OPN_RMK
                    strSql.AppendFormat("           , {0} ", FmMainToolBar2.UserID); //ENT_ID
                    strSql.AppendFormat("           , NOW() ) "); //ENT_DT
                    strSql.AppendFormat("        ON DUPLICATE KEY UPDATE ");
                    strSql.AppendFormat("           OPN_RMK = '{0}' ", sOPN_RMK); //OPN_RMK
                    strSql.AppendFormat("         , MFY_ID = {0} ", FmMainToolBar2.UserID); //MFY_ID
                    strSql.AppendFormat("         , MFY_DT = NOW() "); //MFY_DT

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }
                
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                Cursor = Cursors.Default;
                XtraMessageBox.Show("저장을 완료했습니다.");

                BtnRetr.PerformClick();
                GridViewRetr2.FocusedRowHandle = GridViewRetr2.LocateByDisplayText(0, GridCol2OpnGb, sOpnGb);
                GridViewRetr2.FocusedRowHandle = GridViewRetr2.LocateByDisplayText(GridViewRetr2.FocusedRowHandle, GridCol2MaipcherID, sDealerCd);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewRetr2_CustomDrawCell(object sender, DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
        {
            if(e.Column == GridCol2ItnlGB)
            {
                string sVal = e.CellValue?.ToString() ?? string.Empty;
                if (sVal.Equals("유"))
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
                else
                {
                    e.Appearance.BackColor = Color.PaleGreen;
                }
            }
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}