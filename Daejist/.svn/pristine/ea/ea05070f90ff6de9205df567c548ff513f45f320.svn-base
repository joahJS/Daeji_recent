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
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using ComLib;
using DevExpress.XtraGrid.Columns;
using System.Diagnostics;
using DevExpress.XtraPrinting;

namespace AccAdm
{
    public partial class GE002F00 : DevExpress.XtraEditors.XtraForm
    {
        public GE002F00()
        {
            InitializeComponent();
        }

        private string PROCEDURE_ID = "DP_GE002F00";

        private void GE002F00_Load(object sender, EventArgs e)
        {
            SetLoadFormLayout();
            ComnEtcFunc.gp_SetColorFocused(layoutControl2);
            ComnEtcFunc.SetDateToValue(DateYM);

            GridViewRetr.Appearance.FocusedRow.BackColor = Color.FromArgb(49, 106, 197);
            GridViewRetr2.Appearance.FocusedRow.BackColor = Color.FromArgb(49, 106, 197);
            GridViewRetr9.Appearance.FocusedRow.BackColor = Color.FromArgb(49, 106, 197);

            BtnRetr.PerformClick();
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { GridViewRetr, GridViewRetr2, GridViewRetr3, GridViewRetr4, GridViewRetr5, GridViewRetr6,
                                          GridViewRetr7, GridViewRetr8, GridViewRetr9, GridViewRetr10, GridViewRetr11, GridViewRetr12 };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
        }
        #endregion

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sDate = DateYM.EditValue?.ToString().Substring(0, 10);
            string sYY = sDate.Substring(0, 4);
            string sYM = sDate.Substring(0, 7);
            string sMM = sDate.Substring(5, 2);

            SetDate(sYY, sMM);

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "");
            dicParams.Add("YY", sYY);
            dicParams.Add("YM", sYM);
            dicParams.Add("MM", sMM);

            setGridRetrData(dicParams);
            SetGridRetr2Data(dicParams);
            SetGridRetr3Data(dicParams);
            SetGridRetr10Data(dicParams);
            SetGridRetr4Data(dicParams);
            SetGridRetr11Data(dicParams);
            SetGridRetr5Data(dicParams);
            SetGridRetr6Data(dicParams);
            SetGridRetr7Data(dicParams);
            SetGridRetr8Data(dicParams);
            SetGridRetr9Data(dicParams);
            SetGridRetr12Data(dicParams);
        }

        // 조회기간에 따른 날짜 세팅 메서드
        private void SetDate(string year, string month)
        {
            // 해당 (연도, 월)의 마지막 날짜
            int lastday = DateTime.DaysInMonth(Convert.ToInt32(year), Convert.ToInt32(month));

            // 2월
            if (lastday == 28)
            {
                GridColRet3_29.Visible = false;
                GridColRet3_30.Visible = false;
                GridColRet3_31.Visible = false;

                GridColRet4_29.Visible = false;
                GridColRet4_30.Visible = false;
                GridColRet4_31.Visible = false;

                GridColRet5_29.Visible = false;
                GridColRet5_30.Visible = false;
                GridColRet5_31.Visible = false;

                GridColRet6_29.Visible = false;
                GridColRet6_30.Visible = false;
                GridColRet6_31.Visible = false;

                GridColRet7_29.Visible = false;
                GridColRet7_30.Visible = false;
                GridColRet7_31.Visible = false;

                GridColRet8_29.Visible = false;
                GridColRet8_30.Visible = false;
                GridColRet8_31.Visible = false;

                GridColRet9_29.Visible = false;
                GridColRet9_30.Visible = false;
                GridColRet9_31.Visible = false;

                GridColRet10_29.Visible = false;
                GridColRet10_30.Visible = false;
                GridColRet10_31.Visible = false;

                GridColRet11_29.Visible = false;
                GridColRet11_30.Visible = false;
                GridColRet11_31.Visible = false;

                GridColRet12_29.Visible = false;
                GridColRet12_30.Visible = false;
                GridColRet12_31.Visible = false;
            }
            // 2월 (윤달)
            else if (lastday == 29)
            {
                GridColRet3_29.Visible = true;
                GridColRet3_30.Visible = false;
                GridColRet3_31.Visible = false;

                GridColRet4_29.Visible = true;
                GridColRet4_30.Visible = false;
                GridColRet4_31.Visible = false;

                GridColRet5_29.Visible = true;
                GridColRet5_30.Visible = false;
                GridColRet5_31.Visible = false;

                GridColRet6_29.Visible = true;
                GridColRet6_30.Visible = false;
                GridColRet6_31.Visible = false;

                GridColRet7_29.Visible = true;
                GridColRet7_30.Visible = false;
                GridColRet7_31.Visible = false;

                GridColRet8_29.Visible = true;
                GridColRet8_30.Visible = false;
                GridColRet8_31.Visible = false;

                GridColRet9_29.Visible = true;
                GridColRet9_30.Visible = false;
                GridColRet9_31.Visible = false;

                GridColRet10_29.Visible = true;
                GridColRet10_30.Visible = false;
                GridColRet10_31.Visible = false;

                GridColRet11_29.Visible = true;
                GridColRet11_30.Visible = false;
                GridColRet11_31.Visible = false;

                GridColRet12_29.Visible = true;
                GridColRet12_30.Visible = false;
                GridColRet12_31.Visible = false;
            }
            // 4월, 6월, 9월, 11월
            else if (lastday == 30)
            {
                GridColRet3_29.Visible = true;
                GridColRet3_30.Visible = true;
                GridColRet3_31.Visible = false;

                GridColRet4_29.Visible = true;
                GridColRet4_30.Visible = true;
                GridColRet4_31.Visible = false;

                GridColRet5_29.Visible = true;
                GridColRet5_30.Visible = true;
                GridColRet5_31.Visible = false;

                GridColRet6_29.Visible = true;
                GridColRet6_30.Visible = true;
                GridColRet6_31.Visible = false;

                GridColRet7_29.Visible = true;
                GridColRet7_30.Visible = true;
                GridColRet7_31.Visible = false;

                GridColRet8_29.Visible = true;
                GridColRet8_30.Visible = true;
                GridColRet8_31.Visible = false;

                GridColRet9_29.Visible = true;
                GridColRet9_30.Visible = true;
                GridColRet9_31.Visible = false;

                GridColRet10_29.Visible = true;
                GridColRet10_30.Visible = true;
                GridColRet10_31.Visible = false;

                GridColRet11_29.Visible = true;
                GridColRet11_30.Visible = true;
                GridColRet11_31.Visible = false;

                GridColRet12_29.Visible = true;
                GridColRet12_30.Visible = true;
                GridColRet12_31.Visible = false;
            }
            // 1월, 3월, 5월, 7월, 8월, 10월, 12월
            else if (lastday == 31)
            {
                GridColRet3_29.Visible = true;
                GridColRet3_30.Visible = true;
                GridColRet3_31.Visible = true;

                GridColRet4_29.Visible = true;
                GridColRet4_30.Visible = true;
                GridColRet4_31.Visible = true;

                GridColRet5_29.Visible = true;
                GridColRet5_30.Visible = true;
                GridColRet5_31.Visible = true;

                GridColRet6_29.Visible = true;
                GridColRet6_30.Visible = true;
                GridColRet6_31.Visible = true;

                GridColRet7_29.Visible = true;
                GridColRet7_30.Visible = true;
                GridColRet7_31.Visible = true;

                GridColRet8_29.Visible = true;
                GridColRet8_30.Visible = true;
                GridColRet8_31.Visible = true;

                GridColRet9_29.Visible = true;
                GridColRet9_30.Visible = true;
                GridColRet9_31.Visible = true;

                GridColRet10_29.Visible = true;
                GridColRet10_30.Visible = true;
                GridColRet10_31.Visible = true;

                GridColRet11_29.Visible = true;
                GridColRet11_30.Visible = true;
                GridColRet11_31.Visible = true;

                GridColRet12_29.Visible = true;
                GridColRet12_30.Visible = true;
                GridColRet12_31.Visible = true;
            }

            //
            setColIdx();
        }

        private void setColIdx()
        {
            GridColRet3_NM.VisibleIndex = 0;
            GridColRet3_01.VisibleIndex = 1;
            GridColRet3_02.VisibleIndex = 2;
            GridColRet3_03.VisibleIndex = 3;
            GridColRet3_04.VisibleIndex = 4;
            GridColRet3_05.VisibleIndex = 5;
            GridColRet3_06.VisibleIndex = 6;
            GridColRet3_07.VisibleIndex = 7;
            GridColRet3_08.VisibleIndex = 8;
            GridColRet3_09.VisibleIndex = 9;
            GridColRet3_10.VisibleIndex = 10;
            GridColRet3_11.VisibleIndex = 11;
            GridColRet3_12.VisibleIndex = 12;
            GridColRet3_13.VisibleIndex = 13;
            GridColRet3_14.VisibleIndex = 14;
            GridColRet3_15.VisibleIndex = 15;
            GridColRet3_16.VisibleIndex = 16;
            GridColRet3_17.VisibleIndex = 17;
            GridColRet3_18.VisibleIndex = 18;
            GridColRet3_19.VisibleIndex = 19;
            GridColRet3_20.VisibleIndex = 20;
            GridColRet3_21.VisibleIndex = 21;
            GridColRet3_22.VisibleIndex = 22;
            GridColRet3_23.VisibleIndex = 23;
            GridColRet3_24.VisibleIndex = 24;
            GridColRet3_25.VisibleIndex = 25;
            GridColRet3_26.VisibleIndex = 26;
            GridColRet3_27.VisibleIndex = 27;
            GridColRet3_28.VisibleIndex = 28;
            if (GridColRet3_29.Visible)
                GridColRet3_29.VisibleIndex = 30;
            if (GridColRet3_30.Visible)
                GridColRet3_30.VisibleIndex = 31;
            if (GridColRet3_31.Visible)
                GridColRet3_31.VisibleIndex = 32;
            if (GridColRet3_TOT.Visible)
                GridColRet3_TOT.VisibleIndex = 33;

            GridColRet4_NM.VisibleIndex = 0;
            GridColRet4_01.VisibleIndex = 1;
            GridColRet4_02.VisibleIndex = 2;
            GridColRet4_03.VisibleIndex = 3;
            GridColRet4_04.VisibleIndex = 4;
            GridColRet4_05.VisibleIndex = 5;
            GridColRet4_06.VisibleIndex = 6;
            GridColRet4_07.VisibleIndex = 7;
            GridColRet4_08.VisibleIndex = 8;
            GridColRet4_09.VisibleIndex = 9;
            GridColRet4_10.VisibleIndex = 10;
            GridColRet4_11.VisibleIndex = 11;
            GridColRet4_12.VisibleIndex = 12;
            GridColRet4_13.VisibleIndex = 13;
            GridColRet4_14.VisibleIndex = 14;
            GridColRet4_15.VisibleIndex = 15;
            GridColRet4_16.VisibleIndex = 16;
            GridColRet4_17.VisibleIndex = 17;
            GridColRet4_18.VisibleIndex = 18;
            GridColRet4_19.VisibleIndex = 19;
            GridColRet4_20.VisibleIndex = 20;
            GridColRet4_21.VisibleIndex = 21;
            GridColRet4_22.VisibleIndex = 22;
            GridColRet4_23.VisibleIndex = 23;
            GridColRet4_24.VisibleIndex = 24;
            GridColRet4_25.VisibleIndex = 25;
            GridColRet4_26.VisibleIndex = 26;
            GridColRet4_27.VisibleIndex = 27;
            GridColRet4_28.VisibleIndex = 28;
            if (GridColRet4_29.Visible)
                GridColRet4_29.VisibleIndex = 30;
            if (GridColRet4_30.Visible)
                GridColRet4_30.VisibleIndex = 31;
            if (GridColRet4_31.Visible)
                GridColRet4_31.VisibleIndex = 32;
            if (GridColRet4_TOT.Visible)
                GridColRet4_TOT.VisibleIndex = 33;

            GridColRet5_NM.VisibleIndex = 0;
            GridColRet5_01.VisibleIndex = 1;
            GridColRet5_02.VisibleIndex = 2;
            GridColRet5_03.VisibleIndex = 3;
            GridColRet5_04.VisibleIndex = 4;
            GridColRet5_05.VisibleIndex = 5;
            GridColRet5_06.VisibleIndex = 6;
            GridColRet5_07.VisibleIndex = 7;
            GridColRet5_08.VisibleIndex = 8;
            GridColRet5_09.VisibleIndex = 9;
            GridColRet5_10.VisibleIndex = 10;
            GridColRet5_11.VisibleIndex = 11;
            GridColRet5_12.VisibleIndex = 12;
            GridColRet5_13.VisibleIndex = 13;
            GridColRet5_14.VisibleIndex = 14;
            GridColRet5_15.VisibleIndex = 15;
            GridColRet5_16.VisibleIndex = 16;
            GridColRet5_17.VisibleIndex = 17;
            GridColRet5_18.VisibleIndex = 18;
            GridColRet5_19.VisibleIndex = 19;
            GridColRet5_20.VisibleIndex = 20;
            GridColRet5_21.VisibleIndex = 21;
            GridColRet5_22.VisibleIndex = 22;
            GridColRet5_23.VisibleIndex = 23;
            GridColRet5_24.VisibleIndex = 24;
            GridColRet5_25.VisibleIndex = 25;
            GridColRet5_26.VisibleIndex = 26;
            GridColRet5_27.VisibleIndex = 27;
            GridColRet5_28.VisibleIndex = 28;
            if (GridColRet5_29.Visible)
                GridColRet5_29.VisibleIndex = 30;
            if (GridColRet5_30.Visible)
                GridColRet5_30.VisibleIndex = 31;
            if (GridColRet5_31.Visible)
                GridColRet5_31.VisibleIndex = 32;
            if (GridColRet5_TOT.Visible)
                GridColRet5_TOT.VisibleIndex = 33;

            GridColRet6_NM.VisibleIndex = 0;
            GridColRet6_01.VisibleIndex = 1;
            GridColRet6_02.VisibleIndex = 2;
            GridColRet6_03.VisibleIndex = 3;
            GridColRet6_04.VisibleIndex = 4;
            GridColRet6_05.VisibleIndex = 5;
            GridColRet6_06.VisibleIndex = 6;
            GridColRet6_07.VisibleIndex = 7;
            GridColRet6_08.VisibleIndex = 8;
            GridColRet6_09.VisibleIndex = 9;
            GridColRet6_10.VisibleIndex = 10;
            GridColRet6_11.VisibleIndex = 11;
            GridColRet6_12.VisibleIndex = 12;
            GridColRet6_13.VisibleIndex = 13;
            GridColRet6_14.VisibleIndex = 14;
            GridColRet6_15.VisibleIndex = 15;
            GridColRet6_16.VisibleIndex = 16;
            GridColRet6_17.VisibleIndex = 17;
            GridColRet6_18.VisibleIndex = 18;
            GridColRet6_19.VisibleIndex = 19;
            GridColRet6_20.VisibleIndex = 20;
            GridColRet6_21.VisibleIndex = 21;
            GridColRet6_22.VisibleIndex = 22;
            GridColRet6_23.VisibleIndex = 23;
            GridColRet6_24.VisibleIndex = 24;
            GridColRet6_25.VisibleIndex = 25;
            GridColRet6_26.VisibleIndex = 26;
            GridColRet6_27.VisibleIndex = 27;
            GridColRet6_28.VisibleIndex = 28;
            if (GridColRet6_29.Visible)
                GridColRet6_29.VisibleIndex = 30;
            if (GridColRet6_30.Visible)
                GridColRet6_30.VisibleIndex = 31;
            if (GridColRet6_31.Visible)
                GridColRet6_31.VisibleIndex = 32;
            if (GridColRet6_TOT.Visible)
                GridColRet6_TOT.VisibleIndex = 33;

            GridColRet7_NM.VisibleIndex = 0;
            GridColRet7_01.VisibleIndex = 1;
            GridColRet7_02.VisibleIndex = 2;
            GridColRet7_03.VisibleIndex = 3;
            GridColRet7_04.VisibleIndex = 4;
            GridColRet7_05.VisibleIndex = 5;
            GridColRet7_06.VisibleIndex = 6;
            GridColRet7_07.VisibleIndex = 7;
            GridColRet7_08.VisibleIndex = 8;
            GridColRet7_09.VisibleIndex = 9;
            GridColRet7_10.VisibleIndex = 10;
            GridColRet7_11.VisibleIndex = 11;
            GridColRet7_12.VisibleIndex = 12;
            GridColRet7_13.VisibleIndex = 13;
            GridColRet7_14.VisibleIndex = 14;
            GridColRet7_15.VisibleIndex = 15;
            GridColRet7_16.VisibleIndex = 16;
            GridColRet7_17.VisibleIndex = 17;
            GridColRet7_18.VisibleIndex = 18;
            GridColRet7_19.VisibleIndex = 19;
            GridColRet7_20.VisibleIndex = 20;
            GridColRet7_21.VisibleIndex = 21;
            GridColRet7_22.VisibleIndex = 22;
            GridColRet7_23.VisibleIndex = 23;
            GridColRet7_24.VisibleIndex = 24;
            GridColRet7_25.VisibleIndex = 25;
            GridColRet7_26.VisibleIndex = 26;
            GridColRet7_27.VisibleIndex = 27;
            GridColRet7_28.VisibleIndex = 28;
            if (GridColRet7_29.Visible)
                GridColRet7_29.VisibleIndex = 30;
            if (GridColRet7_30.Visible)
                GridColRet7_30.VisibleIndex = 31;
            if (GridColRet7_31.Visible)
                GridColRet7_31.VisibleIndex = 32;
            if (GridColRet7_TOT.Visible)
                GridColRet7_TOT.VisibleIndex = 33;

            GridColRet8_NM.VisibleIndex = 0;
            GridColRet8_01.VisibleIndex = 1;
            GridColRet8_02.VisibleIndex = 2;
            GridColRet8_03.VisibleIndex = 3;
            GridColRet8_04.VisibleIndex = 4;
            GridColRet8_05.VisibleIndex = 5;
            GridColRet8_06.VisibleIndex = 6;
            GridColRet8_07.VisibleIndex = 7;
            GridColRet8_08.VisibleIndex = 8;
            GridColRet8_09.VisibleIndex = 9;
            GridColRet8_10.VisibleIndex = 10;
            GridColRet8_11.VisibleIndex = 11;
            GridColRet8_12.VisibleIndex = 12;
            GridColRet8_13.VisibleIndex = 13;
            GridColRet8_14.VisibleIndex = 14;
            GridColRet8_15.VisibleIndex = 15;
            GridColRet8_16.VisibleIndex = 16;
            GridColRet8_17.VisibleIndex = 17;
            GridColRet8_18.VisibleIndex = 18;
            GridColRet8_19.VisibleIndex = 19;
            GridColRet8_20.VisibleIndex = 20;
            GridColRet8_21.VisibleIndex = 21;
            GridColRet8_22.VisibleIndex = 22;
            GridColRet8_23.VisibleIndex = 23;
            GridColRet8_24.VisibleIndex = 24;
            GridColRet8_25.VisibleIndex = 25;
            GridColRet8_26.VisibleIndex = 26;
            GridColRet8_27.VisibleIndex = 27;
            GridColRet8_28.VisibleIndex = 28;
            if (GridColRet8_29.Visible)
                GridColRet8_29.VisibleIndex = 30;
            if (GridColRet8_30.Visible)
                GridColRet8_30.VisibleIndex = 31;
            if (GridColRet8_31.Visible)
                GridColRet8_31.VisibleIndex = 32;
            if (GridColRet8_TOT.Visible)
                GridColRet8_TOT.VisibleIndex = 33;

            GridColRet9_NM.VisibleIndex = 0;
            GridColRet9_01.VisibleIndex = 1;
            GridColRet9_02.VisibleIndex = 2;
            GridColRet9_03.VisibleIndex = 3;
            GridColRet9_04.VisibleIndex = 4;
            GridColRet9_05.VisibleIndex = 5;
            GridColRet9_06.VisibleIndex = 6;
            GridColRet9_07.VisibleIndex = 7;
            GridColRet9_08.VisibleIndex = 8;
            GridColRet9_09.VisibleIndex = 9;
            GridColRet9_10.VisibleIndex = 10;
            GridColRet9_11.VisibleIndex = 11;
            GridColRet9_12.VisibleIndex = 12;
            GridColRet9_13.VisibleIndex = 13;
            GridColRet9_14.VisibleIndex = 14;
            GridColRet9_15.VisibleIndex = 15;
            GridColRet9_16.VisibleIndex = 16;
            GridColRet9_17.VisibleIndex = 17;
            GridColRet9_18.VisibleIndex = 18;
            GridColRet9_19.VisibleIndex = 19;
            GridColRet9_20.VisibleIndex = 20;
            GridColRet9_21.VisibleIndex = 21;
            GridColRet9_22.VisibleIndex = 22;
            GridColRet9_23.VisibleIndex = 23;
            GridColRet9_24.VisibleIndex = 24;
            GridColRet9_25.VisibleIndex = 25;
            GridColRet9_26.VisibleIndex = 26;
            GridColRet9_27.VisibleIndex = 27;
            GridColRet9_28.VisibleIndex = 28;
            if (GridColRet9_29.Visible)
                GridColRet9_29.VisibleIndex = 30;
            if (GridColRet9_30.Visible)
                GridColRet9_30.VisibleIndex = 31;
            if (GridColRet9_31.Visible)
                GridColRet9_31.VisibleIndex = 32;
            if (GridColRet9_TOT.Visible)
                GridColRet9_TOT.VisibleIndex = 33;
            if (GridColRet9_FINTOT.Visible)
                GridColRet9_FINTOT.VisibleIndex = 34;

            GridColRet10_NM.VisibleIndex = 0;
            GridColRet10_01.VisibleIndex = 1;
            GridColRet10_02.VisibleIndex = 2;
            GridColRet10_03.VisibleIndex = 3;
            GridColRet10_04.VisibleIndex = 4;
            GridColRet10_05.VisibleIndex = 5;
            GridColRet10_06.VisibleIndex = 6;
            GridColRet10_07.VisibleIndex = 7;
            GridColRet10_08.VisibleIndex = 8;
            GridColRet10_09.VisibleIndex = 9;
            GridColRet10_10.VisibleIndex = 10;
            GridColRet10_11.VisibleIndex = 11;
            GridColRet10_12.VisibleIndex = 12;
            GridColRet10_13.VisibleIndex = 13;
            GridColRet10_14.VisibleIndex = 14;
            GridColRet10_15.VisibleIndex = 15;
            GridColRet10_16.VisibleIndex = 16;
            GridColRet10_17.VisibleIndex = 17;
            GridColRet10_18.VisibleIndex = 18;
            GridColRet10_19.VisibleIndex = 19;
            GridColRet10_20.VisibleIndex = 20;
            GridColRet10_21.VisibleIndex = 21;
            GridColRet10_22.VisibleIndex = 22;
            GridColRet10_23.VisibleIndex = 23;
            GridColRet10_24.VisibleIndex = 24;
            GridColRet10_25.VisibleIndex = 25;
            GridColRet10_26.VisibleIndex = 26;
            GridColRet10_27.VisibleIndex = 27;
            GridColRet10_28.VisibleIndex = 28;
            if (GridColRet10_29.Visible)
                GridColRet10_29.VisibleIndex = 30;
            if (GridColRet10_30.Visible)
                GridColRet10_30.VisibleIndex = 31;
            if (GridColRet10_31.Visible)
                GridColRet10_31.VisibleIndex = 32;
            if (GridColRet10_TOT.Visible)
                GridColRet10_TOT.VisibleIndex = 33;

            GridColRet11_NM.VisibleIndex = 0;
            GridColRet11_01.VisibleIndex = 1;
            GridColRet11_02.VisibleIndex = 2;
            GridColRet11_03.VisibleIndex = 3;
            GridColRet11_04.VisibleIndex = 4;
            GridColRet11_05.VisibleIndex = 5;
            GridColRet11_06.VisibleIndex = 6;
            GridColRet11_07.VisibleIndex = 7;
            GridColRet11_08.VisibleIndex = 8;
            GridColRet11_09.VisibleIndex = 9;
            GridColRet11_10.VisibleIndex = 10;
            GridColRet11_11.VisibleIndex = 11;
            GridColRet11_12.VisibleIndex = 12;
            GridColRet11_13.VisibleIndex = 13;
            GridColRet11_14.VisibleIndex = 14;
            GridColRet11_15.VisibleIndex = 15;
            GridColRet11_16.VisibleIndex = 16;
            GridColRet11_17.VisibleIndex = 17;
            GridColRet11_18.VisibleIndex = 18;
            GridColRet11_19.VisibleIndex = 19;
            GridColRet11_20.VisibleIndex = 20;
            GridColRet11_21.VisibleIndex = 21;
            GridColRet11_22.VisibleIndex = 22;
            GridColRet11_23.VisibleIndex = 23;
            GridColRet11_24.VisibleIndex = 24;
            GridColRet11_25.VisibleIndex = 25;
            GridColRet11_26.VisibleIndex = 26;
            GridColRet11_27.VisibleIndex = 27;
            GridColRet11_28.VisibleIndex = 28;
            if (GridColRet11_29.Visible)
                GridColRet11_29.VisibleIndex = 30;
            if (GridColRet11_30.Visible)
                GridColRet11_30.VisibleIndex = 31;
            if (GridColRet11_31.Visible)
                GridColRet11_31.VisibleIndex = 32;
            if (GridColRet11_TOT.Visible)
                GridColRet11_TOT.VisibleIndex = 33;

            GridColRet12_NM.VisibleIndex = 0;
            GridColRet12_01.VisibleIndex = 1;
            GridColRet12_02.VisibleIndex = 2;
            GridColRet12_03.VisibleIndex = 3;
            GridColRet12_04.VisibleIndex = 4;
            GridColRet12_05.VisibleIndex = 5;
            GridColRet12_06.VisibleIndex = 6;
            GridColRet12_07.VisibleIndex = 7;
            GridColRet12_08.VisibleIndex = 8;
            GridColRet12_09.VisibleIndex = 9;
            GridColRet12_10.VisibleIndex = 10;
            GridColRet12_11.VisibleIndex = 11;
            GridColRet12_12.VisibleIndex = 12;
            GridColRet12_13.VisibleIndex = 13;
            GridColRet12_14.VisibleIndex = 14;
            GridColRet12_15.VisibleIndex = 15;
            GridColRet12_16.VisibleIndex = 16;
            GridColRet12_17.VisibleIndex = 17;
            GridColRet12_18.VisibleIndex = 18;
            GridColRet12_19.VisibleIndex = 19;
            GridColRet12_20.VisibleIndex = 20;
            GridColRet12_21.VisibleIndex = 21;
            GridColRet12_22.VisibleIndex = 22;
            GridColRet12_23.VisibleIndex = 23;
            GridColRet12_24.VisibleIndex = 24;
            GridColRet12_25.VisibleIndex = 25;
            GridColRet12_26.VisibleIndex = 26;
            GridColRet12_27.VisibleIndex = 27;
            GridColRet12_28.VisibleIndex = 28;
            if (GridColRet12_29.Visible)
                GridColRet12_29.VisibleIndex = 30;
            if (GridColRet12_30.Visible)
                GridColRet12_30.VisibleIndex = 31;
            if (GridColRet12_31.Visible)
                GridColRet12_31.VisibleIndex = 32;
            if (GridColRet12_TOT.Visible)
                GridColRet12_TOT.VisibleIndex = 33;
        }


        #region 년집계
        private void setGridRetrData(Dictionary<string, string> dicParams)
        {
            try
            {
                dicParams["CMD"] = "LIST1";
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewRetr_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            if (e.Column.FieldName == "EMP_NM")
            {
                var dr1 = GridViewRetr.GetDataRow(e.RowHandle1);
                var dr2 = GridViewRetr.GetDataRow(e.RowHandle2);

                e.Merge = dr1["EMP_NM"].ToString().Equals(dr2["EMP_NM"].ToString());
                return;
            }
        }

        private void GridViewRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);

            GridView view = sender as GridView;

            // 홀수 행은 전체 행이 원하는대로 되는데, 짝수 행은 선택된 셀만 동작을 하여 추가함
            // 폼 로드 시에 설정하면 안 되었고, 로 스타일에서 설정하여야 정상 동작
            //view.OptionsView.EnableAppearanceEvenRow = true;

            if (((e.State & DevExpress.XtraGrid.Views.Base.GridRowCellState.Focused) != 0) &&
               ((e.State & DevExpress.XtraGrid.Views.Base.GridRowCellState.GridFocused) != 0))
            {
                e.Appearance.Assign(view.PaintAppearance.FocusedRow);
            }
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }
        #endregion

        #region 휴무일수
        private void SetGridRetr2Data(Dictionary<string, string> dicParams)
        {
            try
            {
                dicParams["CMD"] = "LIST2_1";
                DataTable dt1 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                dicParams["CMD"] = "LIST2_2";
                DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if (dt2 != null && dt2.Rows.Count > 0)
                {
                    string sCnt = dt2.Rows[0]["CNT"]?.ToString();

                    if(int.TryParse(sCnt, out int iCnt))
                    {
                        int maxColumnIndex = GridViewRetr2.Columns.Count-1;

                        for(int i = maxColumnIndex; i> 3; i--)
                        {
                            GridViewRetr2.Columns.RemoveAt(i);
                        }

                        int tempcnt = 1;

                        for(int i = tempcnt; i<=iCnt; i++)
                        {
                            GridColumn GridCol = new GridColumn();

                            GridCol.AppearanceCell.Options.UseTextOptions = true;
                            GridCol.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                            GridCol.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                            GridCol.AppearanceHeader.Options.UseTextOptions = true;
                            GridCol.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
                            GridCol.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
                            GridCol.Caption = "사용"+ i;
                            GridCol.FieldName = "USE_"+ i;
                            GridCol.OptionsColumn.AllowEdit = false;
                            GridCol.OptionsColumn.AllowFocus = false;
                            GridCol.OptionsColumn.AllowMerge = DevExpress.Utils.DefaultBoolean.False;
                            GridCol.Visible = true;
                            GridCol.Width = 80;

                            GridViewRetr2.Columns.Add(GridCol);

                            dt1.Columns.Add("USE_" + i);
                        }
                    }


                    int iCnt2 = 1;
                    int iCnt3 = 1;

                    string sCol = "G";

                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        string sSeq = dt1.Rows[i]["SEQ"]?.ToString();
                        string sEmpid = dt1.Rows[i]["EMP_ID"]?.ToString();

                        dt1.Rows[i]["COL"] = sCol;

                        if (sSeq.Equals("0"))
                        {
                            if(dt2 != null)
                            {
                                for(int j = 0; j < dt2.Rows.Count; j++)
                                {
                                    string sEmpid2 = dt2.Rows[j]["EMPID"]?.ToString();
                                    string sBasdt = dt2.Rows[j]["BASDT"]?.ToString();

                                    if (sEmpid.Equals(sEmpid2))
                                    {
                                        dt1.Rows[i]["USE_" + iCnt2++] = sBasdt;
                                    }
                                }
                            }
                        }
                        else if (sSeq.Equals("1"))
                        {
                            if(dt2 != null)
                            {
                                for (int j = 0; j < dt2.Rows.Count; j++)
                                {
                                    string sEmpid2 = dt2.Rows[j]["EMPID"]?.ToString();
                                    string sUscnt = dt2.Rows[j]["USCNT"]?.ToString();

                                    if (sEmpid.Equals(sEmpid2))
                                    {
                                        dt1.Rows[i]["USE_" + iCnt3++] = sUscnt;
                                    }
                                }
                            }
                            if (sCol.Equals("G"))
                            {
                                sCol = "W";
                            }
                            else
                            {
                                sCol = "G";
                            }
                        }

                        iCnt2 = 1;
                        iCnt3 = 1;
                    }

                    GridRetr2.DataSource = dt1;
                }

            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewRetr2_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            string sEmpNm1 = GridViewRetr2.GetRowCellValue(e.RowHandle1, "EMP_NM")?.ToString();
            string sEmpNm2 = GridViewRetr2.GetRowCellValue(e.RowHandle2, "EMP_NM")?.ToString();

            var val1 = GridViewRetr2.GetRowCellValue(e.RowHandle1, e.Column);
            var val2 = GridViewRetr2.GetRowCellValue(e.RowHandle2, e.Column);

            if (e.Column.FieldName.Equals("DEPT_NM") || e.Column.FieldName.Equals("EMP_NM") || e.Column.FieldName.Equals("YNCNT") || e.Column.FieldName.Equals("JCNT"))
            {
                e.Merge = (sEmpNm1.Equals(sEmpNm2)) && (val1.Equals(val2));
                e.Handled = true;
                return;
            }
        }

        private void GridViewRetr9_CellMerge(object sender, CellMergeEventArgs e)
        {
            e.Merge = false;

            string sEmpNm1 = GridViewRetr9.GetRowCellValue(e.RowHandle1, "EMP_NM")?.ToString();
            string sEmpNm2 = GridViewRetr9.GetRowCellValue(e.RowHandle2, "EMP_NM")?.ToString();

            var val1 = GridViewRetr9.GetRowCellValue(e.RowHandle1, e.Column);
            var val2 = GridViewRetr9.GetRowCellValue(e.RowHandle2, e.Column);

            if (e.Column.FieldName.Equals("EMP_NM") || e.Column.FieldName.Equals("FINTOT"))
            {
                e.Merge = (sEmpNm1.Equals(sEmpNm2)) && (val1.Equals(val2));
                e.Handled = true;
                return;
            }
        }

        private void GridViewRetr2_RowStyle(object sender, RowStyleEventArgs e)
        {
            string sCol = GridViewRetr2.GetRowCellValue(e.RowHandle, "COL")?.ToString();

            if (!string.IsNullOrEmpty(sCol) && sCol.Equals("G"))
            {
                if (DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName != "DevExpress Dark Style")
                {
                    e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
                }
                else if (DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName == "DevExpress Dark Style")
                {
                    e.Appearance.BackColor = SystemColors.ControlDark;
                }
            }

            GridView view = sender as GridView;

            // 홀수 행은 전체 행이 원하는대로 되는데, 짝수 행은 선택된 셀만 동작을 하여 추가함
            // 폼 로드 시에 설정하면 안 되었고, 로 스타일에서 설정하여야 정상 동작
            //view.OptionsView.EnableAppearanceEvenRow = false;

            if (((e.State & DevExpress.XtraGrid.Views.Base.GridRowCellState.Focused) != 0) &&
               ((e.State & DevExpress.XtraGrid.Views.Base.GridRowCellState.GridFocused) != 0))
            {
                e.Appearance.Assign(view.PaintAppearance.FocusedRow);
            }
        }
        #endregion

        #region 특근
        private void SetGridRetr10Data(Dictionary<string, string> dicParams)
        {
            try
            {
                dicParams["CMD"] = "LIST4";
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr10.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 월휴무
        private void SetGridRetr3Data(Dictionary<string, string> dicParams)
        {
            try
            {
                dicParams["CMD"] = "LIST3";
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr3.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 병가
        private void SetGridRetr4Data(Dictionary<string, string> dicParams)
        {
            try
            {
                dicParams["CMD"] = "LIST5";
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr4.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 국경일
        private void SetGridRetr11Data(Dictionary<string, string> dicParams)
        {
            try
            {
                dicParams["CMD"] = "LIST6";
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr11.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 반차
        private void SetGridRetr5Data(Dictionary<string, string> dicParams)
        {
            try
            {
                dicParams["CMD"] = "LIST7";
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr5.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 외출
        private void SetGridRetr6Data(Dictionary<string, string> dicParams)
        {
            try
            {
                dicParams["CMD"] = "LIST8";
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr6.DataSource = dt;
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 지각
        private void SetGridRetr7Data(Dictionary<string, string> dicParams)
        {
            try
            {
                dicParams["CMD"] = "LIST9";
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr7.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 조퇴
        private void SetGridRetr8Data(Dictionary<string, string> dicParams)
        {
            try
            {
                dicParams["CMD"] = "LIST10";
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr8.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 잔업
        private void SetGridRetr9Data(Dictionary<string, string> dicParams)
        {
            try
            {
                dicParams["CMD"] = "LIST11";
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr9.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 조출
        private void SetGridRetr12Data(Dictionary<string, string> dicParams)
        {
            try
            {
                dicParams["CMD"] = "LIST12";
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr12.DataSource = dt;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void DateYM_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void GE002F00_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.Escape)
                BtnClose.PerformClick();
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = DateYM.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevMonth(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                DateYM.EditValue = sPrevDate;

                BtnRetr.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = DateYM.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextMonth(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                DateYM.EditValue = sNextDate;

                BtnRetr.PerformClick();
            }
        }

        private void GE002F00_TextChanged(object sender, EventArgs e)
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

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string FileName = string.Empty;
            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                string sFileNM = "월근태 리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    XlsxExportOptions options = new XlsxExportOptions();
                    options.ExportMode = XlsxExportMode.SingleFilePageByPage;
                    //CompositeLink link = new CompositeLink();

                    if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage1"))
                    {
                       
                        GridRetr.ExportToXlsx(FileName, options);
                        Process.Start(FileName);
                    }
                    else if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage2"))
                    {
                        GridRetr2.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                    else if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage3"))
                    {
                        GridRetr3.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                    else if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage4"))
                    {
                        GridRetr4.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                    else if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage5"))
                    {
                        GridRetr5.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                    else if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage6"))
                    {
                        GridRetr6.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                    else if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage7"))
                    {
                        GridRetr7.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                    else if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage8"))
                    {
                        GridRetr8.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                    else if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage9"))
                    {
                        GridRetr9.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                    else if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage10"))
                    {
                        GridRetr10.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                    else if (xtraTabControl1.SelectedTabPage.Name.ToString().Equals("xtraTabPage11"))
                    {
                        GridRetr11.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                    else
                    {
                        GridRetr12.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                }
                fileDlg.Dispose();
            }
            catch (Exception ex)
            {
                if (ex.Message.Equals("Error Opening file"))
                {
                    //파일이 열려있음 체크
                    MessageBox.Show(((ex.InnerException).InnerException).Message);
                }
            }
        }

    }
}
