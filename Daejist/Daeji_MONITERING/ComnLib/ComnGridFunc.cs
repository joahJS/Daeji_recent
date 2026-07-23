using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Daeji_MONITERING
{
    class SAFE_GARDRAIL
    {
        public static void SetInitGridRowColor(GridView view)
        { 
            view.OptionsFind.AllowFindPanel = false;
            view.IndicatorWidth = 40;
        }

        public static DataTable DeleteAllGridViewRows(GridControl grid)
        {
            DataTable dt = (DataTable)grid.DataSource;
            dt.Rows.Clear();

            return dt;
        }

        #region[GridView StripePattern 적용]
        public static void SettingGridViewRowPatternToStripe(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = Color.FromArgb(239, 240, 242);
                //e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
        #endregion[GridView StripePattern 적용]

        #region[GridView Indigator Number 적용]
        public static void SettingGridViewRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            /*
                해당 이벤트 적용 시 GridView.IndicatorWidth = 40으로 따로 적용해여야함
                현 메소드 적용 안됨
             */
            GridView L_View = sender as GridView;

            if (e.RowHandle < 0)
                e.Info.DisplayText = L_View.RowCount.ToString();
            else
                e.Info.DisplayText = (e.RowHandle + 1).ToString();

            e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

            int indiWidth = TextRenderer.MeasureText(L_View.RowCount.ToString(), null).Width;
            L_View.IndicatorWidth = 40;
        }
        #endregion[GridView Indigator Number 적용]

        #region[마지막 Row Focusing 여부]
        public static bool VerificateCheckLastRowFocusing(GridView view)
        {
            if (view.FocusedRowHandle != (view.RowCount - 1))
                return true;
            else
                return false;
        }
        #endregion[마지막 Row Focusing 여부]

        #region [GridView 기본세팅 한번에 하기]

        public static void GridStyleBasicSetting(GridView view)
        {
            view.CustomDrawRowIndicator += (sender, e) =>
            {
                if (e.RowHandle < 0)
                    e.Info.DisplayText = view.RowCount.ToString();
                else
                    e.Info.DisplayText = (e.RowHandle + 1).ToString();

                e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

                int indiWidth = TextRenderer.MeasureText(view.RowCount.ToString(), null).Width;
            };

            view.RowStyle += (sender, e) =>
            {
                if (e.RowHandle % 2 == 0)
                {
                    e.Appearance.BackColor = Color.FromArgb(239, 240, 242);
                }
            };

            //view.CustomDrawRowIndicator += (sender, e) =>
            //{
            //    if(e.Info.Kind == DevExpress.Utils.Drawing.IndicatorKind.Header)
            //    {
            //        //e.Appearance.DrawBackground(e.Cache, e.Bounds);
            //        e.Appearance.DrawString(e.Cache, "순번", e.Bounds);
            //        e.Handled = true;
            //    }
            //};

            view.OptionsNavigation.EnterMoveNextColumn = true;
            view.OptionsView.ColumnAutoWidth = false;
            view.OptionsView.ShowGroupPanel = false;
            view.OptionsFind.AllowFindPanel = false;
            view.IndicatorWidth = 40;           
            view.ColumnPanelRowHeight = 30;
        }

        public static void GridStyleForSelect(GridView view)
        {
            view.CustomDrawRowIndicator += (sender, e) =>
            {
                if (e.RowHandle < 0)
                    e.Info.DisplayText = view.RowCount.ToString();
                else
                    e.Info.DisplayText = (e.RowHandle + 1).ToString();

                e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;

                int indiWidth = TextRenderer.MeasureText(view.RowCount.ToString(), null).Width;
            };

            view.RowStyle += (sender, e) =>
            {
                if (e.RowHandle % 2 == 0)
                {
                    e.Appearance.BackColor = Color.FromArgb(239, 240, 242);
                }
            };

            view.OptionsView.ColumnAutoWidth = false;
            view.OptionsView.ShowGroupPanel = false;
            view.IndicatorWidth = 40;
            view.ColumnPanelRowHeight = 30;
        }

        #endregion

    }
}
