using DevExpress.XtraGrid.Views.Grid;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccAdm
{
    class ComnEventFunc
    {
        #region[GridView StripePattern 적용]
        public static void SettingGridViewRowPatternToStripe(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
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
        }
        #endregion[GridView Indigator Number 적용]
    }
}
