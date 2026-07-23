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
using MySql.Data.MySqlClient;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class SetHoliday : DevExpress.XtraEditors.XtraForm
    {
        public SetHoliday()
        {
            InitializeComponent();
        }

        public string sYmd;

        private void SetHoliday_Load(object sender, EventArgs e)
        {
            CldrWorkingDate.EditValue = sYmd;
        }
        
        private void SetHoliday_Shown(object sender, EventArgs e)
        {
            GetHolidayInfo(sYmd);
        }

        private void GetHolidayInfo(string sYmd)
        {
            Cursor = Cursors.WaitCursor;

            string sDay = sYmd.Replace("-", "").Substring(0, 6);
            int iLastDayInMonth = DateTime.DaysInMonth(Convert.ToInt32(sDay.Substring(0, 4)), Convert.ToInt32(sDay.Substring(4, 2)));
            string sLastDayInMonth = sDay + iLastDayInMonth.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.PRDT_PLN_YMD AS HOLIDAY");
            strSql.AppendLine("   FROM PRDT_PLAN_MGT A ");
            strSql.AppendLine("  WHERE PRDT_PLN_YMD >= '" + sDay + "01' ");
            strSql.AppendLine("    AND PRDT_PLN_YMD <= '" + sLastDayInMonth + "' ");
            strSql.AppendLine("    AND HOLIDAY_YN = 'N' ");
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            
            Cursor = Cursors.Default;

            if (dt.Rows.Count > 0)
            {
                Cursor = Cursors.WaitCursor;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string sHoliday = dt.Rows[i]["HOLIDAY"]?.ToString();
                    dt.Rows[i]["HOLIDAY"] = sHoliday.Substring(0, 4) + "-" + sHoliday.Substring(4, 2) + "-" + sHoliday.Substring(6, 2);
                }

                GridWorkingDate.DataSource = dt;
                Cursor = Cursors.Default;
            }
            else if(dt.Rows.Count == 0)
            {
                SetHoliDate(sYmd);
            }

            Cursor = Cursors.Default;
        }

        private void SetHoliDate(string sYmd)
        {
            Cursor = Cursors.WaitCursor;

            DataTable dt = new DataTable();
            dt.TableName = "Table";
            dt.Columns.Add("HOLIDAY");
            GridWorkingDate.DataSource = dt;
            
            string sDay = sYmd.Replace("-", "").Substring(0, 8);
            int iLastDayInMonth = DateTime.DaysInMonth(Convert.ToInt32(sDay.Substring(0, 4)), Convert.ToInt32(sDay.Substring(4, 2)));

            int iStartDayOfMonth = Convert.ToInt32(sDay.Substring(0, 4) + sDay.Substring(4, 2) + "01");
            int iLastDayOfMonth = Convert.ToInt32(sDay.Substring(0, 4) + sDay.Substring(4, 2) + Convert.ToString(iLastDayInMonth));

            for(int i = iStartDayOfMonth; i <= iLastDayOfMonth; i++)
            {
                Cursor = Cursors.WaitCursor;

                string sWorkYmd = Convert.ToString(i);
                DateTime dtChk = new DateTime(Convert.ToInt32(sWorkYmd.Substring(0, 4)), Convert.ToInt32(sWorkYmd.Substring(4, 2)), Convert.ToInt32(sWorkYmd.Substring(6, 2)));
                var day = dtChk.DayOfWeek;
                
                switch (day)
                {
                    case DayOfWeek.Sunday:
                        GridViewWorkingDate.AddNewRow();
                        GridViewWorkingDate.SetFocusedRowCellValue("HOLIDAY", dtChk.ToString().Substring(0, 10));
                        break;
                }

                Cursor = Cursors.Default;
            }

            Cursor = Cursors.Default;
        }
        
        private void CldrWorkingDate_DateTimeChanged(object sender, EventArgs e)
        {
            string sSelectedDate = CldrWorkingDate.EditValue.ToString().Substring(0, 10);
            string sYyMm = sYmd.Substring(0, 7);

            if (sSelectedDate.Substring(0, 7).Equals(sYyMm))
            {
                for (int i = 0; i < GridViewWorkingDate.RowCount; i++)
                {
                    GridViewWorkingDate.FocusedRowHandle = i;
                    string sWorkingDate = GridViewWorkingDate.GetRowCellValue(i, "HOLIDAY")?.ToString();
                    if (sSelectedDate.Equals(sWorkingDate))
                    {
                        return;
                    }
                }
                GridViewWorkingDate.AddNewRow();
                GridViewWorkingDate.SetFocusedRowCellValue("HOLIDAY", sSelectedDate);
                GridViewWorkingDate.RefreshData();
                
            }
            else
            {
                XtraMessageBox.Show(sYmd.Substring(0, 4) + "년" + sYmd.Substring(6, 2) + "월에 해당하는 날짜만 선택하여주세요.");
            }
        }

        private void BtnConfirm_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                
                StringBuilder strSql = new StringBuilder();

                DataTable dtMerge = (DataTable)GridWorkingDate.DataSource;
                
                if(dtMerge.Rows.Count == 0)
                {
                    Cursor = Cursors.Default;
                    XtraMessageBox.Show("휴무일을 클릭해주세요.");
                    return;
                }

                if (dtMerge != null)
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    string sYyMm = sYmd.Replace("-", "").Substring(0, 6);

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" DELETE FROM PRDT_PLAN_MGT");
                    strSql.AppendLine("  WHERE PRDT_PLN_YMD >= '" + sYyMm + "01' ");
                    strSql.AppendLine("    AND PRDT_PLN_YMD <= '" + sYyMm + "31' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    string sDay = sYmd.Replace("-", "").Substring(0, 8);
                    int iLastDayInMonth = DateTime.DaysInMonth(Convert.ToInt32(sDay.Substring(0, 4)), Convert.ToInt32(sDay.Substring(4, 2)));

                    int iStartDayOfMonth = Convert.ToInt32(sDay.Substring(0, 4) + sDay.Substring(4, 2) + "01");
                    int iLastDayOfMonth = Convert.ToInt32(sDay.Substring(0, 4) + sDay.Substring(4, 2) + Convert.ToString(iLastDayInMonth));

                    for (int i = iStartDayOfMonth; i <= iLastDayOfMonth; i++)
                    {
                        string sWorkYmd = Convert.ToString(i).Replace("-", "").Substring(0, 8);
                        bool bChk = false;
                        for (int j = 0; j < dtMerge.Rows.Count; j++)
                        {
                            string sHoliyDay = dtMerge.Rows[j]["HOLIDAY"].ToString().Replace("-", "").Substring(0, 8);
                            if (sHoliyDay.Equals(sWorkYmd))
                            {
                                bChk = true;
                            }
                        }

                        if (bChk == true)
                        {
                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" INSERT INTO PRDT_PLAN_MGT ");
                            strSql.AppendLine("           ( ");
                            strSql.AppendLine("             PRDT_PLN_YMD ");
                            strSql.AppendLine("           , HOLIDAY_YN ");
                            strSql.AppendLine("           ) ");
                            strSql.AppendLine("      VALUES ");
                            strSql.AppendLine("      	  ( '" + sWorkYmd + "' ");
                            strSql.AppendLine("           , 'N' ");
                            strSql.AppendLine("           ) ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();
                        }
                        else if (bChk == false)
                        {
                            strSql.Clear();
                            strSql.AppendLine(" ");
                            strSql.AppendLine(" INSERT INTO PRDT_PLAN_MGT ");
                            strSql.AppendLine("           ( ");
                            strSql.AppendLine("             PRDT_PLN_YMD ");
                            strSql.AppendLine("           , HOLIDAY_YN ");
                            strSql.AppendLine("           ) ");
                            strSql.AppendLine("      VALUES ");
                            strSql.AppendLine("      	  ( '" + sWorkYmd + "' ");
                            strSql.AppendLine("           , 'Y' ");
                            strSql.AppendLine("           ) ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();
                        }
                    }
                    Cursor = Cursors.Default;

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    XtraMessageBox.Show("저장을 완료했습니다.");
                    Dispose();
                }
                
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;

                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewWorkingDate_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            string sWorkDate = GridViewWorkingDate.GetFocusedRowCellValue("HOLIDAY").ToString().Replace("-", "").Substring(0, 8);
            if (e.Clicks == 2)
            {
                //GridViewWorkingDate.DeleteRow(GridViewWorkingDate.FocusedRowHandle);
                DeleteSelectedRows(GridViewWorkingDate, sWorkDate);
            }
        }

        private void DeleteSelectedRows(DevExpress.XtraGrid.Views.Grid.GridView view, string sWorkDate)
        {
            DataTable dt = (DataTable)GridWorkingDate.DataSource;

            for(int i = 0; i < dt.Rows.Count; i++)
            {
                string sYYMMDD = dt.Rows[i]["HOLIDAY"].ToString().Replace("-", "").Substring(0, 8);
                if (sWorkDate.Equals(sYYMMDD))
                {
                    dt.Rows.RemoveAt(i);
                    return;
                }
            }

            GridWorkingDate.DataSource = dt;
            //if (view == null || view.SelectedRowsCount == 0) return;
            //DataRow[] rows = new DataRow[view.SelectedRowsCount];
            //for (int i = 0; i < view.SelectedRowsCount; i++)
            //    rows[i] = view.GetDataRow(view.GetSelectedRows()[i]);
            //view.BeginSort();
            //try
            //{
            //    foreach (DataRow row in rows)
            //        row.Delete();
            //}
            //finally
            //{
            //    view.EndSort();
            //}
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Dispose();
        }
    }
}