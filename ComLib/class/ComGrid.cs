using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DevExpress.XtraEditors;
using DevExpress.Data;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid;
using DevExpress.XtraEditors.Repository;
using DevExpress.XtraGrid.Columns;

namespace ComLib
{
    public static class ComGrid
    {
        public static DataSet GET_DATASET_FOR_SAVE(DataTable dt)
        {
            DataTable dtAdd = new DataTable();
            DataTable dtMod = new DataTable();

            foreach (DataRow dr in dt.Rows)
            {
                DataRowState drState = dr.RowState;
            }

            dtAdd = dt.Clone();
            dtMod = dt.Clone();

            dtAdd.TableName = "dtAdd";
            dtMod.TableName = "dtMod";

            int iAdd = 0;
            DataRow[] drAddRows = dt.Select(null, null, DataViewRowState.Added);
            foreach (DataRow dr in drAddRows)
            {
                dtAdd.ImportRow(dt.Rows[dt.Rows.IndexOf(drAddRows[iAdd])]);
                iAdd++;
            }

            int iMod = 0;
            DataRow[] drModRows = dt.Select(null, null, DataViewRowState.ModifiedCurrent);
            foreach (DataRow drMod in drModRows)
            {
                dtMod.ImportRow(dt.Rows[dt.Rows.IndexOf(drModRows[iMod])]);
                iMod++;
            }

            DataSet ds = new DataSet();

            ds.Tables.Add(dtAdd);
            ds.Tables.Add(dtMod);

            return ds;
        }

        public static DataSet GET_DATASET_FOR_MERGE(DataTable dt)
        {
            DataTable dtMerge = new DataTable();

            foreach (DataRow dr in dt.Rows)
            {
                DataRowState drState = dr.RowState;
            }

            dtMerge = dt.Clone();

            dtMerge.TableName = "dtMerge";

            int iAdd = 0;
            DataRow[] drAddRows = dt.Select(null, null, DataViewRowState.Added);
            foreach (DataRow dr in drAddRows)
            {
                dtMerge.ImportRow(dt.Rows[dt.Rows.IndexOf(drAddRows[iAdd])]);
                iAdd++;
            }

            int iMod = 0;
            DataRow[] drModRows = dt.Select(null, null, DataViewRowState.ModifiedCurrent);
            foreach (DataRow drMod in drModRows)
            {
                dtMerge.ImportRow(dt.Rows[dt.Rows.IndexOf(drModRows[iMod])]);
                iMod++;
            }

            DataSet ds = new DataSet();

            ds.Tables.Add(dtMerge);

            return ds;
        }

        public static void SetGridLookUpEdit(RepositoryItemGridLookUpEdit repositoryLookUp, DataTable dataTable, GridControl gridControl, GridColumn gridColumn, string sValue, string sDisplay, string sNullText)
        {
            GridView view = new GridView();
         
            repositoryLookUp.DataSource = dataTable;
            repositoryLookUp.ValueMember = sValue;
            repositoryLookUp.DisplayMember = sDisplay;

            gridControl.RepositoryItems.Add(repositoryLookUp);
            gridColumn.ColumnEdit = repositoryLookUp;
            repositoryLookUp.NullText = sNullText;

            repositoryLookUp.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            repositoryLookUp.PopupFilterMode = PopupFilterMode.Contains;
            repositoryLookUp.PopupView.OptionsFilter.AllowFilterIncrementalSearch = true;
            repositoryLookUp.ImmediatePopup = true;
            repositoryLookUp.View.OptionsView.ShowColumnHeaders = false;

        }

        public static void SetLookUpEdit(LookUpEdit lkup, DataTable dt, string sValue, string sDisplay, string sSetIdx)
        {
            lkup.Properties.DataSource = dt;
            lkup.Properties.DisplayMember = sDisplay;
            lkup.Properties.ValueMember = sValue; 

            if (sSetIdx.Equals("Y")) lkup.ItemIndex = 0;
        }

        public static bool GridDataSourceUpdate(DevExpress.XtraGrid.Views.Grid.GridView gvControl)
        {
            gvControl.CloseEditor();
            return gvControl.UpdateCurrentRow();
        }
    }
}
