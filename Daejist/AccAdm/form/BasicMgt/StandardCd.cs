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
using DevExpress.XtraGrid.Views.Grid;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class StandardCd : DevExpress.XtraEditors.XtraForm
    {
        public StandardCd()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void StandardCd_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            //
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            SetCombo();

            DataTable dt = GetLookUpData("");
            ComGrid.SetLookUpEdit(LkupLength, dt, "CD", "CD", "Y");
            LkupLength.Properties.PopulateColumns();
            LkupLength.Properties.Columns[1].Visible = false;

            arrGrdView = new GridView[] { GridViewRetr };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            BtnRetr_Click(null, null);
        }

        private void SetCombo()
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;

            }
            StringBuilder strSql = new StringBuilder();
            strSql.AppendLine("");
            strSql.AppendLine("SELECT Model as Items");
            strSql.AppendLine(" FROM LENGTHS");
            strSql.AppendLine("GROUP BY Model");
            //strSql.AppendLine("GROUP BY Items");
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string Item = dt.Rows[i]["Items"].ToString();
                CboGrade.Properties.Items.Add(Item);
            }
        }

        private DataTable GetLookUpData(string sModel)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");
            strSql.AppendLine("WITH ITEM_INFO AS (");
            strSql.AppendLine(" SELECT '' AS CD");
            strSql.AppendLine("     , '' AS NM");
            strSql.AppendLine(" UNION ALL");

            strSql.AppendLine(" SELECT L1 AS CD                  ");
            strSql.AppendLine(" , '' AS NM                       ");
            strSql.AppendLine(" FROM LENGTHS                     ");
            strSql.AppendLine(" WHERE L1 IS NOT NULL AND L1 != ''");
            if(!sModel.Equals("")) strSql.AppendLine(" AND Model = '" + sModel +"'");
            strSql.AppendLine(" GROUP BY L1                      ");

            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");


            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            GetGridRetr();
        }

        private void GetGridRetr()
        {
            string sLength = LkupLength.Text;
            string sGrade = CboGrade.Text;
            StringBuilder strSql = new StringBuilder();
                             strSql.AppendLine("");
                             strSql.AppendLine("SELECT J_ID ");
                             strSql.AppendLine("     , Model "); 
                             strSql.AppendLine("     , Title");
                             strSql.AppendLine("     , L1 ");
                             strSql.AppendLine("     , J_Unit");
                             strSql.AppendLine("     , P_Unit");
                             strSql.AppendLine(" FROM LENGTHS");
                             strSql.AppendLine(" WHERE 1=1");
                             //strSql.AppendLine(" WHERE 1");
            if (sLength !="") strSql.AppendLine("AND L1 ='"+sLength+"'");
            if (sGrade != "") strSql.AppendLine("AND Model='" + sGrade + "'");
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon,strSql.ToString());
            GridRetr.DataSource = dt;
          
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;

            }

            if(GridViewRetr.RowCount > 0)
            {
                string sModel = GridViewRetr.GetRowCellValue(GridViewRetr.RowCount - 1, GridColModel)?.ToString();

                if (string.IsNullOrEmpty(sModel))
                {
                    XtraMessageBox.Show("품목을 입력해주세요.");
                    return;
                }
            }
            

            GridViewRetr.AddNewRow();
            GridViewRetr.Focus();
            GridViewRetr.FocusedColumn = GridColModel;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            SaveGrid();
        }

        private void SaveGrid()
        {
            Cursor = Cursors.WaitCursor;
            DataTable dt = (DataTable)GridRetr.DataSource;
            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMerge = dsSave.Tables[0];

            for (int i = dtMerge.Rows.Count - 1; i >= 0; i--)
            {
                string sModel = dtMerge.Rows[i]["Model"]?.ToString();

                if (string.IsNullOrEmpty(sModel))
                {
                    dtMerge.Rows[i].Delete();
                }
            }

            string sJid = string.Empty;
            string sItemCls = string.Empty;
            string sGrade = string.Empty;
            string sSize = string.Empty;
            string sPunit = string.Empty;
            string sJunit = string.Empty;

            if (dtMerge.Rows.Count > 0)
            {
                try
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    StringBuilder strSql = new StringBuilder();

                    for (int j = 0; j < dtMerge.Rows.Count; j++)
                    {
                        sJid = Convert.ToString(dtMerge.Rows[j]["J_ID"]);
                        sItemCls = Convert.ToString(dtMerge.Rows[j]["Model"]);
                        sGrade = Convert.ToString(dtMerge.Rows[j]["Title"]);
                        sSize = Convert.ToString(dtMerge.Rows[j]["L1"]);
                        sPunit = Convert.ToString(dtMerge.Rows[j]["P_Unit"]);
                        sJunit = Convert.ToString(dtMerge.Rows[j]["J_Unit"]);
                        strSql.Clear();

                        strSql.AppendLine(" MERGE INTO LENGTHS AS a             ");
                        strSql.AppendLine("     USING(SELECT                    ");
                        strSql.AppendLine("         J_ID = '" + sJid + "'                   ");
                        strSql.AppendLine("         , Model = '" + sItemCls + "'          ");
                        strSql.AppendLine("            , Title = '" + sGrade + "'       ");
                        strSql.AppendLine("            , L1 = '" + sSize + "'            ");
                        strSql.AppendLine("            , P_Unit = '" + sPunit + "'        ");
                        strSql.AppendLine("            , J_Unit = '" + sJunit + "') AS b  ");
                        strSql.AppendLine("     ON a.J_ID = b.J_ID              ");
                        strSql.AppendLine("     WHEN MATCHED THEN UPDATE SET    ");
                        strSql.AppendLine("         Model = '" + sItemCls + "'            ");
                        strSql.AppendLine("            , Title = '" + sGrade + "'       ");
                        strSql.AppendLine("            , L1 = '" + sSize + "'            ");
                        strSql.AppendLine("            , P_Unit = '" + sPunit + "'        ");
                        strSql.AppendLine("            , J_Unit = '" + sJunit + "'        ");
                        strSql.AppendLine("     WHEN NOT MATCHED THEN INSERT(   ");
                        strSql.AppendLine("         Model                       ");
                        strSql.AppendLine("                , Title              ");
                        strSql.AppendLine("                , L1                 ");
                        strSql.AppendLine("                , P_Unit             ");
                        strSql.AppendLine("                , J_Unit)            ");
                        strSql.AppendLine("     VALUES(                         ");
                        strSql.AppendLine("         '" + sItemCls + "'                    ");
                        strSql.AppendLine("                , '" + sGrade + "'           ");
                        strSql.AppendLine("                , '" + sSize + "'             ");
                        strSql.AppendLine("                , '" + sPunit + "'             ");
                        strSql.AppendLine("                , '" + sJunit + "');           ");

                        /*
                        strSql.AppendLine("INSERT INTO LENGTHS");
                        strSql.AppendLine("           (");
                        strSql.AppendLine("            J_ID");
                        strSql.AppendLine("           ,Model");
                        strSql.AppendLine("           ,Title");
                        strSql.AppendLine("           ,L1");
                        strSql.AppendLine("           ,P_Unit");
                        strSql.AppendLine("           ,J_Unit");
                        strSql.AppendLine("           )");
                        strSql.AppendLine("    VALUES (");
            if(sJid !="") strSql.AppendLine("           '" + sJid + "'");
                    else strSql.AppendLine("            NULL");
                        strSql.AppendLine("           ,'" + sItemCls + "'");
                        strSql.AppendLine("           ,'" + sGrade + "'");
                        strSql.AppendLine("           ,'" + sSize + "'");
                        strSql.AppendLine("           ,'" + sPunit + "'");
                        strSql.AppendLine("           ,'" + sJunit + "'");
                        strSql.AppendLine("           )");
                        strSql.AppendLine("ON DUPLICATE KEY UPDATE");
                        strSql.AppendLine("        Model   = '" + sItemCls + "'");
                        strSql.AppendLine("       ,Title   ='" + sGrade + "'");
                        strSql.AppendLine("       ,L1      ='" + sSize + "'");
                        strSql.AppendLine("       ,P_Unit  ='" + sPunit + "'");
                        strSql.AppendLine("       ,J_Unit  ='" + sJunit + "'");
                        */

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        string sLogRmk = "Table:LENGTHS -> J_ID:" + sJid + "Model:" + sItemCls;
                        ClsFunc.LogInsert(DateTime.Now.ToString(), FmMainToolBar2.UserID, (j+1).ToString(), "S", this.Name, sLogRmk, cmd);
                    }

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    MessageBox.Show("저장을 완료했습니다.");

                    GetGridRetr();
                    GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColTitle, dtMerge.Rows[0]["Title"]?.ToString());
               
                }
                catch (Exception ex)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                XtraMessageBox.Show("변경된 내용이 없습니다.");
            }
            Cursor = Cursors.Default;

        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void StandardCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {

            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
            }
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void LkupLength_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void CboGrade_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sModel = CboGrade.EditValue.ToString();

            DataTable dt = GetLookUpData(sModel);
            ComGrid.SetLookUpEdit(LkupLength, dt, "CD", "CD", "Y");
            LkupLength.Properties.PopulateColumns();
            LkupLength.Properties.Columns[1].Visible = false;
        }

        private void StandardCd_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void GridViewRetr_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridViewRetr.UpdateCurrentRow();
        }
    }
}