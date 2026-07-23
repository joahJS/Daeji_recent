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
    public partial class ResetGridViewLayout : DevExpress.XtraEditors.XtraForm
    {
        public ResetGridViewLayout()
        {
            InitializeComponent();
        }

        private void ResetGridViewLayout_Load(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();
            GetAuthorOfLayoutInfo();
        }
        
        private void GetAuthorOfLayoutInfo()
        {
            string sUsrCd = FmMainToolBar2.drUser["USRCD"]?.ToString();

            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.USRCD ");
            strSql.AppendLine("      , C.COM_NM ");
            strSql.AppendLine("      , B.PGMNM ");
            strSql.AppendLine("      , B.PGMID ");
            strSql.AppendLine("   FROM ZPGMAUT A ");
            strSql.AppendLine("   LEFT OUTER JOIN ZPGMLST B ");
            strSql.AppendLine("     ON B.PGMID = A.PGMID ");
            strSql.AppendLine("   LEFT OUTER JOIN COM_BASE_CD C  ");
            strSql.AppendLine("     ON C.COM_CD = B.PGGRP ");
            strSql.AppendLine("    AND C.CD_GB = 'Z1' ");
            strSql.AppendLine("  WHERE A.USRCD = " + sUsrCd + "  ");
            strSql.AppendLine("    AND A.USE_Y = 'Y' ");
            strSql.AppendLine("    AND B.USEYN = 'Y' ");
            strSql.AppendLine("  ORDER BY B.PGGRP ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;

            DataTable dtCopy = dt.Clone();
            GridInitial.DataSource = dtCopy;

            Cursor = Cursors.Default;
        }

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                Cursor = Cursors.WaitCursor;

                DataRow row = GridViewRetr.GetFocusedDataRow();

                DataTable dt = (DataTable)GridInitial.DataSource;

                if (dt.Rows.Count == 0)
                {
                    dt.ImportRow(row);
                    GridInitial.DataSource = dt;
                }
                else
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i]["PGMID"] == row["PGMID"])
                        {
                            Cursor = Cursors.Default;
                            return;
                        }
                    }
                    dt.ImportRow(row);
                    GridInitial.DataSource = dt;
                }
                Cursor = Cursors.Default;
            }
        }

        private void GridViewInitial_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                Cursor = Cursors.WaitCursor;

                DataRow row = GridViewInitial.GetFocusedDataRow();

                DataTable dt = (DataTable)GridInitial.DataSource;

                for(int i = 0; i < dt.Rows.Count; i++)
                {
                    if(dt.Rows[i]["PGMID"] == row["PGMID"])
                    {
                        dt.Rows.RemoveAt(i);

                        Cursor = Cursors.Default;
                        break;
                    }
                }

                GridInitial.DataSource = dt;

                Cursor = Cursors.Default;
            }
        }

        private void BtnInit_Click(object sender, EventArgs e)
        {
            if (XtraMessageBox.Show("선택한 프로그램의 레이아웃 설정값이 사라지게 됩니다. \r\n그래도 초기화 하시겠습니까?"
               , "레이아웃 초기화 여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            DataTable dt = (DataTable)GridInitial.DataSource;
            string sId = FmMainToolBar2.UserID;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string sPgmId = dt.Rows[i]["PGMID"].ToString();

                if (System.IO.Directory.Exists(Application.StartupPath + @"\xaml\" + sId))
                {
                    string[] files = System.IO.Directory.GetFiles(Application.StartupPath + @"\xaml\" + sId);

                    foreach (string s in files)
                    {
                        Cursor = Cursors.WaitCursor;
                        string fileName = System.IO.Path.GetFileName(s);
                        if (fileName.Contains(sPgmId))
                        {
                            string deletefile = Application.StartupPath + @"\xaml\" + sId + @"\" + fileName;
                            System.IO.File.Delete(deletefile);

                            Cursor = Cursors.Default;
                        }
                        Cursor = Cursors.Default;
                    }
                }
                else
                {
                    XtraMessageBox.Show("경로에 해당 폴더가 존재하지 않습니다.");
                    break;
                }
                
            }
            Cursor = Cursors.Default;

            if (dt != null && dt.Rows.Count > 0)
            {
                XtraMessageBox.Show("화면초기화가 완료되었습니다.");
                GetAuthorOfLayoutInfo();
            }
            else
            {
                XtraMessageBox.Show("좌측에서 초기화 할 항목을 선택해주세요.");
            }
        }

        private void BtnAllInit_Click(object sender, EventArgs e)
        {
            if (XtraMessageBox.Show("모든 프로그램의 레이아웃 설정값이 사라지게 됩니다. \r\n그래도 초기화 하시겠습니까?"
               , "레이아웃 전체 초기화 여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            string sId = FmMainToolBar2.UserID;
            if (System.IO.Directory.Exists(Application.StartupPath + @"\xaml\" + sId))
            {
                string[] files = System.IO.Directory.GetFiles(Application.StartupPath + @"\xaml\" + sId);

                foreach (string s in files)
                {
                    Cursor = Cursors.WaitCursor;

                    string fileName = System.IO.Path.GetFileName(s);
                    string deletefile = Application.StartupPath + @"\xaml\" + sId + @"\" + fileName;

                    System.IO.File.Delete(deletefile);

                    Cursor = Cursors.Default;
                }
            }

            XtraMessageBox.Show("전체 화면 초기화가 완료되었습니다.");
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void ResetGridViewLayout_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
            {
            }
            else if(e.KeyCode == Keys.F3)
            {
                BtnInit_Click(null, null);
            }
            else if(e.KeyCode == Keys.F5)
            {
                BtnAllInit_Click(null, null);
            }
        }
    }
}