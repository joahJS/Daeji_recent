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
using DevExpress.DataAccess.Excel;
using System.Collections;
using System.Data.SqlClient;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraGrid.Views.Grid;
using ComLib;
using System.IO;

namespace AccAdm
{
    public partial class IN001F02 : DevExpress.XtraEditors.XtraForm
    {
        public IN001F02()
        {
            InitializeComponent();
        }
        public string _BASYY;
        private string PROCEDURE_ID = "DP_IN001F02";
        
        private void IN001F02_Load(object sender, EventArgs e)
        {
            SetLoadFormLayout();

            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            ComnEtcFunc.gp_SetColorFocused(layoutControl2);
            Dtyyyy.EditValue = DateTime.Today;

            BtnRetr.PerformClick();
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { bandedGridView1 };
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

        private void IN001F02_Shown(object sender, EventArgs e)
        {
            Dtyyyy.Focus();
        }

        public void BtnRetr_Click(object sender, EventArgs e)
        {
            
            string sIYYYY = Dtyyyy.EditValue?.ToString().Substring(0,4);
            string sFIND_IDX = Cbserch.SelectedIndex.ToString();
            string sFIND_WORD = Txserch.EditValue?.ToString().Trim();
            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            
            try
            {
                dicParams.Clear();
                dicParams.Add("CMD", "LIST1");
                dicParams.Add("IYYYY", sIYYYY);
                dicParams.Add("FIND_IDX", sFIND_IDX);
                dicParams.Add("FIND_WORD", sFIND_WORD);
            
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
            
                if (dt != null)
                {
                    GridRetr.DataSource = dt;
                    if (dt.Rows.Count > 0)
                    {
                        GridRetr.Focus();
                    }
                    else
                    {
                        Dtyyyy.Focus();
                    }
            
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString(), "조회 오류");
            }
        }
        
        string _sFileName;
        private void BtnUpload_Click(object sender, EventArgs e)
        {
            IN001F02_POP02 frm = new IN001F02_POP02();

            frm.Owner = this;

            if (frm.ShowDialog() == DialogResult.OK)
            {
                XtraOpenFileDialog fileDlg = new XtraOpenFileDialog();
                fileDlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.Filter = "Excel files (*.xls,*xlsx)|*.xls;*xlsx|All files (*.*)|*.*";
                fileDlg.FilterIndex = 1;
                fileDlg.Title = "업로드 파일 선택";
                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    _sFileName = fileDlg.FileName;

                    try
                    {
                        SplashScreenManager.ShowForm(typeof(WaitForm1));

                        if (!string.IsNullOrEmpty(_sFileName))
                        {
                            DevExpress.DataAccess.Excel.ExcelDataSource excelDataSource = new DevExpress.DataAccess.Excel.ExcelDataSource
                            {
                                FileName = _sFileName
                            };

                            ExcelWorksheetSettings workSheetSettings = new ExcelWorksheetSettings("Sheet1");
                            excelDataSource.SourceOptions = new ExcelSourceOptions(workSheetSettings)
                            {
                                SkipEmptyRows = true,
                                UseFirstRowAsHeader = true
                            };

                            excelDataSource.Fill();

                            DataTable dt = ExcelDataSourceExtension.ToDataTable(excelDataSource);

                            if (dt != null)
                            {
                                string sUSER = FmMainToolBar2.drUser["USRCD"]?.ToString();

                                StringBuilder strSql = new StringBuilder();

                                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                                cmd.Transaction = DBConn.dbTran;

                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    //string sIYYYY = DateTime.Today.ToString().Substring(0, 4);
                                    string sIYYYY = _BASYY;
                                    string sIFROM = string.Empty;
                                    string sITO = string.Empty;
                                    string sINWN01 = string.Empty;
                                    string sINWN02 = string.Empty;
                                    string sINWN03 = string.Empty;
                                    string sINWN04 = string.Empty;
                                    string sINWN05 = string.Empty;
                                    string sINWN06 = string.Empty;
                                    string sINWN07 = string.Empty;
                                    string sINWN08 = string.Empty;
                                    string sINWN09 = string.Empty;
                                    string sINWN10 = string.Empty;
                                    string sINWN11 = string.Empty;
                                    string sRK = string.Empty;

                                    int iIFROM = 0;
                                    int iITO = 0;
                                    int iINWN01 = 0;
                                    int iINWN02 = 0;
                                    int iINWN03 = 0;
                                    int iINWN04 = 0;
                                    int iINWN05 = 0;
                                    int iINWN06 = 0;
                                    int iINWN07 = 0;
                                    int iINWN08 = 0;
                                    int iINWN09 = 0;
                                    int iINWN10 = 0;
                                    int iINWN11 = 0;

                                    sIFROM = dt.Rows[i]["이상"]?.ToString();
                                    sITO = dt.Rows[i]["미만"]?.ToString();

                                    sINWN01 = dt.Rows[i]["1"]?.ToString();
                                    sINWN02 = dt.Rows[i]["2"]?.ToString();
                                    sINWN03 = dt.Rows[i]["3"]?.ToString();
                                    sINWN04 = dt.Rows[i]["4"]?.ToString();
                                    sINWN05 = dt.Rows[i]["5"]?.ToString();
                                    sINWN06 = dt.Rows[i]["6"]?.ToString();
                                    sINWN07 = dt.Rows[i]["7"]?.ToString();
                                    sINWN08 = dt.Rows[i]["8"]?.ToString();
                                    sINWN09 = dt.Rows[i]["9"]?.ToString();
                                    sINWN10 = dt.Rows[i]["10"]?.ToString();
                                    sINWN11 = dt.Rows[i]["11"]?.ToString();

                                    int.TryParse(sIFROM.Replace(",", ""), out iIFROM);
                                    int.TryParse(sITO.Replace(",", ""), out iITO);

                                    int.TryParse(sINWN01.Replace(",",""), out iINWN01);
                                    int.TryParse(sINWN02.Replace(",",""), out iINWN02);
                                    int.TryParse(sINWN03.Replace(",",""), out iINWN03);
                                    int.TryParse(sINWN04.Replace(",",""), out iINWN04);
                                    int.TryParse(sINWN05.Replace(",",""), out iINWN05);
                                    int.TryParse(sINWN06.Replace(",",""), out iINWN06);
                                    int.TryParse(sINWN07.Replace(",",""), out iINWN07);
                                    int.TryParse(sINWN08.Replace(",",""), out iINWN08);
                                    int.TryParse(sINWN09.Replace(",",""), out iINWN09);
                                    int.TryParse(sINWN10.Replace(",",""), out iINWN10);
                                    int.TryParse(sINWN11.Replace(",",""), out iINWN11);

                                    if (!string.IsNullOrEmpty(sIYYYY) && iIFROM != 0 && iITO != 0)
                                    {
                                        strSql.Clear();
                                        strSql.AppendLine(" IF EXISTS(SELECT IYYYY FROM INTAX WHERE IYYYY = '" + sIYYYY + "' AND IFROM = " + iIFROM + ") ");
                                        strSql.AppendLine("    BEGIN                                                           ");
                                        strSql.AppendLine("          UPDATE INTAX                                              ");
                                        strSql.AppendLine("             SET IFROM = " + iIFROM + "                                  ");
                                        strSql.AppendLine("               , ITO = " + iITO + "                                  ");
                                        strSql.AppendLine("               , INWN01 = " + iINWN01 + "                                  ");
                                        strSql.AppendLine("               , INWN02 = " + iINWN02 + "                                  ");
                                        strSql.AppendLine("               , INWN03 = " + iINWN03 + "                                  ");
                                        strSql.AppendLine("               , INWN04 = " + iINWN04 + "                                  ");
                                        strSql.AppendLine("               , INWN05 = " + iINWN05 + "                                  ");
                                        strSql.AppendLine("               , INWN06 = " + iINWN06 + "                                  ");
                                        strSql.AppendLine("               , INWN07 = " + iINWN07 + "                                  ");
                                        strSql.AppendLine("               , INWN08 = " + iINWN08 + "                                  ");
                                        strSql.AppendLine("               , INWN09 = " + iINWN09 + "                                  ");
                                        strSql.AppendLine("               , INWN10 = " + iINWN10 + "                                  ");
                                        strSql.AppendLine("               , INWN11 = " + iINWN11 + "                                  ");
                                        strSql.AppendLine("               , CUSER = '" + sUSER + "'                                  ");
                                        strSql.AppendLine("               , CDATE = CONVERT([varchar](20),getdate(),(20))    ");
                                        strSql.AppendLine("           WHERE IYYYY = '" + sIYYYY + "' AND IFROM =" + iIFROM + "       ");
                                        strSql.AppendLine("      END                                                           ");
                                        strSql.AppendLine(" ELSE                                                               ");
                                        strSql.AppendLine("    BEGIN                                                           ");
                                        strSql.AppendLine("          INSERT INTO INTAX( IYYYY                                  ");
                                        strSql.AppendLine("                           , IFROM                                  ");
                                        strSql.AppendLine("                           , ITO                                  ");
                                        strSql.AppendLine("                           , INWN01                                  ");
                                        strSql.AppendLine("                           , INWN02                                  ");
                                        strSql.AppendLine("                           , INWN03                                  ");
                                        strSql.AppendLine("                           , INWN04                                  ");
                                        strSql.AppendLine("                           , INWN05                                  ");
                                        strSql.AppendLine("                           , INWN06                                  ");
                                        strSql.AppendLine("                           , INWN07                                  ");
                                        strSql.AppendLine("                           , INWN08                                  ");
                                        strSql.AppendLine("                           , INWN09                                  ");
                                        strSql.AppendLine("                           , INWN10                                  ");
                                        strSql.AppendLine("                           , INWN11                                  ");
                                        strSql.AppendLine("                           , CUSER )                                ");
                                        strSql.AppendLine("                     VALUES( '" + sIYYYY + "'                                  ");
                                        strSql.AppendLine("                           , " + iIFROM + "                                  ");
                                        strSql.AppendLine("                           , " + iITO + "                                  ");
                                        strSql.AppendLine("                           , " + iINWN01 + "                                  ");
                                        strSql.AppendLine("                           , " + iINWN02 + "                                  ");
                                        strSql.AppendLine("                           , " + iINWN03 + "                                  ");
                                        strSql.AppendLine("                           , " + iINWN04 + "                                  ");
                                        strSql.AppendLine("                           , " + iINWN05 + "                                  ");
                                        strSql.AppendLine("                           , " + iINWN06 + "                                  ");
                                        strSql.AppendLine("                           , " + iINWN07 + "                                  ");
                                        strSql.AppendLine("                           , " + iINWN08 + "                                  ");
                                        strSql.AppendLine("                           , " + iINWN09 + "                                  ");
                                        strSql.AppendLine("                           , " + iINWN10 + "                                  ");
                                        strSql.AppendLine("                           , " + iINWN11 + "                                  ");
                                        strSql.AppendLine("                           , '" + sUSER + "' )                                ");
                                        strSql.AppendLine("      END                                                           ");

                                        cmd.CommandType = CommandType.Text;
                                        cmd.CommandText = strSql.ToString();
                                        cmd.ExecuteNonQuery();
                                    }
                                }

                                XtraMessageBox.Show("업로드를 완료했습니다.");
                                DBConn.dbTran.Commit();
                                DBConn.dbTran = null;

                                BtnRetr.PerformClick();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Equals("Error Opening file"))
                        {
                            //파일이 열려있음 체크
                            XtraMessageBox.Show(((ex.InnerException).InnerException).Message);
                        }
                        else
                        {
                            XtraMessageBox.Show(ex.Message);
                        }
                        DBConn.dbTran.Rollback();
                        DBConn.dbTran = null;
                    }
                    finally
                    {
                        _sFileName = string.Empty;
                        SplashScreenManager.CloseForm();
                    }
                }
                fileDlg.Dispose();
            }
        }



        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void IN001F02_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F1)
                BtnUpload.PerformClick();
            else if (e.KeyCode == Keys.F8)
                BtnExcel.PerformClick();
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnRetr1.PerformClick();
        }

        private void TxtNmk_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnRetr.Focus();
            }
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            string sPath = string.Format(@"{0}\Temp_File\", Application.StartupPath);
            string sFileName = "간이세액조견표 상세내역.xls";

            ComnEtcFunc.ExportExcelFile(sPath, sFileName, GridRetr);
        }

        private void BtnRetr1_Click(object sender, EventArgs e)
        {
            IN001F02_POP01 frm = new IN001F02_POP01();

            frm.Owner = this;
            frm.Show();
        }

        private void IN001F02_TextChanged(object sender, EventArgs e)
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

        private void Txserch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = Dtyyyy.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevYear(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                Dtyyyy.EditValue = sPrevDate;

                BtnRetr.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = Dtyyyy.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextYear(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                Dtyyyy.EditValue = sNextDate;

                BtnRetr.PerformClick();
            }
        }
    }
}

public static class ExcelDataSourceExtension
{
    public static DataTable ToDataTable(this ExcelDataSource excelDataSource)
    {
        IList list = ((IListSource)excelDataSource).GetList();
        DevExpress.DataAccess.Native.Excel.DataView dataView = (DevExpress.DataAccess.Native.Excel.DataView)list;
        List<PropertyDescriptor> props = dataView.Columns.ToList<PropertyDescriptor>();
        DataTable table = new DataTable();
        for (int i = 0; i < props.Count; i++)
        {
            PropertyDescriptor prop = props[i];
            table.Columns.Add(prop.Name, prop.PropertyType);
        }
        object[] values = new object[props.Count];
        foreach (DevExpress.DataAccess.Native.Excel.ViewRow item in list)
        {
            for (int i = 0; i < values.Length; i++)
            {
                values[i] = props[i].GetValue(item);
            }
            table.Rows.Add(values);
        }
        return table;
    }
}