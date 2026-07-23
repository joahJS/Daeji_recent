using DevExpress.Data;
using DevExpress.Utils;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraGrid.Helpers;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.Drawing;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AccAdm
{
    class ComnGridFunc
    {
        #region[마지막 Row Focusing 여부]
        public static bool VerificateCheckLastRowFocusing(GridView view)
        {
            if (view.FocusedRowHandle != (view.RowCount - 1))
                return true;
            else
                return false;
        }
        #endregion[마지막 Row Focusing 여부]

        public static void SetInitGridRowColor(GridView view)
        {
            //view.Appearance.Empty.BackColor = Color.White;
            //view.Appearance.Preview.BackColor = Color.White;
            //view.Appearance.Preview.ForeColor = Color.Black;
            //view.Appearance.Row.BackColor = Color.White;
            //view.Appearance.Row.ForeColor = Color.Black;
            //view.Appearance.FocusedRow.BackColor = SystemColors.ActiveCaption;
            //view.Appearance.FocusedRow.ForeColor = Color.Black;
            //view.Appearance.FocusedCell.BackColor = SystemColors.ActiveCaption;
            //view.Appearance.FocusedCell.ForeColor = Color.Black;
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
            if(DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName != "DevExpress Dark Style")
            {
                if (e.RowHandle % 2 == 0)
                {
                    e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
                }
            }
            else if(DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName == "DevExpress Dark Style")
            {
                if (e.RowHandle % 2 == 0)
                {
                    e.Appearance.BackColor = SystemColors.ControlDark;
                }
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

        #region cell외곽선 굵게 칠하기
        public static class CellDrawHelper
        {
            public static void DrawCellBorder(DevExpress.XtraGrid.Views.Base.RowCellCustomDrawEventArgs e)
            {
                Brush brush = Brushes.Black;
                e.Cache.FillRectangle(brush, new Rectangle(e.Bounds.Right - 1, e.Bounds.Y - 1, 2, e.Bounds.Height + 2));  // 오른쪽
                                                                                                                          //e.Cache.FillRectangle(brush, new Rectangle(e.Bounds.X - 1, e.Bounds.Y - 1, e.Bounds.Width + 2, 2));
                                                                                                                          //e.Cache.FillRectangle(brush, new Rectangle(e.Bounds.X - 1, e.Bounds.Bottom - 1, e.Bounds.Width + 2, 2));
            }
            public static void DoDefaultDrawCell(GridView view, RowCellCustomDrawEventArgs e)
            {
                e.Appearance.FillRectangle(e.Cache, e.Bounds);
                ((IViewController)view.GridControl).EditorHelper.DrawCellEdit(new GridViewDrawArgs(e.Cache, (view.GetViewInfo() as GridViewInfo), e.Bounds), (e.Cell as GridCellInfo).Editor, (e.Cell as GridCellInfo).ViewInfo, e.Appearance, (e.Cell as GridCellInfo).CellValueRect.Location);
            }
        }
        #endregion

        #region 그리드 밴드 추가관련
        #region 그리드 컨트롤 구하기 - GetGridControl(name)

        /// <summary>
        /// 그리드 컨트롤 구하기
        /// </summary>
        /// <param name="name">명칭</param>
        /// <returns>그리드 컨트롤</returns>
        public static GridControl GetGridControl(string name)
        {
            GridControl gridControl = new GridControl();

            gridControl.Name = name;

            return gridControl;
        }

        #endregion
        #region 초기화 시작하기 - BeginInitialize(gridControl)

        /// <summary>
        /// 초기화 시작하기
        /// </summary>
        /// <param name="gridControl">그리드 컨트롤</param>
        public static void BeginInitialize(GridControl gridControl)
        {
            ((ISupportInitialize)gridControl).BeginInit();
        }

        #endregion
        #region 초기화 종료하기 - EndInitialize(gridControl)

        /// <summary>
        /// 초기화 종료하기
        /// </summary>
        /// <param name="gridControl">그리드 컨트롤</param>
        public static void EndInitialize(GridControl gridControl)
        {
            ((ISupportInitialize)gridControl).EndInit();
        }

        #endregion
        #region 뷰 추가하기 - AddView(gridControl, baseView, isMainView)

        /// <summary>
        /// 뷰 추가하기
        /// </summary>
        /// <param name="gridControl">그리드 컨트롤</param>
        /// <param name="baseView">기본 뷰</param>
        /// <param name="isMainView">메인 뷰 여부</param>
        public static void AddView(GridControl gridControl, BaseView baseView, bool isMainView)
        {
            gridControl.ViewCollection.Add(baseView);

            if (isMainView)
            {
                gridControl.MainView = baseView;
            }

            baseView.GridControl = gridControl;
        }

        #endregion

        #region 개선 밴드 그리드 뷰 구하기 - GetAdvancedBandedGridView(name, editable, multiSelect, showGroupPanel, showIndicator)

        /// <summary>
        /// 개선 밴드 그리드 뷰 구하기
        /// </summary>
        /// <param name="name">명칭</param>
        /// <param name="editable">편집 가능 여부</param>
        /// <param name="multiSelect">복수 선택 여부</param>
        /// <param name="showGroupPanel">그룹 패널 표시 여부</param>
        /// <param name="showIndicator">지시자 표시 여부</param>
        /// <returns>개선 밴드 그리드 뷰</returns>
        public static BandedGridView GetBandedGridView(string name, bool editable, bool multiSelect, bool showGroupPanel,
            bool showIndicator)
        {
            BandedGridView BandedGridView = new BandedGridView();

            BandedGridView.Name = name;
            BandedGridView.OptionsBehavior.Editable = editable;
            BandedGridView.OptionsBehavior.EditorShowMode = EditorShowMode.MouseDown;
            BandedGridView.OptionsCustomization.AllowColumnMoving = false;
            BandedGridView.OptionsCustomization.AllowColumnResizing = false;
            BandedGridView.OptionsCustomization.AllowFilter = false;
            BandedGridView.OptionsCustomization.AllowSort = false;
            BandedGridView.OptionsMenu.EnableColumnMenu = false;
            BandedGridView.OptionsSelection.MultiSelect = multiSelect;
            BandedGridView.OptionsView.ColumnAutoWidth = false;
            BandedGridView.OptionsView.ShowGroupPanel = showGroupPanel;
            BandedGridView.OptionsView.ShowIndicator = showIndicator;

            return BandedGridView;
        }

        #endregion
        #region 초기화 시작하기 - BeginInitialize(bandedGridView)

        /// <summary>
        /// 초기화 시작하기
        /// </summary>
        /// <param name="bandedGridView">개선 밴드 그리드 뷰</param>
        public static void BeginInitialize(BandedGridView bandedGridView)
        {
            ((ISupportInitialize)bandedGridView).BeginInit();
        }

        #endregion
        #region 초기화 종료하기 - EndInitialize(bandedGridView)

        /// <summary>
        /// 초기화 종료하기
        /// </summary>
        /// <param name="bandedGridView">개선 밴드 그리드 뷰</param>
        public static void EndInitialize(BandedGridView bandedGridView)
        {
            ((ISupportInitialize)bandedGridView).EndInit();
        }

        #endregion
        #region 밴드 추가하기 - AddBand(bandedGridView, gridBand)

        /// <summary>
        /// 밴드 추가하기
        /// </summary>
        /// <param name="bandedGridView">개선 밴드 그리드 뷰</param>
        /// <param name="gridBand">그리드 밴드</param>
        public static void AddBand(BandedGridView bandedGridView, GridBand gridBand)
        {
            bandedGridView.Bands.Add(gridBand);
        }

        #endregion

        #region 그리드 밴드 구하기 - GetGridBand(name, width, caption, horzAlignment, vertAlignment)

        /// <summary>
        /// 그리드 밴드 구하기
        /// </summary>
        /// <param name="name">명칭</param>
        /// <param name="width">너비</param>
        /// <param name="caption">제목</param>
        /// <param name="horzAlignment">수평 정렬</param>
        /// <param name="vertAlignment">수직 정렬</param>
        /// <returns>그리드 밴드</returns>
        public static GridBand GetGridBand(string name, int width, string caption, HorzAlignment horzAlignment, VertAlignment vertAlignment)
        {
            GridBand gridBand = new GridBand();

            gridBand.Name = name;
            gridBand.Width = width;
            gridBand.Caption = caption;
            gridBand.AppearanceHeader.Options.UseTextOptions = true;
            gridBand.AppearanceHeader.TextOptions.HAlignment = horzAlignment;
            gridBand.AppearanceHeader.TextOptions.VAlignment = vertAlignment;

            return gridBand;
        }

        #endregion
        #region 그리드 밴드 구하기 - GetGridBand(name, width, caption)

        /// <summary>
        /// 그리드 밴드 구하기
        /// </summary>
        /// <param name="name">명칭</param>
        /// <param name="width">너비</param>
        /// <param name="caption">제목</param>
        /// <returns>그리드 밴드</returns>
        public static GridBand GetGridBand(string name, int width, string caption)
        {
            return GetGridBand(name, width, caption, HorzAlignment.Center, VertAlignment.Center);
        }

        #endregion
        #region 자식 그리드 밴드 추가하기 - AddChildGridBand(parentGridBand, childGridBand)

        /// <summary>
        /// 자식 그리드 밴드 추가하기
        /// </summary>
        /// <param name="parentGridBand">부모 그리드 밴드</param>
        /// <param name="childGridBand">자식 그리드 밴드</param>
        public static void AddChildGridBand(GridBand parentGridBand, GridBand childGridBand)
        {
            parentGridBand.Children.Add(childGridBand);
        }

        #endregion
        #region 밴드 그리드 컬럼 구하기 - GetBandedGridColumn(caption, fieldName, unboundColumnType, width, horzAlignment, allowEdit, formatType, formatString)

        /// <summary>
        /// 밴드 그리드 컬럼 구하기
        /// </summary>
        /// <param name="caption">제목</param>
        /// <param name="fieldName">필드명</param>
        /// <param name="unboundColumnType">언바운드 컬럼 종류</param>
        /// <param name="width">너비</param>
        /// <param name="horzAlignment">수평 정렬</param>
        /// <param name="allowEdit">편집 허용 여부</param>
        /// <param name="formatType">형식 종류</param>
        /// <param name="formatString">형식 문자열</param>
        /// <returns>밴드 그리드 컬럼</returns>
        public static BandedGridColumn GetBandedGridColumn(string caption, string fieldName, UnboundColumnType unboundColumnType, int width,
            HorzAlignment horzAlignment, bool allowEdit, FormatType formatType, string formatString)
        {
            BandedGridColumn bandedGridColumn = new BandedGridColumn();

            bandedGridColumn.Name = string.Empty;
            bandedGridColumn.Caption = caption;
            bandedGridColumn.FieldName = fieldName;
            bandedGridColumn.UnboundType = unboundColumnType;
            bandedGridColumn.Width = width;
            bandedGridColumn.AutoFillDown = true;
            bandedGridColumn.AppearanceHeader.Options.UseTextOptions = true;
            bandedGridColumn.AppearanceHeader.TextOptions.HAlignment = HorzAlignment.Center;
            bandedGridColumn.AppearanceCell.Options.UseTextOptions = true;
            bandedGridColumn.AppearanceCell.TextOptions.HAlignment = horzAlignment;
            bandedGridColumn.OptionsColumn.AllowEdit = allowEdit;
            bandedGridColumn.DisplayFormat.FormatType = formatType;
            bandedGridColumn.DisplayFormat.FormatString = formatString;
            bandedGridColumn.Visible = true;

            return bandedGridColumn;
        }

        #endregion
        #region 그리드 컬럼 추가하기 - AddGridColumn(bandedGridView, gridBand, bandedGridColumn)

        /// <summary>
        /// 그리드 컬럼 추가하기
        /// </summary>
        /// <param name="bandedGridView">개선 밴드 그리드 뷰</param>
        /// <param name="gridBand">그리드 밴드</param>
        /// <param name="bandedGridColumn">밴드 그리드 컬럼</param>
        public static void AddGridColumn(BandedGridView bandedGridView, GridBand gridBand, BandedGridColumn bandedGridColumn)
        {
            bandedGridView.Columns.Add(bandedGridColumn);

            if (gridBand != null)
            {
                gridBand.Columns.Add(bandedGridColumn);
            }
        }
        #endregion
        #endregion

        #region [GridView 새 줄 추가, 삭제]
        /// <summary>
        /// 
        /// </summary>
        /// <param name="gridControl">해당 그리드 컨트롤</param>
        /// <param name="view">해당 그리드 뷰</param>
        /// <param name="focusColumn">개행 후 포커스가 가는 컬럼</param>
        /// <param name="seqFieldName">순번 정보가 있는 필드명</param>
        public static void GridViewAddLine(GridControl gridControl, GridView view, GridColumn focusColumn, string seqFieldName)
        {
            if (view.FocusedRowHandle == view.RowCount - 1)
            {
                DataTable dt = gridControl.DataSource as DataTable;
                dt.Rows.Add();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i][seqFieldName] = i.ToString();
                }
                gridControl.DataSource = dt;
                view.FocusedColumn = focusColumn;
                view.FocusedRowHandle = view.RowCount - 1;
            }
            else { SendKeys.Send("{TAB}"); }
        }

        /// <summary>
        /// GridView에 새 줄을 추가하는 메서드.
        /// 마지막 줄이 아니면 TAB키를 누른 효과를 준다.
        /// </summary>
        /// <param name="gridView1">대상 GridView</param>
        /// <param name="seqColumn">순번 Column</param>
        /// <param name="focusColumn">개행 후 포커스가 가야하는 Column</param>
        public static void GridViewAddLine(GridView gridView1, GridColumn seqColumn, GridColumn focusColumn)
        {
            try
            {
                gridView1.AddNewRow();
                gridView1.UpdateCurrentRow();
                for (int i = 0; i < gridView1.RowCount; i++) { gridView1.SetRowCellValue(i, seqColumn, i + 1); }
                gridView1.SetFocusedRowCellValue(seqColumn, gridView1.FocusedRowHandle + 1);
                gridView1.FocusedColumn = focusColumn;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }


        /// <summary>
        /// GridView에 새 줄을 추가하는 메서드.
        /// 마지막 줄이 아니면 TAB키를 누른 효과를 준다.
        /// </summary>
        /// <param name="gridView1">대상 GridView</param>
        /// <param name="seqColumn">순번 Column</param>
        /// <param name="focusColumn">개행 후 포커스가 가야하는 Column</param>
        public static void GridViewLineAdd_Click(GridView gridView1, GridColumn seqColumn, GridColumn focusColumn)
        {
            try
            {
                if (gridView1.FocusedRowHandle == gridView1.RowCount - 1
                    || gridView1.RowCount == 0)
                {
                    gridView1.AddNewRow();
                    gridView1.UpdateCurrentRow();
                    for (int i = 0; i < gridView1.RowCount; i++) { gridView1.SetRowCellValue(i, seqColumn, i + 1); }
                    gridView1.SetFocusedRowCellValue(seqColumn, gridView1.FocusedRowHandle + 1);
                    gridView1.FocusedColumn = focusColumn;

                }
                else { SendKeys.Send("{TAB}"); }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// GridView에 새 줄을 추가하는 메서드.
        /// 마지막 줄이 아니면 TAB키를 누른 효과를 준다.
        /// 개행 후 주어져야 하는 기본 값을 Dictionary로 받음.
        /// </summary>
        /// <param name="gridView1">대상 GridView</param>
        /// <param name="seqColumn">순번 Column</param>
        /// <param name="focusColumn">개행 후 포커스가 가야하는 Column</param>
        /// <param name="addTextDic">개행 후 기본으로 주어저야 하는 값과 그 컬럼 Dictionary </param>
        public static void GridViewLineAdd_Click(GridView gridView1, GridColumn seqColumn, GridColumn focusColumn, Dictionary<GridColumn, string> addTextDic)
        {
            try
            {
                if (gridView1.FocusedRowHandle == gridView1.RowCount - 1
                    || gridView1.RowCount == 0)
                {
                    gridView1.AddNewRow();
                    gridView1.UpdateCurrentRow();
                    for (int i = 0; i < gridView1.RowCount; i++) { gridView1.SetRowCellValue(i, seqColumn, i + 1); }
                    for (int i = 0; i < addTextDic.Count; i++)
                    {
                        gridView1.SetFocusedRowCellValue(addTextDic.Keys.ElementAt(i), addTextDic.Values.ElementAt(i));
                    }

                    gridView1.FocusedColumn = focusColumn;
                }
                else { SendKeys.Send("{TAB}"); }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }


        /// <summary>
        /// 포커싱된 GridView의 줄을 삭제하는 메서드.
        /// GridControl의 DataSource를 활용한다. (개선가능하다면 개선 필요) 
        /// </summary>
        /// <param name="gridView1">대상 GridView</param>
        /// <param name="gridControl1">대상 GridControl</param>
        /// <param name="focusColumn">Focus되어야 하는 Column</param>
        /// <param name="seqFieldName">순번이 입력되는 FieldName</param>
        public static void GridViewLineDelete_Click(GridView gridView1, GridControl gridControl1, GridColumn focusColumn, string seqFieldName)
        {
            try
            {
                if (gridView1.RowCount < 2)
                    return;

                DataTable dt = gridControl1.DataSource as DataTable;
                dt.Rows.RemoveAt(gridView1.FocusedRowHandle);
                for (int i = 0; i < dt.Rows.Count; i++) { dt.Rows[i][seqFieldName] = i + 1; }
                gridControl1.DataSource = dt;
                gridView1.UpdateCurrentRow();
                gridView1.FocusedColumn = focusColumn;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region [GridView 기본세팅 한번에 하기]


        /// <summary>
        /// 일반적인 (그룹이 없는) 조회형 GridView 세팅
        /// </summary>
        /// <param name="view"></param>
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

            view.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            view.Appearance.SelectedRow.BackColor = Color.LightYellow;
            view.Appearance.SelectedRow.ForeColor = Color.Black;
            view.OptionsNavigation.EnterMoveNextColumn = true;
            view.OptionsView.ColumnAutoWidth = false;
            view.OptionsView.ShowGroupPanel = false;
            view.IndicatorWidth = 50;
            view.ColumnPanelRowHeight = 30;
        }

        /// <summary>
        /// Group 형 GridView에 적용할 Style
        /// </summary>
        /// <param name="view"></param>
        public static void GridStyleBasicSettingForGroup(GridView view)
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



            view.GroupLevelStyle += (sender, e) =>
            {
                int i = view.GroupCount;

                if (i == 1)
                {
                    if (e.Level == 0)
                    {
                        e.LevelAppearance.BackColor = Color.FromArgb(80, 0, 150, 255);
                        e.LevelAppearance.ForeColor = Color.Black;
                    }
                }
                else if (i == 2)
                {
                    if (e.Level == 0)
                    {
                        e.LevelAppearance.BackColor = Color.FromArgb(80, 0, 150, 255);
                        e.LevelAppearance.ForeColor = Color.Black;
                    }
                    else if (e.Level == 1)
                    {
                        e.LevelAppearance.BackColor = Color.FromArgb(50, 255, 150, 0);
                        e.LevelAppearance.ForeColor = Color.Black;
                    }
                }
                else if (i == 3)
                {
                    if (e.Level == 0)
                    {
                        e.LevelAppearance.BackColor = Color.FromArgb(80, 0, 150, 255);
                        e.LevelAppearance.ForeColor = Color.Black;
                    }
                    else if (e.Level == 1)
                    {
                        e.LevelAppearance.BackColor = Color.FromArgb(50, 255, 150, 0);
                        e.LevelAppearance.ForeColor = Color.Black;
                    }
                    else if (e.Level == 2)
                    {
                        e.LevelAppearance.BackColor = Color.LightYellow;
                        e.LevelAppearance.ForeColor = Color.Black;
                    }
                }
            };


            view.Appearance.HeaderPanel.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            //view.Appearance.SelectedRow.BackColor = Color.LightYellow;
            //view.Appearance.SelectedRow.ForeColor = Color.Black;
            view.Appearance.FocusedRow.BackColor = Color.Transparent;
            view.Appearance.FocusedRow.ForeColor = Color.Black;
            view.OptionsNavigation.EnterMoveNextColumn = true;
            view.OptionsView.ColumnAutoWidth = false;
            view.OptionsView.ShowGroupPanel = false;
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

        #region [GridView 저장 전 빈 값 체크]
        /// <summary>
        /// GridView의 특정 컬럼에 빈 값이 있는지 체크하는 메서드
        /// 빈 값이 있다면, 해당 값의 행번-컬럼Caption명에 대한 입력요청 메시지를 리턴한다.
        /// 빈 값이 없다면, string.Empty를 리턴.
        /// </summary>
        /// <param name="gridView">체크하고자 하는 GridView</param>
        /// <param name="gridColumns">체크하고자 하는 컬럼 배열</param>
        /// <returns></returns>
        public static string GridViewEmptyValueCheck(GridView gridView, GridColumn[] gridColumns)
        {
            // strMessage가 Empty라면, 빈 값이 없습니다.
            string strMessage = string.Empty;
            Dictionary<int, GridColumn> dicValue = ValidateRows(gridView, gridColumns);
            if (dicValue.Count > 0)
            {
                int key = 0;
                GridColumn value = null;
                foreach (var item in dicValue)
                {
                    key = item.Key;
                    value = item.Value;
                }
                strMessage = ((key + 1) + " 번째 줄 " + value.Caption.ToString() + "의 값을 입력해 주세요.");
            }
            return strMessage;
        }
        /// <summary>
        /// GridViewEmptyValueCheck를 위한 메서드 
        /// </summary>
        /// <param name="view">체크하고자 하는 GridView</param>
        /// <param name="gridColumns">체크하고자 하는 컬럼 배열</param>
        /// <returns></returns>
        public static Dictionary<int, GridColumn> ValidateRows(GridView view, GridColumn[] gridColumns)
        {
            Dictionary<int, GridColumn> dicValue = new Dictionary<int, GridColumn>();

            if (view.RowCount > 0)
            {
                string[] strArr = new string[gridColumns.Length];
                for (int i = 0; i < view.RowCount; i++)
                {
                    for (int j = 0; j < gridColumns.Length; j++)
                    {
                        strArr[j] = view.GetRowCellValue(i, gridColumns[j])?.ToString();
                        if (string.IsNullOrEmpty(strArr[j]))
                        {
                            dicValue.Add(i, gridColumns[j]);
                            break;
                        }
                    }
                }
            }
            return dicValue;
        }

        /// <summary>
        /// Grid의 0값 체크를 위한 메서드
        /// </summary>
        /// <param name="gridView">체크할 gridview</param>
        /// <param name="gridColumns">체크할 gridColumns</param>
        /// <returns></returns>
        public static string GridViewZeroValueCheck(GridView gridView, GridColumn[] gridColumns)
        {
            // strMessage가 Empty라면, 빈 값이 없습니다.
            string strMessage = string.Empty;
            Dictionary<int, GridColumn> dicValue = ValidateRows_Zero(gridView, gridColumns);
            if (dicValue.Count > 0)
            {
                int key = 0;
                GridColumn value = null;
                foreach (var item in dicValue)
                {
                    key = item.Key;
                    value = item.Value;
                }
                strMessage = ((key + 1) + " 번째 줄 " + value.Caption.ToString() + "의 값을 입력해 주세요.");
            }
            return strMessage;
        }
        public static Dictionary<int, GridColumn> ValidateRows_Zero(GridView view, GridColumn[] gridColumns)
        {
            Dictionary<int, GridColumn> dicValue = new Dictionary<int, GridColumn>();
            double dCheck = 0;
            if (view.RowCount > 0)
            {
                string[] strArr = new string[gridColumns.Length];
                for (int i = 0; i < view.RowCount; i++)
                {
                    for (int j = 0; j < gridColumns.Length; j++)
                    {
                        strArr[j] = view.GetRowCellValue(i, gridColumns[j])?.ToString();
                        dCheck = double.TryParse(strArr[j], out dCheck) ? Convert.ToDouble(strArr[j]) : 0;
                        if (dCheck == 0)
                        {
                            dicValue.Add(i, gridColumns[j]);
                            break;
                        }
                    }
                }
            }
            return dicValue;
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
        #endregion
    }
}
