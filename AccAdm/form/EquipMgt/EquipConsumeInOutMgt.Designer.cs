namespace AccAdm
{
    partial class EquipConsumeInOutMgt
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.GridHistory = new DevExpress.XtraGrid.GridControl();
            this.GridViewHistory = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.GridColConsumeCd = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColConsumeSeq = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColConInOutGb = new DevExpress.XtraGrid.Columns.GridColumn();
            this.RepoConsumeGLkupInOutGb = new DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit();
            this.gridView2 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.GridColConOccurDt = new DevExpress.XtraGrid.Columns.GridColumn();
            this.RepoConsumeDateEditOccurDt = new DevExpress.XtraEditors.Repository.RepositoryItemDateEdit();
            this.GridColConAmt = new DevExpress.XtraGrid.Columns.GridColumn();
            this.RepoConsumeTxtAmt = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
            this.GridColConUnpr = new DevExpress.XtraGrid.Columns.GridColumn();
            this.RepoConsumeTxtUnpr = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
            this.GridColConPrice = new DevExpress.XtraGrid.Columns.GridColumn();
            this.RepoConsumeTxtPrice = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
            this.GridColConSuplDealer = new DevExpress.XtraGrid.Columns.GridColumn();
            this.RepoLkupCv = new DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit();
            this.repositoryItemGridLookUpEdit1View = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.GridColConStorage = new DevExpress.XtraGrid.Columns.GridColumn();
            this.RepoConsumeStorage = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
            this.GridColConNote = new DevExpress.XtraGrid.Columns.GridColumn();
            this.RepoConsumeNote = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit();
            this.GridColConEntDt = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColConEntId = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColConMfyDt = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColConMfyId = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColIdtNo = new DevExpress.XtraGrid.Columns.GridColumn();
            this.RepoConsumeBtnEditSuplDealer = new DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit();
            this.BtnHisNew = new DevExpress.XtraEditors.SimpleButton();
            this.BtnHisSave = new DevExpress.XtraEditors.SimpleButton();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem19 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem20 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.GridHistory)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewHistory)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeGLkupInOutGb)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeDateEditOccurDt)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeDateEditOccurDt.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeTxtAmt)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeTxtUnpr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeTxtPrice)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoLkupCv)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemGridLookUpEdit1View)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeStorage)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeNote)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeBtnEditSuplDealer)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem19)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem20)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.GridHistory);
            this.layoutControl1.Controls.Add(this.BtnHisNew);
            this.layoutControl1.Controls.Add(this.BtnHisSave);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(1216, 579);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // GridHistory
            // 
            this.GridHistory.Location = new System.Drawing.Point(24, 76);
            this.GridHistory.MainView = this.GridViewHistory;
            this.GridHistory.Name = "GridHistory";
            this.GridHistory.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.RepoConsumeDateEditOccurDt,
            this.RepoConsumeTxtAmt,
            this.RepoConsumeStorage,
            this.RepoConsumeNote,
            this.RepoConsumeTxtUnpr,
            this.RepoConsumeTxtPrice,
            this.RepoConsumeBtnEditSuplDealer,
            this.RepoConsumeGLkupInOutGb,
            this.RepoLkupCv});
            this.GridHistory.Size = new System.Drawing.Size(1168, 479);
            this.GridHistory.TabIndex = 5;
            this.GridHistory.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.GridViewHistory});
            this.GridHistory.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.GridHistory_PreviewKeyDown);
            // 
            // GridViewHistory
            // 
            this.GridViewHistory.ColumnPanelRowHeight = 30;
            this.GridViewHistory.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.GridColConsumeCd,
            this.GridColConsumeSeq,
            this.GridColConInOutGb,
            this.GridColConOccurDt,
            this.GridColConAmt,
            this.GridColConUnpr,
            this.GridColConPrice,
            this.GridColConSuplDealer,
            this.GridColConStorage,
            this.GridColConNote,
            this.GridColConEntDt,
            this.GridColConEntId,
            this.GridColConMfyDt,
            this.GridColConMfyId,
            this.GridColIdtNo});
            this.GridViewHistory.GridControl = this.GridHistory;
            this.GridViewHistory.Name = "GridViewHistory";
            this.GridViewHistory.OptionsBehavior.Editable = false;
            this.GridViewHistory.OptionsBehavior.EditingMode = DevExpress.XtraGrid.Views.Grid.GridEditingMode.Inplace;
            this.GridViewHistory.OptionsFind.AllowFindPanel = false;
            this.GridViewHistory.OptionsNavigation.EnterMoveNextColumn = true;
            this.GridViewHistory.OptionsView.ColumnAutoWidth = false;
            this.GridViewHistory.OptionsView.ShowGroupPanel = false;
            this.GridViewHistory.RowStyle += new DevExpress.XtraGrid.Views.Grid.RowStyleEventHandler(this.GridViewHistory_RowStyle);
            this.GridViewHistory.FocusedRowChanged += new DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventHandler(this.GridViewHistory_FocusedRowChanged);
            this.GridViewHistory.CellValueChanged += new DevExpress.XtraGrid.Views.Base.CellValueChangedEventHandler(this.GridViewHistory_CellValueChanged);
            // 
            // GridColConsumeCd
            // 
            this.GridColConsumeCd.AppearanceCell.Options.UseFont = true;
            this.GridColConsumeCd.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConsumeCd.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColConsumeCd.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConsumeCd.AppearanceHeader.Options.UseFont = true;
            this.GridColConsumeCd.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConsumeCd.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConsumeCd.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConsumeCd.Caption = "소모품코드";
            this.GridColConsumeCd.FieldName = "CON_ITEM_CD";
            this.GridColConsumeCd.Name = "GridColConsumeCd";
            this.GridColConsumeCd.OptionsEditForm.Visible = DevExpress.Utils.DefaultBoolean.False;
            this.GridColConsumeCd.Width = 100;
            // 
            // GridColConsumeSeq
            // 
            this.GridColConsumeSeq.AppearanceCell.Options.UseFont = true;
            this.GridColConsumeSeq.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConsumeSeq.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColConsumeSeq.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConsumeSeq.AppearanceHeader.Options.UseFont = true;
            this.GridColConsumeSeq.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConsumeSeq.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConsumeSeq.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConsumeSeq.Caption = "소모품행번";
            this.GridColConsumeSeq.FieldName = "CON_ITEM_SEQ";
            this.GridColConsumeSeq.Name = "GridColConsumeSeq";
            this.GridColConsumeSeq.OptionsEditForm.Visible = DevExpress.Utils.DefaultBoolean.False;
            // 
            // GridColConInOutGb
            // 
            this.GridColConInOutGb.AppearanceCell.Options.UseFont = true;
            this.GridColConInOutGb.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConInOutGb.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColConInOutGb.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConInOutGb.AppearanceHeader.Options.UseFont = true;
            this.GridColConInOutGb.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConInOutGb.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConInOutGb.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConInOutGb.Caption = "입출고";
            this.GridColConInOutGb.ColumnEdit = this.RepoConsumeGLkupInOutGb;
            this.GridColConInOutGb.FieldName = "ITEM_INOUT_GB";
            this.GridColConInOutGb.Name = "GridColConInOutGb";
            this.GridColConInOutGb.Visible = true;
            this.GridColConInOutGb.VisibleIndex = 0;
            this.GridColConInOutGb.Width = 108;
            // 
            // RepoConsumeGLkupInOutGb
            // 
            this.RepoConsumeGLkupInOutGb.Appearance.Options.UseFont = true;
            this.RepoConsumeGLkupInOutGb.AppearanceDisabled.Options.UseFont = true;
            this.RepoConsumeGLkupInOutGb.AppearanceDropDown.Options.UseFont = true;
            this.RepoConsumeGLkupInOutGb.AppearanceFocused.Options.UseFont = true;
            this.RepoConsumeGLkupInOutGb.AppearanceReadOnly.Options.UseFont = true;
            this.RepoConsumeGLkupInOutGb.AutoHeight = false;
            this.RepoConsumeGLkupInOutGb.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.RepoConsumeGLkupInOutGb.Name = "RepoConsumeGLkupInOutGb";
            this.RepoConsumeGLkupInOutGb.NullText = "";
            this.RepoConsumeGLkupInOutGb.PopupView = this.gridView2;
            this.RepoConsumeGLkupInOutGb.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.Standard;
            this.RepoConsumeGLkupInOutGb.Leave += new System.EventHandler(this.RepoConsumeGLkupInOutGb_Leave);
            // 
            // gridView2
            // 
            this.gridView2.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus;
            this.gridView2.Name = "gridView2";
            this.gridView2.OptionsSelection.EnableAppearanceFocusedCell = false;
            this.gridView2.OptionsView.ShowColumnHeaders = false;
            this.gridView2.OptionsView.ShowGroupPanel = false;
            // 
            // GridColConOccurDt
            // 
            this.GridColConOccurDt.AppearanceCell.Options.UseFont = true;
            this.GridColConOccurDt.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConOccurDt.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConOccurDt.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConOccurDt.AppearanceHeader.Options.UseFont = true;
            this.GridColConOccurDt.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConOccurDt.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConOccurDt.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConOccurDt.Caption = "발생일자";
            this.GridColConOccurDt.ColumnEdit = this.RepoConsumeDateEditOccurDt;
            this.GridColConOccurDt.FieldName = "OCCUR_DT";
            this.GridColConOccurDt.Name = "GridColConOccurDt";
            this.GridColConOccurDt.Visible = true;
            this.GridColConOccurDt.VisibleIndex = 1;
            this.GridColConOccurDt.Width = 142;
            // 
            // RepoConsumeDateEditOccurDt
            // 
            this.RepoConsumeDateEditOccurDt.AllowNullInput = DevExpress.Utils.DefaultBoolean.False;
            this.RepoConsumeDateEditOccurDt.Appearance.Options.UseFont = true;
            this.RepoConsumeDateEditOccurDt.AppearanceDisabled.Options.UseFont = true;
            this.RepoConsumeDateEditOccurDt.AppearanceDropDown.Options.UseFont = true;
            this.RepoConsumeDateEditOccurDt.AppearanceFocused.Options.UseFont = true;
            this.RepoConsumeDateEditOccurDt.AppearanceReadOnly.Options.UseFont = true;
            this.RepoConsumeDateEditOccurDt.AutoHeight = false;
            this.RepoConsumeDateEditOccurDt.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.RepoConsumeDateEditOccurDt.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.RepoConsumeDateEditOccurDt.Mask.EditMask = "yyyy-MM-dd";
            this.RepoConsumeDateEditOccurDt.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTimeAdvancingCaret;
            this.RepoConsumeDateEditOccurDt.Mask.UseMaskAsDisplayFormat = true;
            this.RepoConsumeDateEditOccurDt.Name = "RepoConsumeDateEditOccurDt";
            this.RepoConsumeDateEditOccurDt.NullValuePromptShowForEmptyValue = true;
            this.RepoConsumeDateEditOccurDt.EditValueChanged += new System.EventHandler(this.RepoConsumeDateEditOccurDt_EditValueChanged);
            // 
            // GridColConAmt
            // 
            this.GridColConAmt.AppearanceCell.Options.UseFont = true;
            this.GridColConAmt.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConAmt.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.GridColConAmt.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConAmt.AppearanceHeader.Options.UseFont = true;
            this.GridColConAmt.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConAmt.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConAmt.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConAmt.Caption = "수량";
            this.GridColConAmt.ColumnEdit = this.RepoConsumeTxtAmt;
            this.GridColConAmt.FieldName = "CON_ITEM_AMT";
            this.GridColConAmt.Name = "GridColConAmt";
            this.GridColConAmt.Visible = true;
            this.GridColConAmt.VisibleIndex = 2;
            this.GridColConAmt.Width = 111;
            // 
            // RepoConsumeTxtAmt
            // 
            this.RepoConsumeTxtAmt.Appearance.Options.UseFont = true;
            this.RepoConsumeTxtAmt.AppearanceDisabled.Options.UseFont = true;
            this.RepoConsumeTxtAmt.AppearanceFocused.Options.UseFont = true;
            this.RepoConsumeTxtAmt.AppearanceReadOnly.Options.UseFont = true;
            this.RepoConsumeTxtAmt.AutoHeight = false;
            this.RepoConsumeTxtAmt.Mask.EditMask = "n0";
            this.RepoConsumeTxtAmt.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
            this.RepoConsumeTxtAmt.Mask.UseMaskAsDisplayFormat = true;
            this.RepoConsumeTxtAmt.Name = "RepoConsumeTxtAmt";
            this.RepoConsumeTxtAmt.Leave += new System.EventHandler(this.RepoConsumeTxtAmt_Leave);
            // 
            // GridColConUnpr
            // 
            this.GridColConUnpr.AppearanceCell.Options.UseFont = true;
            this.GridColConUnpr.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConUnpr.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.GridColConUnpr.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConUnpr.AppearanceHeader.Options.UseFont = true;
            this.GridColConUnpr.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConUnpr.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConUnpr.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConUnpr.Caption = "단가";
            this.GridColConUnpr.ColumnEdit = this.RepoConsumeTxtUnpr;
            this.GridColConUnpr.FieldName = "CON_ITEM_UNPR";
            this.GridColConUnpr.Name = "GridColConUnpr";
            this.GridColConUnpr.Visible = true;
            this.GridColConUnpr.VisibleIndex = 3;
            this.GridColConUnpr.Width = 102;
            // 
            // RepoConsumeTxtUnpr
            // 
            this.RepoConsumeTxtUnpr.Appearance.Options.UseFont = true;
            this.RepoConsumeTxtUnpr.AppearanceDisabled.Options.UseFont = true;
            this.RepoConsumeTxtUnpr.AppearanceFocused.Options.UseFont = true;
            this.RepoConsumeTxtUnpr.AppearanceReadOnly.Options.UseFont = true;
            this.RepoConsumeTxtUnpr.AutoHeight = false;
            this.RepoConsumeTxtUnpr.Mask.EditMask = "n0";
            this.RepoConsumeTxtUnpr.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
            this.RepoConsumeTxtUnpr.Mask.UseMaskAsDisplayFormat = true;
            this.RepoConsumeTxtUnpr.Name = "RepoConsumeTxtUnpr";
            this.RepoConsumeTxtUnpr.Leave += new System.EventHandler(this.RepoConsumeTxtUnpr_Leave);
            // 
            // GridColConPrice
            // 
            this.GridColConPrice.AppearanceCell.Options.UseFont = true;
            this.GridColConPrice.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConPrice.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.GridColConPrice.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConPrice.AppearanceHeader.Options.UseFont = true;
            this.GridColConPrice.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConPrice.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConPrice.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConPrice.Caption = "금액";
            this.GridColConPrice.ColumnEdit = this.RepoConsumeTxtPrice;
            this.GridColConPrice.FieldName = "CON_ITEM_PRICE";
            this.GridColConPrice.Name = "GridColConPrice";
            this.GridColConPrice.Visible = true;
            this.GridColConPrice.VisibleIndex = 4;
            this.GridColConPrice.Width = 196;
            // 
            // RepoConsumeTxtPrice
            // 
            this.RepoConsumeTxtPrice.Appearance.Options.UseFont = true;
            this.RepoConsumeTxtPrice.AppearanceDisabled.Options.UseFont = true;
            this.RepoConsumeTxtPrice.AppearanceFocused.Options.UseFont = true;
            this.RepoConsumeTxtPrice.AppearanceReadOnly.Options.UseFont = true;
            this.RepoConsumeTxtPrice.AutoHeight = false;
            this.RepoConsumeTxtPrice.Mask.EditMask = "n0";
            this.RepoConsumeTxtPrice.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
            this.RepoConsumeTxtPrice.Mask.UseMaskAsDisplayFormat = true;
            this.RepoConsumeTxtPrice.Name = "RepoConsumeTxtPrice";
            this.RepoConsumeTxtPrice.Leave += new System.EventHandler(this.RepoConsumeTxtPrice_Leave);
            // 
            // GridColConSuplDealer
            // 
            this.GridColConSuplDealer.AppearanceCell.Options.UseFont = true;
            this.GridColConSuplDealer.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConSuplDealer.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColConSuplDealer.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConSuplDealer.AppearanceHeader.Options.UseFont = true;
            this.GridColConSuplDealer.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConSuplDealer.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConSuplDealer.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConSuplDealer.Caption = "소모품 공급처";
            this.GridColConSuplDealer.ColumnEdit = this.RepoLkupCv;
            this.GridColConSuplDealer.FieldName = "CON_SUPL_DEALER";
            this.GridColConSuplDealer.Name = "GridColConSuplDealer";
            this.GridColConSuplDealer.Visible = true;
            this.GridColConSuplDealer.VisibleIndex = 5;
            this.GridColConSuplDealer.Width = 240;
            // 
            // RepoLkupCv
            // 
            this.RepoLkupCv.AutoHeight = false;
            this.RepoLkupCv.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.RepoLkupCv.Name = "RepoLkupCv";
            this.RepoLkupCv.NullText = "";
            this.RepoLkupCv.PopupView = this.repositoryItemGridLookUpEdit1View;
            this.RepoLkupCv.ShowFooter = false;
            // 
            // repositoryItemGridLookUpEdit1View
            // 
            this.repositoryItemGridLookUpEdit1View.FocusRectStyle = DevExpress.XtraGrid.Views.Grid.DrawFocusRectStyle.RowFocus;
            this.repositoryItemGridLookUpEdit1View.Name = "repositoryItemGridLookUpEdit1View";
            this.repositoryItemGridLookUpEdit1View.OptionsSelection.EnableAppearanceFocusedCell = false;
            this.repositoryItemGridLookUpEdit1View.OptionsView.ShowColumnHeaders = false;
            this.repositoryItemGridLookUpEdit1View.OptionsView.ShowGroupPanel = false;
            // 
            // GridColConStorage
            // 
            this.GridColConStorage.AppearanceCell.Options.UseFont = true;
            this.GridColConStorage.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConStorage.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColConStorage.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConStorage.AppearanceHeader.Options.UseFont = true;
            this.GridColConStorage.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConStorage.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConStorage.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConStorage.Caption = "보관장소";
            this.GridColConStorage.ColumnEdit = this.RepoConsumeStorage;
            this.GridColConStorage.FieldName = "CON_STORAGE";
            this.GridColConStorage.Name = "GridColConStorage";
            this.GridColConStorage.Visible = true;
            this.GridColConStorage.VisibleIndex = 6;
            this.GridColConStorage.Width = 100;
            // 
            // RepoConsumeStorage
            // 
            this.RepoConsumeStorage.Appearance.Options.UseFont = true;
            this.RepoConsumeStorage.AppearanceDisabled.Options.UseFont = true;
            this.RepoConsumeStorage.AppearanceFocused.Options.UseFont = true;
            this.RepoConsumeStorage.AppearanceReadOnly.Options.UseFont = true;
            this.RepoConsumeStorage.AutoHeight = false;
            this.RepoConsumeStorage.MaxLength = 20;
            this.RepoConsumeStorage.Name = "RepoConsumeStorage";
            // 
            // GridColConNote
            // 
            this.GridColConNote.AppearanceCell.Options.UseFont = true;
            this.GridColConNote.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConNote.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColConNote.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConNote.AppearanceHeader.Options.UseFont = true;
            this.GridColConNote.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConNote.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConNote.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConNote.Caption = "비고";
            this.GridColConNote.ColumnEdit = this.RepoConsumeNote;
            this.GridColConNote.FieldName = "NOTE";
            this.GridColConNote.Name = "GridColConNote";
            this.GridColConNote.Visible = true;
            this.GridColConNote.VisibleIndex = 7;
            this.GridColConNote.Width = 150;
            // 
            // RepoConsumeNote
            // 
            this.RepoConsumeNote.Appearance.Options.UseFont = true;
            this.RepoConsumeNote.AppearanceDisabled.Options.UseFont = true;
            this.RepoConsumeNote.AppearanceFocused.Options.UseFont = true;
            this.RepoConsumeNote.AppearanceReadOnly.Options.UseFont = true;
            this.RepoConsumeNote.AutoHeight = false;
            this.RepoConsumeNote.MaxLength = 100;
            this.RepoConsumeNote.Name = "RepoConsumeNote";
            // 
            // GridColConEntDt
            // 
            this.GridColConEntDt.AppearanceCell.Options.UseFont = true;
            this.GridColConEntDt.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConEntDt.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConEntDt.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConEntDt.AppearanceHeader.Options.UseFont = true;
            this.GridColConEntDt.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConEntDt.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConEntDt.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConEntDt.Caption = "입력일시";
            this.GridColConEntDt.FieldName = "ENT_DT";
            this.GridColConEntDt.Name = "GridColConEntDt";
            this.GridColConEntDt.OptionsColumn.AllowEdit = false;
            this.GridColConEntDt.OptionsColumn.AllowFocus = false;
            this.GridColConEntDt.OptionsColumn.ReadOnly = true;
            this.GridColConEntDt.OptionsEditForm.Visible = DevExpress.Utils.DefaultBoolean.False;
            this.GridColConEntDt.Visible = true;
            this.GridColConEntDt.VisibleIndex = 8;
            this.GridColConEntDt.Width = 100;
            // 
            // GridColConEntId
            // 
            this.GridColConEntId.AppearanceCell.Options.UseFont = true;
            this.GridColConEntId.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConEntId.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConEntId.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConEntId.AppearanceHeader.Options.UseFont = true;
            this.GridColConEntId.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConEntId.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConEntId.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConEntId.Caption = "입력ID";
            this.GridColConEntId.FieldName = "ENT_ID";
            this.GridColConEntId.Name = "GridColConEntId";
            this.GridColConEntId.OptionsColumn.AllowEdit = false;
            this.GridColConEntId.OptionsColumn.AllowFocus = false;
            this.GridColConEntId.OptionsColumn.ReadOnly = true;
            this.GridColConEntId.OptionsEditForm.Visible = DevExpress.Utils.DefaultBoolean.False;
            this.GridColConEntId.Visible = true;
            this.GridColConEntId.VisibleIndex = 9;
            this.GridColConEntId.Width = 100;
            // 
            // GridColConMfyDt
            // 
            this.GridColConMfyDt.AppearanceCell.Options.UseFont = true;
            this.GridColConMfyDt.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConMfyDt.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConMfyDt.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConMfyDt.AppearanceHeader.Options.UseFont = true;
            this.GridColConMfyDt.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConMfyDt.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConMfyDt.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConMfyDt.Caption = "수정일시";
            this.GridColConMfyDt.FieldName = "MFY_DT";
            this.GridColConMfyDt.Name = "GridColConMfyDt";
            this.GridColConMfyDt.OptionsColumn.AllowEdit = false;
            this.GridColConMfyDt.OptionsColumn.AllowFocus = false;
            this.GridColConMfyDt.OptionsColumn.ReadOnly = true;
            this.GridColConMfyDt.OptionsEditForm.Visible = DevExpress.Utils.DefaultBoolean.False;
            this.GridColConMfyDt.Visible = true;
            this.GridColConMfyDt.VisibleIndex = 10;
            this.GridColConMfyDt.Width = 100;
            // 
            // GridColConMfyId
            // 
            this.GridColConMfyId.AppearanceCell.Options.UseFont = true;
            this.GridColConMfyId.AppearanceCell.Options.UseTextOptions = true;
            this.GridColConMfyId.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConMfyId.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConMfyId.AppearanceHeader.Options.UseFont = true;
            this.GridColConMfyId.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColConMfyId.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColConMfyId.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColConMfyId.Caption = "수정ID";
            this.GridColConMfyId.FieldName = "MFY_ID";
            this.GridColConMfyId.Name = "GridColConMfyId";
            this.GridColConMfyId.OptionsColumn.AllowEdit = false;
            this.GridColConMfyId.OptionsColumn.AllowFocus = false;
            this.GridColConMfyId.OptionsColumn.ReadOnly = true;
            this.GridColConMfyId.OptionsEditForm.Visible = DevExpress.Utils.DefaultBoolean.False;
            this.GridColConMfyId.Visible = true;
            this.GridColConMfyId.VisibleIndex = 11;
            this.GridColConMfyId.Width = 100;
            // 
            // GridColIdtNo
            // 
            this.GridColIdtNo.AppearanceCell.Options.UseFont = true;
            this.GridColIdtNo.AppearanceCell.Options.UseTextOptions = true;
            this.GridColIdtNo.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColIdtNo.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColIdtNo.AppearanceHeader.Options.UseFont = true;
            this.GridColIdtNo.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColIdtNo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColIdtNo.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColIdtNo.Caption = "거래처식별번호";
            this.GridColIdtNo.FieldName = "IDT_NO";
            this.GridColIdtNo.Name = "GridColIdtNo";
            // 
            // RepoConsumeBtnEditSuplDealer
            // 
            this.RepoConsumeBtnEditSuplDealer.Appearance.Options.UseFont = true;
            this.RepoConsumeBtnEditSuplDealer.AppearanceDisabled.Options.UseFont = true;
            this.RepoConsumeBtnEditSuplDealer.AppearanceFocused.Options.UseFont = true;
            this.RepoConsumeBtnEditSuplDealer.AppearanceReadOnly.Options.UseFont = true;
            this.RepoConsumeBtnEditSuplDealer.AutoHeight = false;
            this.RepoConsumeBtnEditSuplDealer.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton()});
            this.RepoConsumeBtnEditSuplDealer.Name = "RepoConsumeBtnEditSuplDealer";
            this.RepoConsumeBtnEditSuplDealer.ButtonClick += new DevExpress.XtraEditors.Controls.ButtonPressedEventHandler(this.RepoConsumeBtnEditSuplDealer_ButtonClick);
            // 
            // BtnHisNew
            // 
            this.BtnHisNew.Appearance.Options.UseFont = true;
            this.BtnHisNew.AppearanceDisabled.Options.UseFont = true;
            this.BtnHisNew.AppearanceHovered.Options.UseFont = true;
            this.BtnHisNew.AppearancePressed.Options.UseFont = true;
            this.BtnHisNew.ImageOptions.Image = global::AccAdm.Properties.Resources.add_16x16;
            this.BtnHisNew.Location = new System.Drawing.Point(1028, 44);
            this.BtnHisNew.Name = "BtnHisNew";
            this.BtnHisNew.Size = new System.Drawing.Size(80, 28);
            this.BtnHisNew.StyleController = this.layoutControl1;
            this.BtnHisNew.TabIndex = 4;
            this.BtnHisNew.TabStop = false;
            this.BtnHisNew.Text = "신규";
            this.BtnHisNew.Click += new System.EventHandler(this.BtnHisNew_Click);
            // 
            // BtnHisSave
            // 
            this.BtnHisSave.Appearance.Options.UseFont = true;
            this.BtnHisSave.AppearanceDisabled.Options.UseFont = true;
            this.BtnHisSave.AppearanceHovered.Options.UseFont = true;
            this.BtnHisSave.AppearancePressed.Options.UseFont = true;
            this.BtnHisSave.ImageOptions.Image = global::AccAdm.Properties.Resources.save_16x16;
            this.BtnHisSave.Location = new System.Drawing.Point(1112, 44);
            this.BtnHisSave.Name = "BtnHisSave";
            this.BtnHisSave.Size = new System.Drawing.Size(80, 28);
            this.BtnHisSave.StyleController = this.layoutControl1;
            this.BtnHisSave.TabIndex = 3;
            this.BtnHisSave.TabStop = false;
            this.BtnHisSave.Text = "저장";
            this.BtnHisSave.Click += new System.EventHandler(this.BtnHisSave_Click);
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlGroup1});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(1216, 579);
            this.Root.TextVisible = false;
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.AppearanceGroup.Options.UseFont = true;
            this.layoutControlGroup1.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlGroup1.AppearanceTabPage.Header.Options.UseFont = true;
            this.layoutControlGroup1.AppearanceTabPage.HeaderActive.Options.UseFont = true;
            this.layoutControlGroup1.AppearanceTabPage.HeaderDisabled.Options.UseFont = true;
            this.layoutControlGroup1.AppearanceTabPage.HeaderHotTracked.Options.UseFont = true;
            this.layoutControlGroup1.AppearanceTabPage.PageClient.Options.UseFont = true;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem19,
            this.layoutControlItem20,
            this.emptySpaceItem1,
            this.layoutControlItem2});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.OptionsPrint.AppearanceGroupCaption.Options.UseFont = true;
            this.layoutControlGroup1.OptionsPrint.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlGroup1.Size = new System.Drawing.Size(1196, 559);
            this.layoutControlGroup1.Text = "입출고 등록";
            // 
            // layoutControlItem19
            // 
            this.layoutControlItem19.Control = this.BtnHisNew;
            this.layoutControlItem19.ControlAlignment = System.Drawing.ContentAlignment.TopCenter;
            this.layoutControlItem19.CustomizationFormText = "layoutControlItem13";
            this.layoutControlItem19.Location = new System.Drawing.Point(1004, 0);
            this.layoutControlItem19.MaxSize = new System.Drawing.Size(84, 32);
            this.layoutControlItem19.MinSize = new System.Drawing.Size(84, 32);
            this.layoutControlItem19.Name = "layoutControlItem19";
            this.layoutControlItem19.Size = new System.Drawing.Size(84, 32);
            this.layoutControlItem19.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem19.Text = "layoutControlItem13";
            this.layoutControlItem19.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem19.TextVisible = false;
            this.layoutControlItem19.TrimClientAreaToControl = false;
            // 
            // layoutControlItem20
            // 
            this.layoutControlItem20.Control = this.BtnHisSave;
            this.layoutControlItem20.ControlAlignment = System.Drawing.ContentAlignment.TopCenter;
            this.layoutControlItem20.CustomizationFormText = "layoutControlItem12";
            this.layoutControlItem20.Location = new System.Drawing.Point(1088, 0);
            this.layoutControlItem20.MaxSize = new System.Drawing.Size(84, 32);
            this.layoutControlItem20.MinSize = new System.Drawing.Size(84, 32);
            this.layoutControlItem20.Name = "layoutControlItem20";
            this.layoutControlItem20.Size = new System.Drawing.Size(84, 32);
            this.layoutControlItem20.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem20.Text = "layoutControlItem12";
            this.layoutControlItem20.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem20.TextVisible = false;
            this.layoutControlItem20.TrimClientAreaToControl = false;
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(0, 0);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(1004, 32);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.GridHistory;
            this.layoutControlItem2.CustomizationFormText = "layoutControlItem2";
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 32);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(1172, 483);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // EquipConsumeInOutMgt
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1216, 579);
            this.Controls.Add(this.layoutControl1);
            this.KeyPreview = true;
            this.Name = "EquipConsumeInOutMgt";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "소모품입출고등록";
            this.Load += new System.EventHandler(this.EquipConsumeInOutMgt_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.EquipConsumeInOutMgt_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.GridHistory)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewHistory)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeGLkupInOutGb)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeDateEditOccurDt.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeDateEditOccurDt)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeTxtAmt)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeTxtUnpr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeTxtPrice)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoLkupCv)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemGridLookUpEdit1View)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeStorage)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeNote)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoConsumeBtnEditSuplDealer)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem19)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem20)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraGrid.GridControl GridHistory;
        private DevExpress.XtraGrid.Views.Grid.GridView GridViewHistory;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConsumeCd;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConsumeSeq;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConInOutGb;
        private DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit RepoConsumeGLkupInOutGb;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView2;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConOccurDt;
        private DevExpress.XtraEditors.Repository.RepositoryItemDateEdit RepoConsumeDateEditOccurDt;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConAmt;
        private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit RepoConsumeTxtAmt;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConUnpr;
        private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit RepoConsumeTxtPrice;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConPrice;
        private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit RepoConsumeTxtUnpr;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConSuplDealer;
        private DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit RepoConsumeBtnEditSuplDealer;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConStorage;
        private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit RepoConsumeStorage;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConNote;
        private DevExpress.XtraEditors.Repository.RepositoryItemTextEdit RepoConsumeNote;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConEntDt;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConEntId;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConMfyDt;
        private DevExpress.XtraGrid.Columns.GridColumn GridColConMfyId;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraEditors.SimpleButton BtnHisNew;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem19;
        private DevExpress.XtraEditors.SimpleButton BtnHisSave;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem20;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraGrid.Columns.GridColumn GridColIdtNo;
        private DevExpress.XtraEditors.Repository.RepositoryItemGridLookUpEdit RepoLkupCv;
        private DevExpress.XtraGrid.Views.Grid.GridView repositoryItemGridLookUpEdit1View;
    }
}