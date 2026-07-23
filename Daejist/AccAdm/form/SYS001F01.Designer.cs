namespace AccAdm
{
    partial class SYS001F01
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SYS001F01));
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.GridRetr = new DevExpress.XtraGrid.GridControl();
            this.GridViewRetr = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.GridColVersionId = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColUploadDt = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColFile = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColFileName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColFileByte = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColVersionRmk = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColEntId = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColEntDt = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColModId = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColModDt = new DevExpress.XtraGrid.Columns.GridColumn();
            this.RepoDateEditTime = new DevExpress.XtraEditors.Repository.RepositoryItemDateEdit();
            this.BtnRetr = new DevExpress.XtraEditors.SimpleButton();
            this.DateEditFrom = new DevExpress.XtraEditors.DateEdit();
            this.DateEditTo = new DevExpress.XtraEditors.DateEdit();
            this.BtnAdd = new DevExpress.XtraEditors.SimpleButton();
            this.BtnClose = new DevExpress.XtraEditors.SimpleButton();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlGroup2 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.GridRetr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewRetr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoDateEditTime)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoDateEditTime.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DateEditFrom.Properties.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DateEditFrom.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DateEditTo.Properties.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DateEditTo.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.GridRetr);
            this.layoutControl1.Controls.Add(this.BtnRetr);
            this.layoutControl1.Controls.Add(this.DateEditFrom);
            this.layoutControl1.Controls.Add(this.DateEditTo);
            this.layoutControl1.Controls.Add(this.BtnAdd);
            this.layoutControl1.Controls.Add(this.BtnClose);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(1250, 888);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // GridRetr
            // 
            this.GridRetr.Location = new System.Drawing.Point(24, 122);
            this.GridRetr.MainView = this.GridViewRetr;
            this.GridRetr.Name = "GridRetr";
            this.GridRetr.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.RepoDateEditTime});
            this.GridRetr.Size = new System.Drawing.Size(1202, 742);
            this.GridRetr.TabIndex = 6;
            this.GridRetr.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.GridViewRetr});
            // 
            // GridViewRetr
            // 
            this.GridViewRetr.ColumnPanelRowHeight = 30;
            this.GridViewRetr.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.GridColVersionId,
            this.GridColUploadDt,
            this.GridColFile,
            this.GridColFileName,
            this.GridColFileByte,
            this.GridColVersionRmk,
            this.GridColEntId,
            this.GridColEntDt,
            this.GridColModId,
            this.GridColModDt});
            this.GridViewRetr.GridControl = this.GridRetr;
            this.GridViewRetr.Name = "GridViewRetr";
            this.GridViewRetr.OptionsView.ColumnAutoWidth = false;
            this.GridViewRetr.OptionsView.ShowGroupPanel = false;
            this.GridViewRetr.RowClick += new DevExpress.XtraGrid.Views.Grid.RowClickEventHandler(this.GridViewRetr_RowClick);
            this.GridViewRetr.RowStyle += new DevExpress.XtraGrid.Views.Grid.RowStyleEventHandler(this.GridViewRetr_RowStyle);
            // 
            // GridColVersionId
            // 
            this.GridColVersionId.AppearanceCell.Options.UseTextOptions = true;
            this.GridColVersionId.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColVersionId.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColVersionId.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColVersionId.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColVersionId.Caption = "버전번호";
            this.GridColVersionId.FieldName = "VERSION_ID";
            this.GridColVersionId.Name = "GridColVersionId";
            this.GridColVersionId.OptionsColumn.AllowEdit = false;
            this.GridColVersionId.Visible = true;
            this.GridColVersionId.VisibleIndex = 0;
            this.GridColVersionId.Width = 107;
            // 
            // GridColUploadDt
            // 
            this.GridColUploadDt.AppearanceCell.Options.UseTextOptions = true;
            this.GridColUploadDt.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColUploadDt.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColUploadDt.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColUploadDt.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColUploadDt.Caption = "업로드일자";
            this.GridColUploadDt.FieldName = "UPLOAD_DT";
            this.GridColUploadDt.Name = "GridColUploadDt";
            this.GridColUploadDt.OptionsColumn.AllowEdit = false;
            this.GridColUploadDt.Visible = true;
            this.GridColUploadDt.VisibleIndex = 1;
            this.GridColUploadDt.Width = 110;
            // 
            // GridColFile
            // 
            this.GridColFile.AppearanceCell.Options.UseTextOptions = true;
            this.GridColFile.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColFile.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColFile.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColFile.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColFile.Caption = "파일";
            this.GridColFile.FieldName = "EXEFILE";
            this.GridColFile.Name = "GridColFile";
            this.GridColFile.OptionsColumn.AllowEdit = false;
            this.GridColFile.Visible = true;
            this.GridColFile.VisibleIndex = 2;
            this.GridColFile.Width = 113;
            // 
            // GridColFileName
            // 
            this.GridColFileName.AppearanceCell.Options.UseTextOptions = true;
            this.GridColFileName.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColFileName.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColFileName.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColFileName.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColFileName.Caption = "파일명";
            this.GridColFileName.FieldName = "FILE_NAME";
            this.GridColFileName.Name = "GridColFileName";
            this.GridColFileName.OptionsColumn.AllowEdit = false;
            this.GridColFileName.Visible = true;
            this.GridColFileName.VisibleIndex = 3;
            this.GridColFileName.Width = 146;
            // 
            // GridColFileByte
            // 
            this.GridColFileByte.AppearanceCell.Options.UseTextOptions = true;
            this.GridColFileByte.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColFileByte.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColFileByte.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColFileByte.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColFileByte.Caption = "크기";
            this.GridColFileByte.FieldName = "FILE_BYTE";
            this.GridColFileByte.Name = "GridColFileByte";
            this.GridColFileByte.OptionsColumn.AllowEdit = false;
            this.GridColFileByte.Visible = true;
            this.GridColFileByte.VisibleIndex = 4;
            this.GridColFileByte.Width = 114;
            // 
            // GridColVersionRmk
            // 
            this.GridColVersionRmk.AppearanceCell.Options.UseTextOptions = true;
            this.GridColVersionRmk.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColVersionRmk.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColVersionRmk.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColVersionRmk.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColVersionRmk.Caption = "비고";
            this.GridColVersionRmk.FieldName = "VERSION_RMK";
            this.GridColVersionRmk.Name = "GridColVersionRmk";
            this.GridColVersionRmk.OptionsColumn.AllowEdit = false;
            this.GridColVersionRmk.Visible = true;
            this.GridColVersionRmk.VisibleIndex = 5;
            this.GridColVersionRmk.Width = 212;
            // 
            // GridColEntId
            // 
            this.GridColEntId.AppearanceCell.Options.UseTextOptions = true;
            this.GridColEntId.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColEntId.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColEntId.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColEntId.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColEntId.Caption = "입력ID";
            this.GridColEntId.FieldName = "ENT_ID";
            this.GridColEntId.Name = "GridColEntId";
            this.GridColEntId.OptionsColumn.AllowEdit = false;
            this.GridColEntId.Visible = true;
            this.GridColEntId.VisibleIndex = 6;
            this.GridColEntId.Width = 85;
            // 
            // GridColEntDt
            // 
            this.GridColEntDt.AppearanceCell.Options.UseTextOptions = true;
            this.GridColEntDt.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColEntDt.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColEntDt.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColEntDt.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColEntDt.Caption = "입력일시";
            this.GridColEntDt.FieldName = "ENT_DT";
            this.GridColEntDt.Name = "GridColEntDt";
            this.GridColEntDt.OptionsColumn.AllowEdit = false;
            this.GridColEntDt.Visible = true;
            this.GridColEntDt.VisibleIndex = 7;
            this.GridColEntDt.Width = 180;
            // 
            // GridColModId
            // 
            this.GridColModId.AppearanceCell.Options.UseTextOptions = true;
            this.GridColModId.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColModId.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColModId.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColModId.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColModId.Caption = "수정ID";
            this.GridColModId.FieldName = "MOD_ID";
            this.GridColModId.Name = "GridColModId";
            this.GridColModId.OptionsColumn.AllowEdit = false;
            this.GridColModId.Visible = true;
            this.GridColModId.VisibleIndex = 8;
            // 
            // GridColModDt
            // 
            this.GridColModDt.AppearanceCell.Options.UseTextOptions = true;
            this.GridColModDt.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColModDt.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColModDt.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColModDt.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColModDt.Caption = "수정일시";
            this.GridColModDt.FieldName = "MOD_DT";
            this.GridColModDt.Name = "GridColModDt";
            this.GridColModDt.OptionsColumn.AllowEdit = false;
            this.GridColModDt.Visible = true;
            this.GridColModDt.VisibleIndex = 9;
            this.GridColModDt.Width = 180;
            // 
            // RepoDateEditTime
            // 
            this.RepoDateEditTime.AutoHeight = false;
            this.RepoDateEditTime.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.RepoDateEditTime.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.RepoDateEditTime.Mask.EditMask = "G";
            this.RepoDateEditTime.Mask.UseMaskAsDisplayFormat = true;
            this.RepoDateEditTime.Name = "RepoDateEditTime";
            // 
            // BtnRetr
            // 
            this.BtnRetr.ImageOptions.Image = global::AccAdm.Properties.Resources.zoom_16x16;
            this.BtnRetr.Location = new System.Drawing.Point(963, 44);
            this.BtnRetr.Name = "BtnRetr";
            this.BtnRetr.Size = new System.Drawing.Size(85, 30);
            this.BtnRetr.StyleController = this.layoutControl1;
            this.BtnRetr.TabIndex = 5;
            this.BtnRetr.Text = "조회(F5)";
            this.BtnRetr.Click += new System.EventHandler(this.BtnRetr_Click);
            // 
            // DateEditFrom
            // 
            this.DateEditFrom.EditValue = null;
            this.DateEditFrom.EnterMoveNextControl = true;
            this.DateEditFrom.Location = new System.Drawing.Point(87, 44);
            this.DateEditFrom.Name = "DateEditFrom";
            this.DateEditFrom.Properties.AppearanceFocused.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.DateEditFrom.Properties.AppearanceFocused.Options.UseBackColor = true;
            this.DateEditFrom.Properties.AutoHeight = false;
            this.DateEditFrom.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.DateEditFrom.Properties.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.DateEditFrom.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTimeAdvancingCaret;
            this.DateEditFrom.Properties.Mask.UseMaskAsDisplayFormat = true;
            this.DateEditFrom.Size = new System.Drawing.Size(101, 22);
            this.DateEditFrom.StyleController = this.layoutControl1;
            this.DateEditFrom.TabIndex = 4;
            // 
            // DateEditTo
            // 
            this.DateEditTo.EditValue = null;
            this.DateEditTo.EnterMoveNextControl = true;
            this.DateEditTo.Location = new System.Drawing.Point(204, 44);
            this.DateEditTo.Name = "DateEditTo";
            this.DateEditTo.Properties.AppearanceFocused.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.DateEditTo.Properties.AppearanceFocused.Options.UseBackColor = true;
            this.DateEditTo.Properties.AutoHeight = false;
            this.DateEditTo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.DateEditTo.Properties.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.DateEditTo.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTimeAdvancingCaret;
            this.DateEditTo.Properties.Mask.UseMaskAsDisplayFormat = true;
            this.DateEditTo.Size = new System.Drawing.Size(100, 22);
            this.DateEditTo.StyleController = this.layoutControl1;
            this.DateEditTo.TabIndex = 4;
            this.DateEditTo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.DateEditTo_KeyDown);
            // 
            // BtnAdd
            // 
            this.BtnAdd.ImageOptions.Image = global::AccAdm.Properties.Resources.addfile_16x16;
            this.BtnAdd.Location = new System.Drawing.Point(1052, 44);
            this.BtnAdd.Name = "BtnAdd";
            this.BtnAdd.Size = new System.Drawing.Size(85, 30);
            this.BtnAdd.StyleController = this.layoutControl1;
            this.BtnAdd.TabIndex = 5;
            this.BtnAdd.TabStop = false;
            this.BtnAdd.Text = "추가(F1)";
            this.BtnAdd.Click += new System.EventHandler(this.BtnAdd_Click);
            // 
            // BtnClose
            // 
            this.BtnClose.ImageOptions.Image = global::AccAdm.Properties.Resources.cancel_16x16;
            this.BtnClose.Location = new System.Drawing.Point(1141, 44);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(85, 30);
            this.BtnClose.StyleController = this.layoutControl1;
            this.BtnClose.TabIndex = 5;
            this.BtnClose.TabStop = false;
            this.BtnClose.Text = "닫기";
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlGroup1,
            this.layoutControlGroup2});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(1250, 888);
            this.Root.TextVisible = false;
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem2,
            this.emptySpaceItem2,
            this.layoutControlItem3,
            this.layoutControlItem4,
            this.layoutControlItem5});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(1230, 78);
            this.layoutControlGroup1.Text = "검색조건";
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.DateEditFrom;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.MaxSize = new System.Drawing.Size(168, 26);
            this.layoutControlItem1.MinSize = new System.Drawing.Size(168, 26);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(168, 34);
            this.layoutControlItem1.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem1.Text = "업로드일자";
            this.layoutControlItem1.TextSize = new System.Drawing.Size(60, 15);
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.DateEditTo;
            this.layoutControlItem2.CustomizationFormText = "layoutControlItem1";
            this.layoutControlItem2.Location = new System.Drawing.Point(168, 0);
            this.layoutControlItem2.MaxSize = new System.Drawing.Size(116, 26);
            this.layoutControlItem2.MinSize = new System.Drawing.Size(116, 26);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(116, 34);
            this.layoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem2.Text = "~";
            this.layoutControlItem2.TextAlignMode = DevExpress.XtraLayout.TextAlignModeItem.AutoSize;
            this.layoutControlItem2.TextSize = new System.Drawing.Size(8, 15);
            this.layoutControlItem2.TextToControlDistance = 4;
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(284, 0);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(655, 34);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.BtnRetr;
            this.layoutControlItem3.Location = new System.Drawing.Point(939, 0);
            this.layoutControlItem3.MaxSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem3.MinSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(89, 34);
            this.layoutControlItem3.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.BtnAdd;
            this.layoutControlItem4.CustomizationFormText = "layoutControlItem3";
            this.layoutControlItem4.Location = new System.Drawing.Point(1028, 0);
            this.layoutControlItem4.MaxSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem4.MinSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(89, 34);
            this.layoutControlItem4.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem4.Text = "layoutControlItem3";
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.BtnClose;
            this.layoutControlItem5.CustomizationFormText = "layoutControlItem3";
            this.layoutControlItem5.Location = new System.Drawing.Point(1117, 0);
            this.layoutControlItem5.MaxSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem5.MinSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(89, 34);
            this.layoutControlItem5.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem5.Text = "layoutControlItem3";
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // layoutControlGroup2
            // 
            this.layoutControlGroup2.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem6});
            this.layoutControlGroup2.Location = new System.Drawing.Point(0, 78);
            this.layoutControlGroup2.Name = "layoutControlGroup2";
            this.layoutControlGroup2.Size = new System.Drawing.Size(1230, 790);
            this.layoutControlGroup2.Text = "버전리스트";
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.GridRetr;
            this.layoutControlItem6.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(1206, 746);
            this.layoutControlItem6.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem6.TextVisible = false;
            // 
            // SYS001F01
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1250, 888);
            this.Controls.Add(this.layoutControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.Name = "SYS001F01";
            this.Text = "프로그램 버전관리";
            this.Load += new System.EventHandler(this.SYS001F01_Load);
            this.TextChanged += new System.EventHandler(this.SYS001F01_TextChanged);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.SYS001F01_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.GridRetr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewRetr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoDateEditTime.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RepoDateEditTime)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DateEditFrom.Properties.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DateEditFrom.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DateEditTo.Properties.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DateEditTo.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraEditors.DateEdit DateEditFrom;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraEditors.SimpleButton BtnRetr;
        private DevExpress.XtraEditors.DateEdit DateEditTo;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraEditors.SimpleButton BtnAdd;
        private DevExpress.XtraEditors.SimpleButton BtnClose;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraGrid.GridControl GridRetr;
        private DevExpress.XtraGrid.Views.Grid.GridView GridViewRetr;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraGrid.Columns.GridColumn GridColVersionId;
        private DevExpress.XtraGrid.Columns.GridColumn GridColUploadDt;
        private DevExpress.XtraGrid.Columns.GridColumn GridColFile;
        private DevExpress.XtraGrid.Columns.GridColumn GridColFileName;
        private DevExpress.XtraGrid.Columns.GridColumn GridColFileByte;
        private DevExpress.XtraGrid.Columns.GridColumn GridColVersionRmk;
        private DevExpress.XtraGrid.Columns.GridColumn GridColEntId;
        private DevExpress.XtraGrid.Columns.GridColumn GridColEntDt;
        private DevExpress.XtraGrid.Columns.GridColumn GridColModId;
        private DevExpress.XtraGrid.Columns.GridColumn GridColModDt;
        private DevExpress.XtraEditors.Repository.RepositoryItemDateEdit RepoDateEditTime;
    }
}