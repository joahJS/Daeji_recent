namespace AccAdm
{
    partial class MesGradeCdDev
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
            this.BtnAdd = new DevExpress.XtraEditors.SimpleButton();
            this.BtnSave = new DevExpress.XtraEditors.SimpleButton();
            this.BtnClose = new DevExpress.XtraEditors.SimpleButton();
            this.BtnRetr = new DevExpress.XtraEditors.SimpleButton();
            this.LkupEditDetGb = new DevExpress.XtraEditors.LookUpEdit();
            this.LkupEditGb = new DevExpress.XtraEditors.LookUpEdit();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.GridRetr = new DevExpress.XtraGrid.GridControl();
            this.GridViewRetr = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.GridColGbCd = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColDetCd = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColItemCd = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColItemNm = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColRmk = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColUseYn = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColSeq = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemLookUpEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
            this.repositoryItemLookUpEdit2 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.LkupEditDetGb.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LkupEditGb.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridRetr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewRetr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.BtnAdd);
            this.layoutControl1.Controls.Add(this.BtnSave);
            this.layoutControl1.Controls.Add(this.BtnClose);
            this.layoutControl1.Controls.Add(this.BtnRetr);
            this.layoutControl1.Controls.Add(this.LkupEditDetGb);
            this.layoutControl1.Controls.Add(this.LkupEditGb);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(823, 52);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // BtnAdd
            // 
            this.BtnAdd.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnAdd.Appearance.Options.UseFont = true;
            this.BtnAdd.Location = new System.Drawing.Point(604, 12);
            this.BtnAdd.Name = "BtnAdd";
            this.BtnAdd.Size = new System.Drawing.Size(66, 28);
            this.BtnAdd.StyleController = this.layoutControl1;
            this.BtnAdd.TabIndex = 9;
            this.BtnAdd.Text = "추가";
            this.BtnAdd.Click += new System.EventHandler(this.BtnAdd_Click);
            // 
            // BtnSave
            // 
            this.BtnSave.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnSave.Appearance.Options.UseFont = true;
            this.BtnSave.Location = new System.Drawing.Point(674, 12);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(66, 28);
            this.BtnSave.StyleController = this.layoutControl1;
            this.BtnSave.TabIndex = 8;
            this.BtnSave.Text = "저장";
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // BtnClose
            // 
            this.BtnClose.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnClose.Appearance.Options.UseFont = true;
            this.BtnClose.Location = new System.Drawing.Point(744, 12);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(67, 28);
            this.BtnClose.StyleController = this.layoutControl1;
            this.BtnClose.TabIndex = 7;
            this.BtnClose.Text = "닫기";
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // BtnRetr
            // 
            this.BtnRetr.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnRetr.Appearance.Options.UseFont = true;
            this.BtnRetr.Location = new System.Drawing.Point(534, 12);
            this.BtnRetr.Name = "BtnRetr";
            this.BtnRetr.Size = new System.Drawing.Size(66, 28);
            this.BtnRetr.StyleController = this.layoutControl1;
            this.BtnRetr.TabIndex = 6;
            this.BtnRetr.Text = "조회";
            this.BtnRetr.Click += new System.EventHandler(this.BtnRetr_Click);
            // 
            // LkupEditDetGb
            // 
            this.LkupEditDetGb.Location = new System.Drawing.Point(180, 12);
            this.LkupEditDetGb.Name = "LkupEditDetGb";
            this.LkupEditDetGb.Properties.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LkupEditDetGb.Properties.Appearance.Options.UseFont = true;
            this.LkupEditDetGb.Properties.AppearanceDropDown.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LkupEditDetGb.Properties.AppearanceDropDown.Options.UseFont = true;
            this.LkupEditDetGb.Properties.AutoHeight = false;
            this.LkupEditDetGb.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.LkupEditDetGb.Properties.ShowHeader = false;
            this.LkupEditDetGb.Size = new System.Drawing.Size(105, 28);
            this.LkupEditDetGb.StyleController = this.layoutControl1;
            this.LkupEditDetGb.TabIndex = 5;
            this.LkupEditDetGb.EditValueChanged += new System.EventHandler(this.LkupEditDetGb_EditValueChanged);
            // 
            // LkupEditGb
            // 
            this.LkupEditGb.ImeMode = System.Windows.Forms.ImeMode.Hangul;
            this.LkupEditGb.Location = new System.Drawing.Point(53, 12);
            this.LkupEditGb.Name = "LkupEditGb";
            this.LkupEditGb.Properties.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LkupEditGb.Properties.Appearance.Options.UseFont = true;
            this.LkupEditGb.Properties.Appearance.Options.UseTextOptions = true;
            this.LkupEditGb.Properties.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.LkupEditGb.Properties.AppearanceDropDown.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LkupEditGb.Properties.AppearanceDropDown.Options.UseFont = true;
            this.LkupEditGb.Properties.AutoHeight = false;
            this.LkupEditGb.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.LkupEditGb.Properties.NullText = "";
            this.LkupEditGb.Properties.ShowHeader = false;
            this.LkupEditGb.Size = new System.Drawing.Size(82, 28);
            this.LkupEditGb.StyleController = this.layoutControl1;
            this.LkupEditGb.TabIndex = 4;
            this.LkupEditGb.EditValueChanged += new System.EventHandler(this.LkupEditGb_EditValueChanged);
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.emptySpaceItem1,
            this.layoutControlItem1,
            this.layoutControlItem2,
            this.layoutControlItem3,
            this.layoutControlItem4,
            this.layoutControlItem5,
            this.layoutControlItem6});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(823, 52);
            this.Root.TextVisible = false;
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(277, 0);
            this.emptySpaceItem1.MinSize = new System.Drawing.Size(10, 11);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(245, 32);
            this.emptySpaceItem1.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.AppearanceItemCaption.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.layoutControlItem1.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem1.Control = this.LkupEditGb;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.MinSize = new System.Drawing.Size(1, 1);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(127, 32);
            this.layoutControlItem1.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem1.Text = "대분류";
            this.layoutControlItem1.TextAlignMode = DevExpress.XtraLayout.TextAlignModeItem.AutoSize;
            this.layoutControlItem1.TextSize = new System.Drawing.Size(36, 15);
            this.layoutControlItem1.TextToControlDistance = 5;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.AppearanceItemCaption.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.layoutControlItem2.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem2.Control = this.LkupEditDetGb;
            this.layoutControlItem2.Location = new System.Drawing.Point(127, 0);
            this.layoutControlItem2.MinSize = new System.Drawing.Size(1, 1);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(150, 32);
            this.layoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem2.Text = "중분류";
            this.layoutControlItem2.TextAlignMode = DevExpress.XtraLayout.TextAlignModeItem.AutoSize;
            this.layoutControlItem2.TextSize = new System.Drawing.Size(36, 15);
            this.layoutControlItem2.TextToControlDistance = 5;
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.BtnRetr;
            this.layoutControlItem3.Location = new System.Drawing.Point(522, 0);
            this.layoutControlItem3.MinSize = new System.Drawing.Size(1, 1);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(70, 32);
            this.layoutControlItem3.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.BtnClose;
            this.layoutControlItem4.Location = new System.Drawing.Point(732, 0);
            this.layoutControlItem4.MinSize = new System.Drawing.Size(1, 1);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(71, 32);
            this.layoutControlItem4.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.BtnSave;
            this.layoutControlItem5.Location = new System.Drawing.Point(662, 0);
            this.layoutControlItem5.MinSize = new System.Drawing.Size(1, 1);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(70, 32);
            this.layoutControlItem5.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem5.Text = "저장";
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.BtnAdd;
            this.layoutControlItem6.Location = new System.Drawing.Point(592, 0);
            this.layoutControlItem6.MinSize = new System.Drawing.Size(1, 1);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(70, 32);
            this.layoutControlItem6.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem6.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem6.TextVisible = false;
            // 
            // GridRetr
            // 
            this.GridRetr.Dock = System.Windows.Forms.DockStyle.Fill;
            this.GridRetr.Location = new System.Drawing.Point(0, 52);
            this.GridRetr.MainView = this.GridViewRetr;
            this.GridRetr.Name = "GridRetr";
            this.GridRetr.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemLookUpEdit1,
            this.repositoryItemLookUpEdit2});
            this.GridRetr.Size = new System.Drawing.Size(823, 596);
            this.GridRetr.TabIndex = 1;
            this.GridRetr.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.GridViewRetr});
            this.GridRetr.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.GridRetr_PreviewKeyDown);
            // 
            // GridViewRetr
            // 
            this.GridViewRetr.Appearance.HeaderPanel.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridViewRetr.Appearance.HeaderPanel.Options.UseFont = true;
            this.GridViewRetr.Appearance.Row.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridViewRetr.Appearance.Row.Options.UseFont = true;
            this.GridViewRetr.ColumnPanelRowHeight = 30;
            this.GridViewRetr.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.GridColGbCd,
            this.GridColDetCd,
            this.GridColItemCd,
            this.GridColItemNm,
            this.GridColRmk,
            this.GridColUseYn,
            this.GridColSeq});
            this.GridViewRetr.GridControl = this.GridRetr;
            this.GridViewRetr.Name = "GridViewRetr";
            this.GridViewRetr.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.True;
            this.GridViewRetr.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.True;
            this.GridViewRetr.OptionsBehavior.EditingMode = DevExpress.XtraGrid.Views.Grid.GridEditingMode.EditFormInplace;
            this.GridViewRetr.OptionsBehavior.EditorShowMode = DevExpress.Utils.EditorShowMode.Click;
            this.GridViewRetr.RowStyle += new DevExpress.XtraGrid.Views.Grid.RowStyleEventHandler(this.GridViewRetr_RowStyle);
            // 
            // GridColGbCd
            // 
            this.GridColGbCd.AppearanceCell.Options.UseTextOptions = true;
            this.GridColGbCd.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColGbCd.Caption = "대분류";
            this.GridColGbCd.FieldName = "GB_CD";
            this.GridColGbCd.Name = "GridColGbCd";
            this.GridColGbCd.Visible = true;
            this.GridColGbCd.VisibleIndex = 0;
            this.GridColGbCd.Width = 87;
            // 
            // GridColDetCd
            // 
            this.GridColDetCd.AppearanceCell.Options.UseTextOptions = true;
            this.GridColDetCd.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColDetCd.Caption = "중분류";
            this.GridColDetCd.FieldName = "DET_CD";
            this.GridColDetCd.Name = "GridColDetCd";
            this.GridColDetCd.Visible = true;
            this.GridColDetCd.VisibleIndex = 1;
            this.GridColDetCd.Width = 78;
            // 
            // GridColItemCd
            // 
            this.GridColItemCd.AppearanceCell.Options.UseTextOptions = true;
            this.GridColItemCd.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColItemCd.Caption = "항목코드";
            this.GridColItemCd.FieldName = "ITEM_CD";
            this.GridColItemCd.Name = "GridColItemCd";
            this.GridColItemCd.OptionsColumn.AllowEdit = false;
            this.GridColItemCd.OptionsColumn.AllowFocus = false;
            this.GridColItemCd.Visible = true;
            this.GridColItemCd.VisibleIndex = 2;
            this.GridColItemCd.Width = 64;
            // 
            // GridColItemNm
            // 
            this.GridColItemNm.AppearanceCell.Options.UseTextOptions = true;
            this.GridColItemNm.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColItemNm.Caption = "항목명";
            this.GridColItemNm.FieldName = "ITEM_NM";
            this.GridColItemNm.Name = "GridColItemNm";
            this.GridColItemNm.Visible = true;
            this.GridColItemNm.VisibleIndex = 3;
            this.GridColItemNm.Width = 79;
            // 
            // GridColRmk
            // 
            this.GridColRmk.AppearanceCell.Options.UseTextOptions = true;
            this.GridColRmk.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColRmk.Caption = "상세내용";
            this.GridColRmk.FieldName = "RMK";
            this.GridColRmk.Name = "GridColRmk";
            this.GridColRmk.Visible = true;
            this.GridColRmk.VisibleIndex = 4;
            this.GridColRmk.Width = 495;
            // 
            // GridColUseYn
            // 
            this.GridColUseYn.AppearanceCell.Options.UseTextOptions = true;
            this.GridColUseYn.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColUseYn.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColUseYn.Caption = "사용여부";
            this.GridColUseYn.FieldName = "USE_YN";
            this.GridColUseYn.Name = "GridColUseYn";
            this.GridColUseYn.Visible = true;
            this.GridColUseYn.VisibleIndex = 5;
            // 
            // GridColSeq
            // 
            this.GridColSeq.AppearanceCell.Options.UseTextOptions = true;
            this.GridColSeq.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColSeq.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColSeq.Caption = "순서";
            this.GridColSeq.FieldName = "SORT_SEQ";
            this.GridColSeq.Name = "GridColSeq";
            this.GridColSeq.Visible = true;
            this.GridColSeq.VisibleIndex = 6;
            // 
            // repositoryItemLookUpEdit1
            // 
            this.repositoryItemLookUpEdit1.AutoHeight = false;
            this.repositoryItemLookUpEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemLookUpEdit1.Name = "repositoryItemLookUpEdit1";
            // 
            // repositoryItemLookUpEdit2
            // 
            this.repositoryItemLookUpEdit2.AutoHeight = false;
            this.repositoryItemLookUpEdit2.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemLookUpEdit2.Name = "repositoryItemLookUpEdit2";
            // 
            // MesGradeCdDev
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(823, 648);
            this.Controls.Add(this.GridRetr);
            this.Controls.Add(this.layoutControl1);
            this.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.KeyPreview = true;
            this.Name = "MesGradeCdDev";
            this.Text = "항목코드분류(등급분류)";
            this.Load += new System.EventHandler(this.MesGradeCdDev_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.MesGradeCdDev_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.LkupEditDetGb.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LkupEditGb.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridRetr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewRetr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraEditors.LookUpEdit LkupEditGb;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraEditors.SimpleButton BtnSave;
        private DevExpress.XtraEditors.SimpleButton BtnClose;
        private DevExpress.XtraEditors.SimpleButton BtnRetr;
        private DevExpress.XtraEditors.LookUpEdit LkupEditDetGb;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraGrid.GridControl GridRetr;
        private DevExpress.XtraGrid.Views.Grid.GridView GridViewRetr;
        private DevExpress.XtraGrid.Columns.GridColumn GridColGbCd;
        private DevExpress.XtraGrid.Columns.GridColumn GridColDetCd;
        private DevExpress.XtraGrid.Columns.GridColumn GridColItemCd;
        private DevExpress.XtraGrid.Columns.GridColumn GridColItemNm;
        private DevExpress.XtraGrid.Columns.GridColumn GridColRmk;
        private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit1;
        private DevExpress.XtraEditors.SimpleButton BtnAdd;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit2;
        private DevExpress.XtraGrid.Columns.GridColumn GridColUseYn;
        private DevExpress.XtraGrid.Columns.GridColumn GridColSeq;
    }
}