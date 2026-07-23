namespace AccAdm
{
    partial class PopUpEmployeeCd
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
            this.BtnRetr = new DevExpress.XtraEditors.SimpleButton();
            this.TxtEmpNm = new DevExpress.XtraEditors.TextEdit();
            this.GridRetr = new DevExpress.XtraGrid.GridControl();
            this.GridViewRetr = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.GridColEmpId = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridEmpNm = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColDeptCd = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColEntranceYmd = new DevExpress.XtraGrid.Columns.GridColumn();
            this.BtnConfirm = new DevExpress.XtraEditors.SimpleButton();
            this.BtnClose = new DevExpress.XtraEditors.SimpleButton();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem3 = new DevExpress.XtraLayout.EmptySpaceItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TxtEmpNm.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridRetr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewRetr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.BtnRetr);
            this.layoutControl1.Controls.Add(this.TxtEmpNm);
            this.layoutControl1.Controls.Add(this.GridRetr);
            this.layoutControl1.Controls.Add(this.BtnConfirm);
            this.layoutControl1.Controls.Add(this.BtnClose);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(638, 582);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // BtnRetr
            // 
            this.BtnRetr.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.BtnRetr.Appearance.Options.UseFont = true;
            this.BtnRetr.AppearanceDisabled.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.BtnRetr.AppearanceDisabled.Options.UseFont = true;
            this.BtnRetr.AppearanceHovered.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.BtnRetr.AppearanceHovered.Options.UseFont = true;
            this.BtnRetr.AppearancePressed.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.BtnRetr.AppearancePressed.Options.UseFont = true;
            this.BtnRetr.Location = new System.Drawing.Point(560, 12);
            this.BtnRetr.Name = "BtnRetr";
            this.BtnRetr.Size = new System.Drawing.Size(66, 28);
            this.BtnRetr.StyleController = this.layoutControl1;
            this.BtnRetr.TabIndex = 6;
            this.BtnRetr.TabStop = false;
            this.BtnRetr.Text = "조회";
            this.BtnRetr.Click += new System.EventHandler(this.BtnRetr_Click);
            // 
            // TxtEmpNm
            // 
            this.TxtEmpNm.Location = new System.Drawing.Point(51, 12);
            this.TxtEmpNm.Name = "TxtEmpNm";
            this.TxtEmpNm.Properties.AutoHeight = false;
            this.TxtEmpNm.Size = new System.Drawing.Size(136, 28);
            this.TxtEmpNm.StyleController = this.layoutControl1;
            this.TxtEmpNm.TabIndex = 5;
            this.TxtEmpNm.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxtEmpNm_KeyDown);
            this.TxtEmpNm.Leave += new System.EventHandler(this.TxtEmpNm_Leave);
            // 
            // GridRetr
            // 
            this.GridRetr.Location = new System.Drawing.Point(12, 44);
            this.GridRetr.MainView = this.GridViewRetr;
            this.GridRetr.Name = "GridRetr";
            this.GridRetr.Size = new System.Drawing.Size(614, 494);
            this.GridRetr.TabIndex = 4;
            this.GridRetr.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.GridViewRetr});
            this.GridRetr.DoubleClick += new System.EventHandler(this.GridRetr_DoubleClick);
            this.GridRetr.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.GridRetr_PreviewKeyDown);
            // 
            // GridViewRetr
            // 
            this.GridViewRetr.ColumnPanelRowHeight = 30;
            this.GridViewRetr.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.GridColEmpId,
            this.GridEmpNm,
            this.GridColDeptCd,
            this.GridColEntranceYmd});
            this.GridViewRetr.GridControl = this.GridRetr;
            this.GridViewRetr.Name = "GridViewRetr";
            this.GridViewRetr.OptionsView.ColumnAutoWidth = false;
            this.GridViewRetr.RowStyle += new DevExpress.XtraGrid.Views.Grid.RowStyleEventHandler(this.GridViewRetr_RowStyle);
            // 
            // GridColEmpId
            // 
            this.GridColEmpId.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridColEmpId.AppearanceCell.Options.UseFont = true;
            this.GridColEmpId.AppearanceCell.Options.UseTextOptions = true;
            this.GridColEmpId.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColEmpId.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColEmpId.AppearanceHeader.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridColEmpId.AppearanceHeader.Options.UseFont = true;
            this.GridColEmpId.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColEmpId.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColEmpId.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColEmpId.Caption = "직원ID";
            this.GridColEmpId.FieldName = "EMP_ID";
            this.GridColEmpId.Name = "GridColEmpId";
            this.GridColEmpId.OptionsColumn.AllowEdit = false;
            this.GridColEmpId.Visible = true;
            this.GridColEmpId.VisibleIndex = 0;
            this.GridColEmpId.Width = 111;
            // 
            // GridEmpNm
            // 
            this.GridEmpNm.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridEmpNm.AppearanceCell.Options.UseFont = true;
            this.GridEmpNm.AppearanceCell.Options.UseTextOptions = true;
            this.GridEmpNm.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridEmpNm.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridEmpNm.AppearanceHeader.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridEmpNm.AppearanceHeader.Options.UseFont = true;
            this.GridEmpNm.AppearanceHeader.Options.UseTextOptions = true;
            this.GridEmpNm.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridEmpNm.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridEmpNm.Caption = "직원명";
            this.GridEmpNm.FieldName = "EMP_NM";
            this.GridEmpNm.Name = "GridEmpNm";
            this.GridEmpNm.OptionsColumn.AllowEdit = false;
            this.GridEmpNm.Visible = true;
            this.GridEmpNm.VisibleIndex = 1;
            this.GridEmpNm.Width = 188;
            // 
            // GridColDeptCd
            // 
            this.GridColDeptCd.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridColDeptCd.AppearanceCell.Options.UseFont = true;
            this.GridColDeptCd.AppearanceCell.Options.UseTextOptions = true;
            this.GridColDeptCd.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColDeptCd.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColDeptCd.AppearanceHeader.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridColDeptCd.AppearanceHeader.Options.UseFont = true;
            this.GridColDeptCd.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColDeptCd.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColDeptCd.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColDeptCd.Caption = "부서명";
            this.GridColDeptCd.FieldName = "DEPT_CD";
            this.GridColDeptCd.Name = "GridColDeptCd";
            this.GridColDeptCd.OptionsColumn.AllowEdit = false;
            this.GridColDeptCd.Visible = true;
            this.GridColDeptCd.VisibleIndex = 2;
            this.GridColDeptCd.Width = 132;
            // 
            // GridColEntranceYmd
            // 
            this.GridColEntranceYmd.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridColEntranceYmd.AppearanceCell.Options.UseFont = true;
            this.GridColEntranceYmd.AppearanceCell.Options.UseTextOptions = true;
            this.GridColEntranceYmd.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColEntranceYmd.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColEntranceYmd.AppearanceHeader.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridColEntranceYmd.AppearanceHeader.Options.UseFont = true;
            this.GridColEntranceYmd.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColEntranceYmd.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColEntranceYmd.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColEntranceYmd.Caption = "입사일자";
            this.GridColEntranceYmd.FieldName = "ENTRANCE_YMD";
            this.GridColEntranceYmd.Name = "GridColEntranceYmd";
            this.GridColEntranceYmd.OptionsColumn.AllowEdit = false;
            this.GridColEntranceYmd.Visible = true;
            this.GridColEntranceYmd.VisibleIndex = 3;
            this.GridColEntranceYmd.Width = 164;
            // 
            // BtnConfirm
            // 
            this.BtnConfirm.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.BtnConfirm.Appearance.Options.UseFont = true;
            this.BtnConfirm.AppearanceDisabled.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.BtnConfirm.AppearanceDisabled.Options.UseFont = true;
            this.BtnConfirm.AppearanceHovered.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.BtnConfirm.AppearanceHovered.Options.UseFont = true;
            this.BtnConfirm.AppearancePressed.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.BtnConfirm.AppearancePressed.Options.UseFont = true;
            this.BtnConfirm.Location = new System.Drawing.Point(274, 542);
            this.BtnConfirm.Name = "BtnConfirm";
            this.BtnConfirm.Size = new System.Drawing.Size(66, 28);
            this.BtnConfirm.StyleController = this.layoutControl1;
            this.BtnConfirm.TabIndex = 6;
            this.BtnConfirm.TabStop = false;
            this.BtnConfirm.Text = "확인";
            this.BtnConfirm.Click += new System.EventHandler(this.BtnConfirm_Click);
            // 
            // BtnClose
            // 
            this.BtnClose.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.BtnClose.Appearance.Options.UseFont = true;
            this.BtnClose.AppearanceDisabled.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.BtnClose.AppearanceDisabled.Options.UseFont = true;
            this.BtnClose.AppearanceHovered.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.BtnClose.AppearanceHovered.Options.UseFont = true;
            this.BtnClose.AppearancePressed.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.BtnClose.AppearancePressed.Options.UseFont = true;
            this.BtnClose.Location = new System.Drawing.Point(344, 542);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(66, 28);
            this.BtnClose.StyleController = this.layoutControl1;
            this.BtnClose.TabIndex = 6;
            this.BtnClose.TabStop = false;
            this.BtnClose.Text = "취소";
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem2,
            this.emptySpaceItem1,
            this.layoutControlItem3,
            this.layoutControlItem4,
            this.layoutControlItem5,
            this.emptySpaceItem2,
            this.emptySpaceItem3});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(638, 582);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.GridRetr;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 32);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(618, 498);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.AppearanceItemCaption.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.layoutControlItem2.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem2.Control = this.TxtEmpNm;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem2.MinSize = new System.Drawing.Size(1, 1);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.OptionsPrint.AppearanceItemCaption.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.layoutControlItem2.OptionsPrint.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem2.Size = new System.Drawing.Size(179, 32);
            this.layoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem2.Text = "직원명";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(36, 15);
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(179, 0);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(369, 32);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.BtnRetr;
            this.layoutControlItem3.Location = new System.Drawing.Point(548, 0);
            this.layoutControlItem3.MaxSize = new System.Drawing.Size(70, 32);
            this.layoutControlItem3.MinSize = new System.Drawing.Size(70, 32);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(70, 32);
            this.layoutControlItem3.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.BtnConfirm;
            this.layoutControlItem4.CustomizationFormText = "layoutControlItem3";
            this.layoutControlItem4.Location = new System.Drawing.Point(262, 530);
            this.layoutControlItem4.MaxSize = new System.Drawing.Size(70, 32);
            this.layoutControlItem4.MinSize = new System.Drawing.Size(70, 32);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(70, 32);
            this.layoutControlItem4.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem4.Text = "layoutControlItem3";
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.BtnClose;
            this.layoutControlItem5.CustomizationFormText = "layoutControlItem3";
            this.layoutControlItem5.Location = new System.Drawing.Point(332, 530);
            this.layoutControlItem5.MaxSize = new System.Drawing.Size(70, 32);
            this.layoutControlItem5.MinSize = new System.Drawing.Size(70, 32);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(70, 32);
            this.layoutControlItem5.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem5.Text = "layoutControlItem3";
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(0, 530);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(262, 32);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem3
            // 
            this.emptySpaceItem3.AllowHotTrack = false;
            this.emptySpaceItem3.Location = new System.Drawing.Point(402, 530);
            this.emptySpaceItem3.Name = "emptySpaceItem3";
            this.emptySpaceItem3.Size = new System.Drawing.Size(216, 32);
            this.emptySpaceItem3.TextSize = new System.Drawing.Size(0, 0);
            // 
            // PopUpEmployeeCd
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(638, 582);
            this.Controls.Add(this.layoutControl1);
            this.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.KeyPreview = true;
            this.Name = "PopUpEmployeeCd";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "사원코드";
            this.Load += new System.EventHandler(this.PopUpEmployeeCd_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.PopUpEmployeeCd_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.TxtEmpNm.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridRetr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewRetr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraEditors.TextEdit TxtEmpNm;
        private DevExpress.XtraGrid.GridControl GridRetr;
        private DevExpress.XtraGrid.Views.Grid.GridView GridViewRetr;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraEditors.SimpleButton BtnRetr;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraEditors.SimpleButton BtnConfirm;
        private DevExpress.XtraEditors.SimpleButton BtnClose;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
        private DevExpress.XtraGrid.Columns.GridColumn GridColEmpId;
        private DevExpress.XtraGrid.Columns.GridColumn GridEmpNm;
        private DevExpress.XtraGrid.Columns.GridColumn GridColDeptCd;
        private DevExpress.XtraGrid.Columns.GridColumn GridColEntranceYmd;
    }
}