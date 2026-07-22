namespace AccAdm
{
    partial class PopUpGradeCd
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
            this.TxtGrade = new DevExpress.XtraEditors.TextEdit();
            this.GridRetr = new DevExpress.XtraGrid.GridControl();
            this.GridViewRetr = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.GridColGradeCd = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColGradeNm = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColBigCate = new DevExpress.XtraGrid.Columns.GridColumn();
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
            ((System.ComponentModel.ISupportInitialize)(this.TxtGrade.Properties)).BeginInit();
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
            this.layoutControl1.Controls.Add(this.TxtGrade);
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
            this.BtnRetr.Text = "조회";
            this.BtnRetr.Click += new System.EventHandler(this.BtnRetr_Click);
            // 
            // TxtGrade
            // 
            this.TxtGrade.Location = new System.Drawing.Point(51, 12);
            this.TxtGrade.Name = "TxtGrade";
            this.TxtGrade.Properties.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.TxtGrade.Properties.Appearance.Options.UseFont = true;
            this.TxtGrade.Properties.AutoHeight = false;
            this.TxtGrade.Size = new System.Drawing.Size(116, 28);
            this.TxtGrade.StyleController = this.layoutControl1;
            this.TxtGrade.TabIndex = 5;
            this.TxtGrade.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxtGrade_KeyDown);
            this.TxtGrade.Leave += new System.EventHandler(this.TxtGrade_Leave);
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
            this.GridColGradeCd,
            this.GridColGradeNm,
            this.GridColBigCate});
            this.GridViewRetr.GridControl = this.GridRetr;
            this.GridViewRetr.Name = "GridViewRetr";
            this.GridViewRetr.OptionsView.ColumnAutoWidth = false;
            this.GridViewRetr.RowStyle += new DevExpress.XtraGrid.Views.Grid.RowStyleEventHandler(this.GridViewRetr_RowStyle);
            // 
            // GridColGradeCd
            // 
            this.GridColGradeCd.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridColGradeCd.AppearanceCell.Options.UseFont = true;
            this.GridColGradeCd.AppearanceCell.Options.UseTextOptions = true;
            this.GridColGradeCd.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColGradeCd.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColGradeCd.AppearanceHeader.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridColGradeCd.AppearanceHeader.Options.UseFont = true;
            this.GridColGradeCd.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColGradeCd.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColGradeCd.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColGradeCd.Caption = "등급코드";
            this.GridColGradeCd.FieldName = "J_SERIAL";
            this.GridColGradeCd.Name = "GridColGradeCd";
            this.GridColGradeCd.OptionsColumn.AllowEdit = false;
            this.GridColGradeCd.Visible = true;
            this.GridColGradeCd.VisibleIndex = 0;
            this.GridColGradeCd.Width = 161;
            // 
            // GridColGradeNm
            // 
            this.GridColGradeNm.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridColGradeNm.AppearanceCell.Options.UseFont = true;
            this.GridColGradeNm.AppearanceCell.Options.UseTextOptions = true;
            this.GridColGradeNm.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColGradeNm.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColGradeNm.AppearanceHeader.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridColGradeNm.AppearanceHeader.Options.UseFont = true;
            this.GridColGradeNm.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColGradeNm.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColGradeNm.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColGradeNm.Caption = "등급이름";
            this.GridColGradeNm.FieldName = "GUBUN1";
            this.GridColGradeNm.Name = "GridColGradeNm";
            this.GridColGradeNm.OptionsColumn.AllowEdit = false;
            this.GridColGradeNm.Visible = true;
            this.GridColGradeNm.VisibleIndex = 1;
            this.GridColGradeNm.Width = 265;
            // 
            // GridColBigCate
            // 
            this.GridColBigCate.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridColBigCate.AppearanceCell.Options.UseFont = true;
            this.GridColBigCate.AppearanceCell.Options.UseTextOptions = true;
            this.GridColBigCate.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColBigCate.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColBigCate.AppearanceHeader.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.GridColBigCate.AppearanceHeader.Options.UseFont = true;
            this.GridColBigCate.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColBigCate.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColBigCate.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColBigCate.Caption = "대구분";
            this.GridColBigCate.FieldName = "DAEGUBUN";
            this.GridColBigCate.Name = "GridColBigCate";
            this.GridColBigCate.OptionsColumn.AllowEdit = false;
            this.GridColBigCate.Visible = true;
            this.GridColBigCate.VisibleIndex = 2;
            this.GridColBigCate.Width = 172;
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
            this.BtnConfirm.Location = new System.Drawing.Point(270, 542);
            this.BtnConfirm.Name = "BtnConfirm";
            this.BtnConfirm.Size = new System.Drawing.Size(66, 28);
            this.BtnConfirm.StyleController = this.layoutControl1;
            this.BtnConfirm.TabIndex = 6;
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
            this.BtnClose.Location = new System.Drawing.Point(340, 542);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(66, 28);
            this.BtnClose.StyleController = this.layoutControl1;
            this.BtnClose.TabIndex = 6;
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
            this.layoutControlItem2.Control = this.TxtGrade;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem2.MinSize = new System.Drawing.Size(1, 1);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.OptionsPrint.AppearanceItemCaption.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.layoutControlItem2.OptionsPrint.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem2.Size = new System.Drawing.Size(159, 32);
            this.layoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem2.Text = "등급명";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(36, 15);
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(159, 0);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(389, 32);
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
            this.layoutControlItem4.Location = new System.Drawing.Point(258, 530);
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
            this.layoutControlItem5.Location = new System.Drawing.Point(328, 530);
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
            this.emptySpaceItem2.Size = new System.Drawing.Size(258, 32);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem3
            // 
            this.emptySpaceItem3.AllowHotTrack = false;
            this.emptySpaceItem3.Location = new System.Drawing.Point(398, 530);
            this.emptySpaceItem3.Name = "emptySpaceItem3";
            this.emptySpaceItem3.Size = new System.Drawing.Size(220, 32);
            this.emptySpaceItem3.TextSize = new System.Drawing.Size(0, 0);
            // 
            // PopUpGradeCd
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(638, 582);
            this.Controls.Add(this.layoutControl1);
            this.Font = new System.Drawing.Font("맑은 고딕", 9F);
            this.KeyPreview = true;
            this.Name = "PopUpGradeCd";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "등급코드";
            this.Load += new System.EventHandler(this.PopUpGradeCd_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.PopUpGradeCd_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.TxtGrade.Properties)).EndInit();
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
        private DevExpress.XtraEditors.SimpleButton BtnRetr;
        private DevExpress.XtraEditors.TextEdit TxtGrade;
        private DevExpress.XtraGrid.GridControl GridRetr;
        private DevExpress.XtraGrid.Views.Grid.GridView GridViewRetr;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraEditors.SimpleButton BtnConfirm;
        private DevExpress.XtraEditors.SimpleButton BtnClose;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
        private DevExpress.XtraGrid.Columns.GridColumn GridColGradeCd;
        private DevExpress.XtraGrid.Columns.GridColumn GridColGradeNm;
        private DevExpress.XtraGrid.Columns.GridColumn GridColBigCate;
    }
}