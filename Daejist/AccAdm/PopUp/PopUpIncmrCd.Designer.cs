namespace AccAdm
{
    partial class PopUpIncmrCd
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
            this.BtnCancel = new DevExpress.XtraEditors.SimpleButton();
            this.BtnConFirm = new DevExpress.XtraEditors.SimpleButton();
            this.GridIncmrCd = new DevExpress.XtraGrid.GridControl();
            this.GridViewIncmrCd = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.GridColIncmrIdtNo = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColIncmrNm = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColCoNm = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColTelNo = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColEmail = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColIncmrCd = new DevExpress.XtraGrid.Columns.GridColumn();
            this.BtnRetr = new DevExpress.XtraEditors.SimpleButton();
            this.TxtIncmrNm = new DevExpress.XtraEditors.TextEdit();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlGroup2 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem3 = new DevExpress.XtraLayout.EmptySpaceItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.GridIncmrCd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewIncmrCd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtIncmrNm.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.BtnCancel);
            this.layoutControl1.Controls.Add(this.BtnConFirm);
            this.layoutControl1.Controls.Add(this.GridIncmrCd);
            this.layoutControl1.Controls.Add(this.BtnRetr);
            this.layoutControl1.Controls.Add(this.TxtIncmrNm);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(569, 524);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // BtnCancel
            // 
            this.BtnCancel.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnCancel.Appearance.Options.UseFont = true;
            this.BtnCancel.Location = new System.Drawing.Point(290, 484);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(66, 28);
            this.BtnCancel.StyleController = this.layoutControl1;
            this.BtnCancel.TabIndex = 9;
            this.BtnCancel.TabStop = false;
            this.BtnCancel.Text = "취소";
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // BtnConFirm
            // 
            this.BtnConFirm.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnConFirm.Appearance.Options.UseFont = true;
            this.BtnConFirm.Location = new System.Drawing.Point(220, 484);
            this.BtnConFirm.Name = "BtnConFirm";
            this.BtnConFirm.Size = new System.Drawing.Size(66, 28);
            this.BtnConFirm.StyleController = this.layoutControl1;
            this.BtnConFirm.TabIndex = 8;
            this.BtnConFirm.TabStop = false;
            this.BtnConFirm.Text = "확인";
            this.BtnConFirm.Click += new System.EventHandler(this.BtnConFirm_Click);
            // 
            // GridIncmrCd
            // 
            this.GridIncmrCd.Location = new System.Drawing.Point(24, 56);
            this.GridIncmrCd.MainView = this.GridViewIncmrCd;
            this.GridIncmrCd.Name = "GridIncmrCd";
            this.GridIncmrCd.Size = new System.Drawing.Size(521, 412);
            this.GridIncmrCd.TabIndex = 1;
            this.GridIncmrCd.TabStop = false;
            this.GridIncmrCd.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.GridViewIncmrCd});
            this.GridIncmrCd.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.GridIncmrCd_PreviewKeyDown);
            // 
            // GridViewIncmrCd
            // 
            this.GridViewIncmrCd.Appearance.HeaderPanel.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridViewIncmrCd.Appearance.HeaderPanel.Options.UseFont = true;
            this.GridViewIncmrCd.Appearance.Row.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridViewIncmrCd.Appearance.Row.Options.UseFont = true;
            this.GridViewIncmrCd.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.GridColIncmrIdtNo,
            this.GridColIncmrNm,
            this.GridColCoNm,
            this.GridColTelNo,
            this.GridColEmail,
            this.GridColIncmrCd});
            this.GridViewIncmrCd.GridControl = this.GridIncmrCd;
            this.GridViewIncmrCd.Name = "GridViewIncmrCd";
            this.GridViewIncmrCd.OptionsBehavior.Editable = false;
            this.GridViewIncmrCd.OptionsBehavior.ReadOnly = true;
            this.GridViewIncmrCd.RowStyle += new DevExpress.XtraGrid.Views.Grid.RowStyleEventHandler(this.GridViewIncmrCd_RowStyle);
            this.GridViewIncmrCd.DoubleClick += new System.EventHandler(this.GridViewIncmrCd_DoubleClick);
            // 
            // GridColIncmrIdtNo
            // 
            this.GridColIncmrIdtNo.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridColIncmrIdtNo.AppearanceCell.Options.UseFont = true;
            this.GridColIncmrIdtNo.Caption = "식별번호";
            this.GridColIncmrIdtNo.FieldName = "INCMR_IDT_NO";
            this.GridColIncmrIdtNo.Name = "GridColIncmrIdtNo";
            this.GridColIncmrIdtNo.Visible = true;
            this.GridColIncmrIdtNo.VisibleIndex = 0;
            this.GridColIncmrIdtNo.Width = 127;
            // 
            // GridColIncmrNm
            // 
            this.GridColIncmrNm.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridColIncmrNm.AppearanceCell.Options.UseFont = true;
            this.GridColIncmrNm.Caption = "소득자명";
            this.GridColIncmrNm.FieldName = "INCMR_NM";
            this.GridColIncmrNm.Name = "GridColIncmrNm";
            this.GridColIncmrNm.Visible = true;
            this.GridColIncmrNm.VisibleIndex = 1;
            this.GridColIncmrNm.Width = 71;
            // 
            // GridColCoNm
            // 
            this.GridColCoNm.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridColCoNm.AppearanceCell.Options.UseFont = true;
            this.GridColCoNm.Caption = "상호";
            this.GridColCoNm.FieldName = "CO_NM";
            this.GridColCoNm.Name = "GridColCoNm";
            this.GridColCoNm.Visible = true;
            this.GridColCoNm.VisibleIndex = 2;
            this.GridColCoNm.Width = 134;
            // 
            // GridColTelNo
            // 
            this.GridColTelNo.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridColTelNo.AppearanceCell.Options.UseFont = true;
            this.GridColTelNo.Caption = "전화번호";
            this.GridColTelNo.FieldName = "TEL_NO";
            this.GridColTelNo.Name = "GridColTelNo";
            this.GridColTelNo.Visible = true;
            this.GridColTelNo.VisibleIndex = 3;
            this.GridColTelNo.Width = 109;
            // 
            // GridColEmail
            // 
            this.GridColEmail.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridColEmail.AppearanceCell.Options.UseFont = true;
            this.GridColEmail.Caption = "E-MALE";
            this.GridColEmail.FieldName = "EMAIL";
            this.GridColEmail.Name = "GridColEmail";
            this.GridColEmail.Visible = true;
            this.GridColEmail.VisibleIndex = 4;
            this.GridColEmail.Width = 62;
            // 
            // GridColIncmrCd
            // 
            this.GridColIncmrCd.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridColIncmrCd.AppearanceCell.Options.UseFont = true;
            this.GridColIncmrCd.Caption = "소득자코드";
            this.GridColIncmrCd.FieldName = "INCMR_CD";
            this.GridColIncmrCd.Name = "GridColIncmrCd";
            // 
            // BtnRetr
            // 
            this.BtnRetr.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnRetr.Appearance.Options.UseFont = true;
            this.BtnRetr.Location = new System.Drawing.Point(487, 12);
            this.BtnRetr.Name = "BtnRetr";
            this.BtnRetr.Size = new System.Drawing.Size(70, 28);
            this.BtnRetr.StyleController = this.layoutControl1;
            this.BtnRetr.TabIndex = 6;
            this.BtnRetr.TabStop = false;
            this.BtnRetr.Text = "조회";
            this.BtnRetr.Click += new System.EventHandler(this.BtnRetr_Click);
            // 
            // TxtIncmrNm
            // 
            this.TxtIncmrNm.Location = new System.Drawing.Point(63, 12);
            this.TxtIncmrNm.Name = "TxtIncmrNm";
            this.TxtIncmrNm.Properties.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtIncmrNm.Properties.Appearance.Options.UseFont = true;
            this.TxtIncmrNm.Properties.AutoHeight = false;
            this.TxtIncmrNm.Size = new System.Drawing.Size(170, 28);
            this.TxtIncmrNm.StyleController = this.layoutControl1;
            this.TxtIncmrNm.TabIndex = 5;
            this.TxtIncmrNm.Leave += new System.EventHandler(this.TxtIncmrNm_Leave);
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.emptySpaceItem1,
            this.layoutControlGroup1,
            this.layoutControlGroup2,
            this.layoutControlItem5,
            this.layoutControlItem6,
            this.emptySpaceItem3});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(569, 524);
            this.Root.TextVisible = false;
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(0, 472);
            this.emptySpaceItem1.MaxSize = new System.Drawing.Size(208, 32);
            this.emptySpaceItem1.MinSize = new System.Drawing.Size(208, 32);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(208, 32);
            this.emptySpaceItem1.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem2,
            this.emptySpaceItem2,
            this.layoutControlItem3});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(549, 32);
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.AppearanceItemCaption.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.layoutControlItem2.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem2.Control = this.TxtIncmrNm;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem2.MaxSize = new System.Drawing.Size(225, 32);
            this.layoutControlItem2.MinSize = new System.Drawing.Size(225, 32);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(225, 32);
            this.layoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem2.Text = "소득자명";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(48, 15);
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(225, 0);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(250, 32);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.BtnRetr;
            this.layoutControlItem3.Location = new System.Drawing.Point(475, 0);
            this.layoutControlItem3.MaxSize = new System.Drawing.Size(74, 32);
            this.layoutControlItem3.MinSize = new System.Drawing.Size(74, 32);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(74, 32);
            this.layoutControlItem3.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // layoutControlGroup2
            // 
            this.layoutControlGroup2.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem4});
            this.layoutControlGroup2.Location = new System.Drawing.Point(0, 32);
            this.layoutControlGroup2.Name = "layoutControlGroup2";
            this.layoutControlGroup2.Size = new System.Drawing.Size(549, 440);
            this.layoutControlGroup2.TextVisible = false;
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.GridIncmrCd;
            this.layoutControlItem4.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(525, 416);
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.BtnConFirm;
            this.layoutControlItem5.Location = new System.Drawing.Point(208, 472);
            this.layoutControlItem5.MaxSize = new System.Drawing.Size(70, 32);
            this.layoutControlItem5.MinSize = new System.Drawing.Size(70, 32);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(70, 32);
            this.layoutControlItem5.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.BtnCancel;
            this.layoutControlItem6.Location = new System.Drawing.Point(278, 472);
            this.layoutControlItem6.MaxSize = new System.Drawing.Size(70, 32);
            this.layoutControlItem6.MinSize = new System.Drawing.Size(70, 32);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(70, 32);
            this.layoutControlItem6.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem6.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem6.TextVisible = false;
            // 
            // emptySpaceItem3
            // 
            this.emptySpaceItem3.AllowHotTrack = false;
            this.emptySpaceItem3.Location = new System.Drawing.Point(348, 472);
            this.emptySpaceItem3.MaxSize = new System.Drawing.Size(201, 32);
            this.emptySpaceItem3.MinSize = new System.Drawing.Size(201, 32);
            this.emptySpaceItem3.Name = "emptySpaceItem3";
            this.emptySpaceItem3.Size = new System.Drawing.Size(201, 32);
            this.emptySpaceItem3.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.emptySpaceItem3.TextSize = new System.Drawing.Size(0, 0);
            // 
            // PopUpIncmrCd
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(569, 524);
            this.Controls.Add(this.layoutControl1);
            this.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.KeyPreview = true;
            this.Name = "PopUpIncmrCd";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "회계계정코드 조회";
            this.Load += new System.EventHandler(this.PopUpIncmrCd_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.PopUpIncmrCd_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.GridIncmrCd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewIncmrCd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtIncmrNm.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraEditors.SimpleButton BtnRetr;
        private DevExpress.XtraEditors.TextEdit TxtIncmrNm;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraEditors.SimpleButton BtnCancel;
        private DevExpress.XtraEditors.SimpleButton BtnConFirm;
        private DevExpress.XtraGrid.GridControl GridIncmrCd;
        private DevExpress.XtraGrid.Views.Grid.GridView GridViewIncmrCd;
        private DevExpress.XtraGrid.Columns.GridColumn GridColIncmrIdtNo;
        private DevExpress.XtraGrid.Columns.GridColumn GridColIncmrNm;
        private DevExpress.XtraGrid.Columns.GridColumn GridColCoNm;
        private DevExpress.XtraGrid.Columns.GridColumn GridColTelNo;
        private DevExpress.XtraGrid.Columns.GridColumn GridColEmail;
        private DevExpress.XtraGrid.Columns.GridColumn GridColIncmrCd;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
    }
}