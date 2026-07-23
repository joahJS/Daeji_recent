namespace AccAdm
{
    partial class PopUpMgmtCd
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
            this.GridMgmtCd = new DevExpress.XtraGrid.GridControl();
            this.GridViewMgmtCd = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.GridColMgmtCd = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColMgmtNm = new DevExpress.XtraGrid.Columns.GridColumn();
            this.BtnRetr = new DevExpress.XtraEditors.SimpleButton();
            this.TxtMgmtNm = new DevExpress.XtraEditors.TextEdit();
            this.TxtMgmtCd = new DevExpress.XtraEditors.TextEdit();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
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
            ((System.ComponentModel.ISupportInitialize)(this.GridMgmtCd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewMgmtCd)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtMgmtNm.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtMgmtCd.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
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
            this.layoutControl1.Controls.Add(this.GridMgmtCd);
            this.layoutControl1.Controls.Add(this.BtnRetr);
            this.layoutControl1.Controls.Add(this.TxtMgmtNm);
            this.layoutControl1.Controls.Add(this.TxtMgmtCd);
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
            // GridMgmtCd
            // 
            this.GridMgmtCd.Location = new System.Drawing.Point(24, 56);
            this.GridMgmtCd.MainView = this.GridViewMgmtCd;
            this.GridMgmtCd.Name = "GridMgmtCd";
            this.GridMgmtCd.Size = new System.Drawing.Size(521, 412);
            this.GridMgmtCd.TabIndex = 1;
            this.GridMgmtCd.TabStop = false;
            this.GridMgmtCd.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.GridViewMgmtCd});
            this.GridMgmtCd.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.GridMgmtCd_PreviewKeyDown);
            // 
            // GridViewMgmtCd
            // 
            this.GridViewMgmtCd.Appearance.HeaderPanel.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridViewMgmtCd.Appearance.HeaderPanel.Options.UseFont = true;
            this.GridViewMgmtCd.Appearance.Row.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridViewMgmtCd.Appearance.Row.Options.UseFont = true;
            this.GridViewMgmtCd.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.GridColMgmtCd,
            this.GridColMgmtNm});
            this.GridViewMgmtCd.GridControl = this.GridMgmtCd;
            this.GridViewMgmtCd.Name = "GridViewMgmtCd";
            this.GridViewMgmtCd.OptionsBehavior.Editable = false;
            this.GridViewMgmtCd.OptionsBehavior.ReadOnly = true;
            this.GridViewMgmtCd.RowStyle += new DevExpress.XtraGrid.Views.Grid.RowStyleEventHandler(this.GridViewMgmtCd_RowStyle);
            this.GridViewMgmtCd.DoubleClick += new System.EventHandler(this.GridViewMgmtCd_DoubleClick);
            // 
            // GridColMgmtCd
            // 
            this.GridColMgmtCd.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridColMgmtCd.AppearanceCell.Options.UseFont = true;
            this.GridColMgmtCd.Caption = "관리항목코드";
            this.GridColMgmtCd.FieldName = "CD";
            this.GridColMgmtCd.Name = "GridColMgmtCd";
            this.GridColMgmtCd.Visible = true;
            this.GridColMgmtCd.VisibleIndex = 0;
            this.GridColMgmtCd.Width = 80;
            // 
            // GridColMgmtNm
            // 
            this.GridColMgmtNm.AppearanceCell.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GridColMgmtNm.AppearanceCell.Options.UseFont = true;
            this.GridColMgmtNm.Caption = "관리항목명";
            this.GridColMgmtNm.FieldName = "NM";
            this.GridColMgmtNm.Name = "GridColMgmtNm";
            this.GridColMgmtNm.Visible = true;
            this.GridColMgmtNm.VisibleIndex = 1;
            this.GridColMgmtNm.Width = 160;
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
            // TxtMgmtNm
            // 
            this.TxtMgmtNm.Location = new System.Drawing.Point(281, 12);
            this.TxtMgmtNm.Name = "TxtMgmtNm";
            this.TxtMgmtNm.Properties.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtMgmtNm.Properties.Appearance.Options.UseFont = true;
            this.TxtMgmtNm.Properties.AutoHeight = false;
            this.TxtMgmtNm.Size = new System.Drawing.Size(137, 28);
            this.TxtMgmtNm.StyleController = this.layoutControl1;
            this.TxtMgmtNm.TabIndex = 5;
            this.TxtMgmtNm.Leave += new System.EventHandler(this.TxtMgmtNm_Leave);
            // 
            // TxtMgmtCd
            // 
            this.TxtMgmtCd.EnterMoveNextControl = true;
            this.TxtMgmtCd.Location = new System.Drawing.Point(89, 12);
            this.TxtMgmtCd.Name = "TxtMgmtCd";
            this.TxtMgmtCd.Properties.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtMgmtCd.Properties.Appearance.Options.UseFont = true;
            this.TxtMgmtCd.Properties.AutoHeight = false;
            this.TxtMgmtCd.Size = new System.Drawing.Size(123, 28);
            this.TxtMgmtCd.StyleController = this.layoutControl1;
            this.TxtMgmtCd.TabIndex = 4;
            this.TxtMgmtCd.Leave += new System.EventHandler(this.TxtMgmtCd_Leave);
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
            this.layoutControlItem1,
            this.layoutControlItem2,
            this.emptySpaceItem2,
            this.layoutControlItem3});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(549, 32);
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.AppearanceItemCaption.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.layoutControlItem1.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem1.Control = this.TxtMgmtCd;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.MaxSize = new System.Drawing.Size(204, 32);
            this.layoutControlItem1.MinSize = new System.Drawing.Size(204, 32);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(204, 32);
            this.layoutControlItem1.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem1.Text = "관리항목코드";
            this.layoutControlItem1.TextAlignMode = DevExpress.XtraLayout.TextAlignModeItem.AutoSize;
            this.layoutControlItem1.TextSize = new System.Drawing.Size(72, 15);
            this.layoutControlItem1.TextToControlDistance = 5;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.AppearanceItemCaption.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.layoutControlItem2.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem2.Control = this.TxtMgmtNm;
            this.layoutControlItem2.Location = new System.Drawing.Point(204, 0);
            this.layoutControlItem2.MaxSize = new System.Drawing.Size(206, 32);
            this.layoutControlItem2.MinSize = new System.Drawing.Size(206, 32);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(206, 32);
            this.layoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem2.Text = "관리항목명";
            this.layoutControlItem2.TextAlignMode = DevExpress.XtraLayout.TextAlignModeItem.AutoSize;
            this.layoutControlItem2.TextSize = new System.Drawing.Size(60, 15);
            this.layoutControlItem2.TextToControlDistance = 5;
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(410, 0);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(65, 32);
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
            this.layoutControlItem4.Control = this.GridMgmtCd;
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
            // PopUpMgmtCd
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(569, 524);
            this.Controls.Add(this.layoutControl1);
            this.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.KeyPreview = true;
            this.Name = "PopUpMgmtCd";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "관리항목코드 조회";
            this.Load += new System.EventHandler(this.PopUpMgmtCd_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.PopUpMgmtCd_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.GridMgmtCd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewMgmtCd)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtMgmtNm.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtMgmtCd.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
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
        private DevExpress.XtraEditors.TextEdit TxtMgmtCd;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraEditors.SimpleButton BtnRetr;
        private DevExpress.XtraEditors.TextEdit TxtMgmtNm;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraEditors.SimpleButton BtnCancel;
        private DevExpress.XtraEditors.SimpleButton BtnConFirm;
        private DevExpress.XtraGrid.GridControl GridMgmtCd;
        private DevExpress.XtraGrid.Views.Grid.GridView GridViewMgmtCd;
        private DevExpress.XtraGrid.Columns.GridColumn GridColMgmtCd;
        private DevExpress.XtraGrid.Columns.GridColumn GridColMgmtNm;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
    }
}