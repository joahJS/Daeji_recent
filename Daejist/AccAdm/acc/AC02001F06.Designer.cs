namespace AccAdm
{
    partial class AC02001F06
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AC02001F06));
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.BtnRetr = new DevExpress.XtraEditors.SimpleButton();
            this.TxtFindWord = new DevExpress.XtraEditors.TextEdit();
            this.GridRetr = new DevExpress.XtraGrid.GridControl();
            this.GridViewRetr = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.GridColDeptCd = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColDeptNm = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColUpDeptNm = new DevExpress.XtraGrid.Columns.GridColumn();
            this.CboFindSubject = new DevExpress.XtraEditors.ComboBoxEdit();
            this.BtnApply = new DevExpress.XtraEditors.SimpleButton();
            this.BtnClose = new DevExpress.XtraEditors.SimpleButton();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlGroup2 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem3 = new DevExpress.XtraLayout.EmptySpaceItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TxtFindWord.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridRetr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewRetr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboFindSubject.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.BtnRetr);
            this.layoutControl1.Controls.Add(this.TxtFindWord);
            this.layoutControl1.Controls.Add(this.GridRetr);
            this.layoutControl1.Controls.Add(this.CboFindSubject);
            this.layoutControl1.Controls.Add(this.BtnApply);
            this.layoutControl1.Controls.Add(this.BtnClose);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(415, 458);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // BtnRetr
            // 
            this.BtnRetr.ImageOptions.Image = global::AccAdm.Properties.Resources.zoom_16x16;
            this.BtnRetr.Location = new System.Drawing.Point(306, 44);
            this.BtnRetr.Name = "BtnRetr";
            this.BtnRetr.Size = new System.Drawing.Size(85, 30);
            this.BtnRetr.StyleController = this.layoutControl1;
            this.BtnRetr.TabIndex = 2;
            this.BtnRetr.Text = "조회(F5)";
            this.BtnRetr.Click += new System.EventHandler(this.BtnRetr_Click);
            // 
            // TxtFindWord
            // 
            this.TxtFindWord.EnterMoveNextControl = true;
            this.TxtFindWord.Location = new System.Drawing.Point(159, 44);
            this.TxtFindWord.Name = "TxtFindWord";
            this.TxtFindWord.Size = new System.Drawing.Size(120, 22);
            this.TxtFindWord.StyleController = this.layoutControl1;
            this.TxtFindWord.TabIndex = 1;
            // 
            // GridRetr
            // 
            this.GridRetr.Location = new System.Drawing.Point(24, 122);
            this.GridRetr.MainView = this.GridViewRetr;
            this.GridRetr.Name = "GridRetr";
            this.GridRetr.Size = new System.Drawing.Size(367, 278);
            this.GridRetr.TabIndex = 5;
            this.GridRetr.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.GridViewRetr});
            // 
            // GridViewRetr
            // 
            this.GridViewRetr.ColumnPanelRowHeight = 30;
            this.GridViewRetr.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.GridColDeptCd,
            this.GridColDeptNm,
            this.GridColUpDeptNm});
            this.GridViewRetr.GridControl = this.GridRetr;
            this.GridViewRetr.Name = "GridViewRetr";
            this.GridViewRetr.OptionsBehavior.Editable = false;
            this.GridViewRetr.OptionsCustomization.AllowFilter = false;
            this.GridViewRetr.OptionsCustomization.AllowGroup = false;
            this.GridViewRetr.OptionsCustomization.AllowSort = false;
            this.GridViewRetr.OptionsFind.AllowFindPanel = false;
            this.GridViewRetr.OptionsView.ShowGroupPanel = false;
            this.GridViewRetr.RowClick += new DevExpress.XtraGrid.Views.Grid.RowClickEventHandler(this.GridViewRetr_RowClick);
            this.GridViewRetr.KeyDown += new System.Windows.Forms.KeyEventHandler(this.GridViewRetr_KeyDown);
            // 
            // GridColDeptCd
            // 
            this.GridColDeptCd.AppearanceCell.Options.UseTextOptions = true;
            this.GridColDeptCd.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColDeptCd.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColDeptCd.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColDeptCd.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColDeptCd.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColDeptCd.Caption = "부서코드";
            this.GridColDeptCd.FieldName = "DEPT_CD";
            this.GridColDeptCd.Name = "GridColDeptCd";
            this.GridColDeptCd.Visible = true;
            this.GridColDeptCd.VisibleIndex = 0;
            this.GridColDeptCd.Width = 83;
            // 
            // GridColDeptNm
            // 
            this.GridColDeptNm.AppearanceCell.Options.UseTextOptions = true;
            this.GridColDeptNm.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColDeptNm.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColDeptNm.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColDeptNm.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColDeptNm.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColDeptNm.Caption = "부서명";
            this.GridColDeptNm.FieldName = "DEPT_NM";
            this.GridColDeptNm.Name = "GridColDeptNm";
            this.GridColDeptNm.Visible = true;
            this.GridColDeptNm.VisibleIndex = 1;
            this.GridColDeptNm.Width = 129;
            // 
            // GridColUpDeptNm
            // 
            this.GridColUpDeptNm.AppearanceCell.Options.UseTextOptions = true;
            this.GridColUpDeptNm.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColUpDeptNm.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColUpDeptNm.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColUpDeptNm.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColUpDeptNm.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColUpDeptNm.Caption = "상위부서명";
            this.GridColUpDeptNm.FieldName = "UP_DEPT_NM";
            this.GridColUpDeptNm.Name = "GridColUpDeptNm";
            this.GridColUpDeptNm.Visible = true;
            this.GridColUpDeptNm.VisibleIndex = 2;
            this.GridColUpDeptNm.Width = 137;
            // 
            // CboFindSubject
            // 
            this.CboFindSubject.EditValue = "부서명";
            this.CboFindSubject.EnterMoveNextControl = true;
            this.CboFindSubject.Location = new System.Drawing.Point(24, 44);
            this.CboFindSubject.Name = "CboFindSubject";
            this.CboFindSubject.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.CboFindSubject.Properties.Items.AddRange(new object[] {
            "부서코드",
            "부서명"});
            this.CboFindSubject.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
            this.CboFindSubject.Size = new System.Drawing.Size(80, 22);
            this.CboFindSubject.StyleController = this.layoutControl1;
            this.CboFindSubject.TabIndex = 0;
            // 
            // BtnApply
            // 
            this.BtnApply.ImageOptions.Image = global::AccAdm.Properties.Resources.apply_16x16;
            this.BtnApply.Location = new System.Drawing.Point(123, 416);
            this.BtnApply.Name = "BtnApply";
            this.BtnApply.Size = new System.Drawing.Size(85, 30);
            this.BtnApply.StyleController = this.layoutControl1;
            this.BtnApply.TabIndex = 7;
            this.BtnApply.TabStop = false;
            this.BtnApply.Text = "적용(F3)";
            this.BtnApply.Click += new System.EventHandler(this.BtnApply_Click);
            // 
            // BtnClose
            // 
            this.BtnClose.ImageOptions.Image = global::AccAdm.Properties.Resources.cancel_16x16;
            this.BtnClose.Location = new System.Drawing.Point(212, 416);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(85, 30);
            this.BtnClose.StyleController = this.layoutControl1;
            this.BtnClose.TabIndex = 7;
            this.BtnClose.TabStop = false;
            this.BtnClose.Text = "닫기(ESC)";
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlGroup1,
            this.layoutControlGroup2,
            this.layoutControlItem5,
            this.layoutControlItem6,
            this.emptySpaceItem2,
            this.emptySpaceItem3});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(415, 458);
            this.Root.TextVisible = false;
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem3,
            this.layoutControlItem4,
            this.emptySpaceItem1});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(395, 78);
            this.layoutControlGroup1.Text = "검색조건";
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.CboFindSubject;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.MaxSize = new System.Drawing.Size(84, 26);
            this.layoutControlItem1.MinSize = new System.Drawing.Size(84, 26);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(84, 34);
            this.layoutControlItem1.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem1.Text = "찾을단어";
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.TxtFindWord;
            this.layoutControlItem3.Location = new System.Drawing.Point(84, 0);
            this.layoutControlItem3.MaxSize = new System.Drawing.Size(175, 26);
            this.layoutControlItem3.MinSize = new System.Drawing.Size(175, 26);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(175, 34);
            this.layoutControlItem3.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem3.Text = "찾을단어";
            this.layoutControlItem3.TextSize = new System.Drawing.Size(48, 15);
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.BtnRetr;
            this.layoutControlItem4.Location = new System.Drawing.Point(282, 0);
            this.layoutControlItem4.MaxSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem4.MinSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(89, 34);
            this.layoutControlItem4.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(259, 0);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(23, 34);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlGroup2
            // 
            this.layoutControlGroup2.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem2});
            this.layoutControlGroup2.Location = new System.Drawing.Point(0, 78);
            this.layoutControlGroup2.Name = "layoutControlGroup2";
            this.layoutControlGroup2.Size = new System.Drawing.Size(395, 326);
            this.layoutControlGroup2.Text = "부서리스트";
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.GridRetr;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(371, 282);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.BtnApply;
            this.layoutControlItem5.CustomizationFormText = "layoutControlItem4";
            this.layoutControlItem5.Location = new System.Drawing.Point(111, 404);
            this.layoutControlItem5.MaxSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem5.MinSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(89, 34);
            this.layoutControlItem5.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem5.Text = "layoutControlItem4";
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.BtnClose;
            this.layoutControlItem6.CustomizationFormText = "layoutControlItem4";
            this.layoutControlItem6.Location = new System.Drawing.Point(200, 404);
            this.layoutControlItem6.MaxSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem6.MinSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(89, 34);
            this.layoutControlItem6.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem6.Text = "layoutControlItem4";
            this.layoutControlItem6.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem6.TextVisible = false;
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(289, 404);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(106, 34);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem3
            // 
            this.emptySpaceItem3.AllowHotTrack = false;
            this.emptySpaceItem3.Location = new System.Drawing.Point(0, 404);
            this.emptySpaceItem3.Name = "emptySpaceItem3";
            this.emptySpaceItem3.Size = new System.Drawing.Size(111, 34);
            this.emptySpaceItem3.TextSize = new System.Drawing.Size(0, 0);
            // 
            // AC02001F06
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(415, 458);
            this.Controls.Add(this.layoutControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.Name = "AC02001F06";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "부서코드";
            this.Load += new System.EventHandler(this.AC02001F06_Load);
            this.Shown += new System.EventHandler(this.AC02001F06_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.TxtFindWord.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridRetr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewRetr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboFindSubject.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraEditors.ComboBoxEdit CboFindSubject;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraEditors.TextEdit TxtFindWord;
        private DevExpress.XtraGrid.GridControl GridRetr;
        private DevExpress.XtraGrid.Views.Grid.GridView GridViewRetr;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraEditors.SimpleButton BtnRetr;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraEditors.SimpleButton BtnApply;
        private DevExpress.XtraEditors.SimpleButton BtnClose;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
        private DevExpress.XtraGrid.Columns.GridColumn GridColDeptCd;
        private DevExpress.XtraGrid.Columns.GridColumn GridColDeptNm;
        private DevExpress.XtraGrid.Columns.GridColumn GridColUpDeptNm;
    }
}