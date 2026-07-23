namespace AccAdm
{
    partial class PP002F00
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
            this.GridRetr = new DevExpress.XtraGrid.GridControl();
            this.GridViewRetr = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.GridColUsrId = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColUsrNm = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColDeptNm = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColJkwiNm = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColMoblNo = new DevExpress.XtraGrid.Columns.GridColumn();
            this.GridColRk = new DevExpress.XtraGrid.Columns.GridColumn();
            this.TxtFindWord = new DevExpress.XtraEditors.TextEdit();
            this.CboFindIdx = new DevExpress.XtraEditors.ComboBoxEdit();
            this.BtnApply = new DevExpress.XtraEditors.SimpleButton();
            this.BtnClose = new DevExpress.XtraEditors.SimpleButton();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlGroup2 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.GridRetr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewRetr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtFindWord.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboFindIdx.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.BtnRetr);
            this.layoutControl1.Controls.Add(this.GridRetr);
            this.layoutControl1.Controls.Add(this.TxtFindWord);
            this.layoutControl1.Controls.Add(this.CboFindIdx);
            this.layoutControl1.Controls.Add(this.BtnApply);
            this.layoutControl1.Controls.Add(this.BtnClose);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(1059, 532);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // BtnRetr
            // 
            this.BtnRetr.ImageOptions.Image = global::AccAdm.Properties.Resources.zoom_16x16;
            this.BtnRetr.Location = new System.Drawing.Point(950, 44);
            this.BtnRetr.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BtnRetr.Name = "BtnRetr";
            this.BtnRetr.Size = new System.Drawing.Size(85, 30);
            this.BtnRetr.StyleController = this.layoutControl1;
            this.BtnRetr.TabIndex = 2;
            this.BtnRetr.Text = "조회(F5)";
            this.BtnRetr.Click += new System.EventHandler(this.BtnRetr_Click);
            // 
            // GridRetr
            // 
            this.GridRetr.EmbeddedNavigator.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.GridRetr.Location = new System.Drawing.Point(24, 122);
            this.GridRetr.MainView = this.GridViewRetr;
            this.GridRetr.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.GridRetr.Name = "GridRetr";
            this.GridRetr.Size = new System.Drawing.Size(1011, 352);
            this.GridRetr.TabIndex = 3;
            this.GridRetr.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.GridViewRetr});
            // 
            // GridViewRetr
            // 
            this.GridViewRetr.ColumnPanelRowHeight = 30;
            this.GridViewRetr.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.GridColUsrId,
            this.GridColUsrNm,
            this.GridColDeptNm,
            this.GridColJkwiNm,
            this.GridColMoblNo,
            this.GridColRk});
            this.GridViewRetr.GridControl = this.GridRetr;
            this.GridViewRetr.IndicatorWidth = 40;
            this.GridViewRetr.Name = "GridViewRetr";
            this.GridViewRetr.OptionsFind.AllowFindPanel = false;
            this.GridViewRetr.OptionsSelection.MultiSelect = true;
            this.GridViewRetr.OptionsView.ColumnAutoWidth = false;
            this.GridViewRetr.OptionsView.ShowGroupPanel = false;
            this.GridViewRetr.RowClick += new DevExpress.XtraGrid.Views.Grid.RowClickEventHandler(this.GridViewRetr_RowClick);
            this.GridViewRetr.CustomDrawRowIndicator += new DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventHandler(this.GridViewRetr_CustomDrawRowIndicator);
            this.GridViewRetr.RowStyle += new DevExpress.XtraGrid.Views.Grid.RowStyleEventHandler(this.GridViewRetr_RowStyle);
            this.GridViewRetr.KeyDown += new System.Windows.Forms.KeyEventHandler(this.GridViewRetr_KeyDown);
            // 
            // GridColUsrId
            // 
            this.GridColUsrId.AppearanceCell.Options.UseTextOptions = true;
            this.GridColUsrId.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColUsrId.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColUsrId.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColUsrId.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColUsrId.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColUsrId.Caption = "ID";
            this.GridColUsrId.FieldName = "USRID";
            this.GridColUsrId.Name = "GridColUsrId";
            this.GridColUsrId.OptionsColumn.AllowEdit = false;
            this.GridColUsrId.Visible = true;
            this.GridColUsrId.VisibleIndex = 0;
            this.GridColUsrId.Width = 115;
            // 
            // GridColUsrNm
            // 
            this.GridColUsrNm.AppearanceCell.Options.UseTextOptions = true;
            this.GridColUsrNm.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColUsrNm.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColUsrNm.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColUsrNm.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColUsrNm.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColUsrNm.Caption = "사용자명";
            this.GridColUsrNm.FieldName = "USRNM";
            this.GridColUsrNm.Name = "GridColUsrNm";
            this.GridColUsrNm.OptionsColumn.AllowEdit = false;
            this.GridColUsrNm.Visible = true;
            this.GridColUsrNm.VisibleIndex = 1;
            this.GridColUsrNm.Width = 122;
            // 
            // GridColDeptNm
            // 
            this.GridColDeptNm.AppearanceCell.Options.UseTextOptions = true;
            this.GridColDeptNm.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColDeptNm.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColDeptNm.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColDeptNm.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColDeptNm.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColDeptNm.Caption = "부서";
            this.GridColDeptNm.FieldName = "DEPT_NM";
            this.GridColDeptNm.Name = "GridColDeptNm";
            this.GridColDeptNm.OptionsColumn.AllowEdit = false;
            this.GridColDeptNm.Visible = true;
            this.GridColDeptNm.VisibleIndex = 2;
            this.GridColDeptNm.Width = 114;
            // 
            // GridColJkwiNm
            // 
            this.GridColJkwiNm.AppearanceCell.Options.UseTextOptions = true;
            this.GridColJkwiNm.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColJkwiNm.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColJkwiNm.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColJkwiNm.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColJkwiNm.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColJkwiNm.Caption = "직위";
            this.GridColJkwiNm.FieldName = "JKWINM";
            this.GridColJkwiNm.Name = "GridColJkwiNm";
            this.GridColJkwiNm.OptionsColumn.AllowEdit = false;
            this.GridColJkwiNm.Visible = true;
            this.GridColJkwiNm.VisibleIndex = 3;
            this.GridColJkwiNm.Width = 91;
            // 
            // GridColMoblNo
            // 
            this.GridColMoblNo.AppearanceCell.Options.UseTextOptions = true;
            this.GridColMoblNo.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColMoblNo.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColMoblNo.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColMoblNo.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColMoblNo.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColMoblNo.Caption = "연락처";
            this.GridColMoblNo.FieldName = "MOBLNO";
            this.GridColMoblNo.Name = "GridColMoblNo";
            this.GridColMoblNo.OptionsColumn.AllowEdit = false;
            this.GridColMoblNo.Visible = true;
            this.GridColMoblNo.VisibleIndex = 4;
            this.GridColMoblNo.Width = 154;
            // 
            // GridColRk
            // 
            this.GridColRk.AppearanceCell.Options.UseTextOptions = true;
            this.GridColRk.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColRk.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColRk.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColRk.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColRk.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColRk.Caption = "비고";
            this.GridColRk.FieldName = "RK";
            this.GridColRk.Name = "GridColRk";
            this.GridColRk.OptionsColumn.AllowEdit = false;
            this.GridColRk.Visible = true;
            this.GridColRk.VisibleIndex = 5;
            this.GridColRk.Width = 226;
            // 
            // TxtFindWord
            // 
            this.TxtFindWord.EnterMoveNextControl = true;
            this.TxtFindWord.Location = new System.Drawing.Point(199, 44);
            this.TxtFindWord.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.TxtFindWord.Name = "TxtFindWord";
            this.TxtFindWord.Properties.AppearanceFocused.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.TxtFindWord.Properties.AppearanceFocused.Options.UseBackColor = true;
            this.TxtFindWord.Size = new System.Drawing.Size(120, 22);
            this.TxtFindWord.StyleController = this.layoutControl1;
            this.TxtFindWord.TabIndex = 1;
            // 
            // CboFindIdx
            // 
            this.CboFindIdx.EditValue = "사용자명";
            this.CboFindIdx.EnterMoveNextControl = true;
            this.CboFindIdx.Location = new System.Drawing.Point(75, 44);
            this.CboFindIdx.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.CboFindIdx.Name = "CboFindIdx";
            this.CboFindIdx.Properties.AppearanceFocused.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.CboFindIdx.Properties.AppearanceFocused.Options.UseBackColor = true;
            this.CboFindIdx.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.CboFindIdx.Properties.Items.AddRange(new object[] {
            "사용자명",
            "ID",
            "사용자코드",
            "부서",
            "직위"});
            this.CboFindIdx.Size = new System.Drawing.Size(120, 22);
            this.CboFindIdx.StyleController = this.layoutControl1;
            this.CboFindIdx.TabIndex = 0;
            this.CboFindIdx.TabStop = false;
            // 
            // BtnApply
            // 
            this.BtnApply.ImageOptions.Image = global::AccAdm.Properties.Resources.apply_16x16;
            this.BtnApply.Location = new System.Drawing.Point(873, 490);
            this.BtnApply.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BtnApply.Name = "BtnApply";
            this.BtnApply.Size = new System.Drawing.Size(85, 30);
            this.BtnApply.StyleController = this.layoutControl1;
            this.BtnApply.TabIndex = 7;
            this.BtnApply.Text = "적용(F3)";
            this.BtnApply.Click += new System.EventHandler(this.BtnApply_Click);
            // 
            // BtnClose
            // 
            this.BtnClose.ImageOptions.Image = global::AccAdm.Properties.Resources.cancel_16x16;
            this.BtnClose.Location = new System.Drawing.Point(962, 490);
            this.BtnClose.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(85, 30);
            this.BtnClose.StyleController = this.layoutControl1;
            this.BtnClose.TabIndex = 7;
            this.BtnClose.Text = "닫기(ESC)";
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlGroup1,
            this.layoutControlItem5,
            this.layoutControlItem6,
            this.emptySpaceItem2,
            this.layoutControlGroup2});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(1059, 532);
            this.Root.TextVisible = false;
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem2,
            this.layoutControlItem4,
            this.emptySpaceItem1});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(1039, 78);
            this.layoutControlGroup1.Text = "검색조건";
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.CboFindIdx;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.MaxSize = new System.Drawing.Size(175, 26);
            this.layoutControlItem1.MinSize = new System.Drawing.Size(175, 26);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(175, 34);
            this.layoutControlItem1.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem1.Text = "찾을항목";
            this.layoutControlItem1.TextSize = new System.Drawing.Size(48, 15);
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.TxtFindWord;
            this.layoutControlItem2.Location = new System.Drawing.Point(175, 0);
            this.layoutControlItem2.MaxSize = new System.Drawing.Size(124, 26);
            this.layoutControlItem2.MinSize = new System.Drawing.Size(124, 26);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(124, 34);
            this.layoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.BtnRetr;
            this.layoutControlItem4.Location = new System.Drawing.Point(926, 0);
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
            this.emptySpaceItem1.Location = new System.Drawing.Point(299, 0);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(627, 34);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.BtnApply;
            this.layoutControlItem5.CustomizationFormText = "layoutControlItem4";
            this.layoutControlItem5.Location = new System.Drawing.Point(861, 478);
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
            this.layoutControlItem6.Location = new System.Drawing.Point(950, 478);
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
            this.emptySpaceItem2.Location = new System.Drawing.Point(0, 478);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(861, 34);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlGroup2
            // 
            this.layoutControlGroup2.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem3});
            this.layoutControlGroup2.Location = new System.Drawing.Point(0, 78);
            this.layoutControlGroup2.Name = "layoutControlGroup2";
            this.layoutControlGroup2.Size = new System.Drawing.Size(1039, 400);
            this.layoutControlGroup2.Text = "사용자리스트";
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.GridRetr;
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(1015, 356);
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // PP002F00
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1059, 532);
            this.Controls.Add(this.layoutControl1);
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "PP002F00";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "사용자검색";
            this.Load += new System.EventHandler(this.PP005F00_Load);
            this.Shown += new System.EventHandler(this.PP005F00_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.PP005F00_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.GridRetr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewRetr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtFindWord.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboFindIdx.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraEditors.SimpleButton BtnRetr;
        private DevExpress.XtraGrid.GridControl GridRetr;
        private DevExpress.XtraGrid.Views.Grid.GridView GridViewRetr;
        private DevExpress.XtraEditors.TextEdit TxtFindWord;
        private DevExpress.XtraEditors.ComboBoxEdit CboFindIdx;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraEditors.SimpleButton BtnApply;
        private DevExpress.XtraEditors.SimpleButton BtnClose;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup2;
        private DevExpress.XtraGrid.Columns.GridColumn GridColUsrId;
        private DevExpress.XtraGrid.Columns.GridColumn GridColUsrNm;
        private DevExpress.XtraGrid.Columns.GridColumn GridColDeptNm;
        private DevExpress.XtraGrid.Columns.GridColumn GridColJkwiNm;
        private DevExpress.XtraGrid.Columns.GridColumn GridColMoblNo;
        private DevExpress.XtraGrid.Columns.GridColumn GridColRk;
    }
}