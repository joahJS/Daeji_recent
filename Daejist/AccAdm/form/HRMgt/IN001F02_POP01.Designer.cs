namespace AccAdm
{
    partial class IN001F02_POP01
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
            this.BtnClose = new DevExpress.XtraEditors.SimpleButton();
            this.Txttotal = new DevExpress.XtraEditors.TextEdit();
            this.BtnRetr = new DevExpress.XtraEditors.SimpleButton();
            this.Cbfam = new DevExpress.XtraEditors.ComboBoxEdit();
            this.TxtCost = new DevExpress.XtraEditors.TextEdit();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem4 = new DevExpress.XtraLayout.EmptySpaceItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Txttotal.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Cbfam.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtCost.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem4)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.AutoScroll = false;
            this.layoutControl1.Controls.Add(this.BtnClose);
            this.layoutControl1.Controls.Add(this.Txttotal);
            this.layoutControl1.Controls.Add(this.BtnRetr);
            this.layoutControl1.Controls.Add(this.Cbfam);
            this.layoutControl1.Controls.Add(this.TxtCost);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Margin = new System.Windows.Forms.Padding(2);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(368, 193);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // BtnClose
            // 
            this.BtnClose.ImageOptions.Image = global::AccAdm.Properties.Resources.cancel_16x16;
            this.BtnClose.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.LeftCenter;
            this.BtnClose.Location = new System.Drawing.Point(274, 12);
            this.BtnClose.Margin = new System.Windows.Forms.Padding(2);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(82, 26);
            this.BtnClose.StyleController = this.layoutControl1;
            this.BtnClose.TabIndex = 8;
            this.BtnClose.TabStop = false;
            this.BtnClose.Text = "닫기";
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // Txttotal
            // 
            this.Txttotal.Location = new System.Drawing.Point(218, 147);
            this.Txttotal.Margin = new System.Windows.Forms.Padding(2);
            this.Txttotal.Name = "Txttotal";
            this.Txttotal.Properties.Appearance.Options.UseTextOptions = true;
            this.Txttotal.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.Txttotal.Properties.Mask.EditMask = "n0";
            this.Txttotal.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
            this.Txttotal.Properties.Mask.UseMaskAsDisplayFormat = true;
            this.Txttotal.Properties.ReadOnly = true;
            this.Txttotal.Size = new System.Drawing.Size(126, 22);
            this.Txttotal.StyleController = this.layoutControl1;
            this.Txttotal.TabIndex = 7;
            // 
            // BtnRetr
            // 
            this.BtnRetr.ImageOptions.Image = global::AccAdm.Properties.Resources.zoom_16x16;
            this.BtnRetr.ImageOptions.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.LeftCenter;
            this.BtnRetr.Location = new System.Drawing.Point(260, 100);
            this.BtnRetr.Margin = new System.Windows.Forms.Padding(2);
            this.BtnRetr.Name = "BtnRetr";
            this.BtnRetr.Size = new System.Drawing.Size(82, 22);
            this.BtnRetr.StyleController = this.layoutControl1;
            this.BtnRetr.TabIndex = 6;
            this.BtnRetr.Text = "조회";
            this.BtnRetr.Click += new System.EventHandler(this.BtnRetr_Click);
            // 
            // Cbfam
            // 
            this.Cbfam.EditValue = "1";
            this.Cbfam.EnterMoveNextControl = true;
            this.Cbfam.Location = new System.Drawing.Point(218, 100);
            this.Cbfam.Margin = new System.Windows.Forms.Padding(2);
            this.Cbfam.Name = "Cbfam";
            this.Cbfam.Properties.Appearance.Options.UseTextOptions = true;
            this.Cbfam.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.Cbfam.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.Cbfam.Properties.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11"});
            this.Cbfam.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor;
            this.Cbfam.Size = new System.Drawing.Size(38, 22);
            this.Cbfam.StyleController = this.layoutControl1;
            this.Cbfam.TabIndex = 5;
            // 
            // TxtCost
            // 
            this.TxtCost.EnterMoveNextControl = true;
            this.TxtCost.Location = new System.Drawing.Point(218, 74);
            this.TxtCost.Margin = new System.Windows.Forms.Padding(2);
            this.TxtCost.Name = "TxtCost";
            this.TxtCost.Properties.Appearance.Options.UseTextOptions = true;
            this.TxtCost.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.TxtCost.Properties.Mask.EditMask = "n0";
            this.TxtCost.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.Numeric;
            this.TxtCost.Properties.Mask.UseMaskAsDisplayFormat = true;
            this.TxtCost.Properties.MaxLength = 9;
            this.TxtCost.Size = new System.Drawing.Size(126, 22);
            this.TxtCost.StyleController = this.layoutControl1;
            this.TxtCost.TabIndex = 4;
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlGroup1,
            this.layoutControlItem5,
            this.emptySpaceItem4});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(368, 193);
            this.Root.TextVisible = false;
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem2,
            this.emptySpaceItem2,
            this.layoutControlItem4,
            this.layoutControlItem1,
            this.layoutControlItem3});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 30);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(348, 143);
            this.layoutControlGroup1.Text = "기본사항";
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.Cbfam;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 26);
            this.layoutControlItem2.MaxSize = new System.Drawing.Size(236, 26);
            this.layoutControlItem2.MinSize = new System.Drawing.Size(236, 26);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(236, 26);
            this.layoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem2.Text = "전체 공제대상 가족 수 (본인포함) : ";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(191, 15);
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(0, 52);
            this.emptySpaceItem2.MaxSize = new System.Drawing.Size(321, 21);
            this.emptySpaceItem2.MinSize = new System.Drawing.Size(321, 21);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(324, 21);
            this.emptySpaceItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.Txttotal;
            this.layoutControlItem4.Location = new System.Drawing.Point(0, 73);
            this.layoutControlItem4.MaxSize = new System.Drawing.Size(0, 26);
            this.layoutControlItem4.MinSize = new System.Drawing.Size(1, 26);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(324, 26);
            this.layoutControlItem4.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem4.Text = "한달에 납부하는 세금(원) : ";
            this.layoutControlItem4.TextSize = new System.Drawing.Size(191, 15);
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.TxtCost;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.MaxSize = new System.Drawing.Size(0, 26);
            this.layoutControlItem1.MinSize = new System.Drawing.Size(1, 26);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(324, 26);
            this.layoutControlItem1.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem1.Text = "월 급여액(원) : ";
            this.layoutControlItem1.TextSize = new System.Drawing.Size(191, 15);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.BtnRetr;
            this.layoutControlItem3.Location = new System.Drawing.Point(236, 26);
            this.layoutControlItem3.MaxSize = new System.Drawing.Size(86, 26);
            this.layoutControlItem3.MinSize = new System.Drawing.Size(86, 26);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(88, 26);
            this.layoutControlItem3.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.BtnClose;
            this.layoutControlItem5.Location = new System.Drawing.Point(262, 0);
            this.layoutControlItem5.MaxSize = new System.Drawing.Size(86, 30);
            this.layoutControlItem5.MinSize = new System.Drawing.Size(86, 30);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(86, 30);
            this.layoutControlItem5.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // emptySpaceItem4
            // 
            this.emptySpaceItem4.AllowHotTrack = false;
            this.emptySpaceItem4.Location = new System.Drawing.Point(0, 0);
            this.emptySpaceItem4.Name = "emptySpaceItem4";
            this.emptySpaceItem4.Size = new System.Drawing.Size(262, 30);
            this.emptySpaceItem4.TextSize = new System.Drawing.Size(0, 0);
            // 
            // IN001F02_POP01
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(368, 193);
            this.Controls.Add(this.layoutControl1);
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "IN001F02_POP01";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "간단조회";
            this.Load += new System.EventHandler(this.IN001F02_POP01_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.IN001F02_POP01_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.Txttotal.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Cbfam.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtCost.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem4)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraEditors.TextEdit Txttotal;
        private DevExpress.XtraEditors.SimpleButton BtnRetr;
        private DevExpress.XtraEditors.ComboBoxEdit Cbfam;
        private DevExpress.XtraEditors.TextEdit TxtCost;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraEditors.SimpleButton BtnClose;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem4;
    }
}