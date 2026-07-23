namespace AccAdm
{
    partial class GE003F01
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
            this.RadiYNGUB = new DevExpress.XtraEditors.RadioGroup();
            this.BtnSaveMt = new DevExpress.XtraEditors.SimpleButton();
            this.BtnClose = new DevExpress.XtraEditors.SimpleButton();
            this.BtnSave = new DevExpress.XtraEditors.SimpleButton();
            this.MemoRk = new DevExpress.XtraEditors.MemoEdit();
            this.TxtYncnt = new DevExpress.XtraEditors.TextEdit();
            this.DateBasdt = new DevExpress.XtraEditors.DateEdit();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem7 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.RadiYNGUB.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.MemoRk.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtYncnt.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DateBasdt.Properties.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DateBasdt.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem7)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.RadiYNGUB);
            this.layoutControl1.Controls.Add(this.BtnSaveMt);
            this.layoutControl1.Controls.Add(this.BtnClose);
            this.layoutControl1.Controls.Add(this.BtnSave);
            this.layoutControl1.Controls.Add(this.MemoRk);
            this.layoutControl1.Controls.Add(this.TxtYncnt);
            this.layoutControl1.Controls.Add(this.DateBasdt);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(340, 184);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // RadiYNGUB
            // 
            this.RadiYNGUB.EditValue = "0";
            this.RadiYNGUB.Location = new System.Drawing.Point(63, 12);
            this.RadiYNGUB.Name = "RadiYNGUB";
            this.RadiYNGUB.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
            new DevExpress.XtraEditors.Controls.RadioGroupItem("0", "연차"),
            new DevExpress.XtraEditors.Controls.RadioGroupItem("1", "특별휴가")});
            this.RadiYNGUB.Size = new System.Drawing.Size(265, 25);
            this.RadiYNGUB.StyleController = this.layoutControl1;
            this.RadiYNGUB.TabIndex = 10;
            // 
            // BtnSaveMt
            // 
            this.BtnSaveMt.ImageOptions.Image = global::AccAdm.Properties.Resources.saveall_16x16;
            this.BtnSaveMt.Location = new System.Drawing.Point(55, 142);
            this.BtnSaveMt.Name = "BtnSaveMt";
            this.BtnSaveMt.Size = new System.Drawing.Size(95, 30);
            this.BtnSaveMt.StyleController = this.layoutControl1;
            this.BtnSaveMt.TabIndex = 9;
            this.BtnSaveMt.Text = "연속저장(F1)";
            this.BtnSaveMt.Click += new System.EventHandler(this.BtnSaveMt_Click);
            // 
            // BtnClose
            // 
            this.BtnClose.ImageOptions.Image = global::AccAdm.Properties.Resources.cancel_16x16;
            this.BtnClose.Location = new System.Drawing.Point(243, 142);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(85, 30);
            this.BtnClose.StyleController = this.layoutControl1;
            this.BtnClose.TabIndex = 8;
            this.BtnClose.Text = "닫기";
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // BtnSave
            // 
            this.BtnSave.ImageOptions.Image = global::AccAdm.Properties.Resources.save_16x16;
            this.BtnSave.Location = new System.Drawing.Point(154, 142);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(85, 30);
            this.BtnSave.StyleController = this.layoutControl1;
            this.BtnSave.TabIndex = 7;
            this.BtnSave.Text = "저장(F3)";
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // MemoRk
            // 
            this.MemoRk.Location = new System.Drawing.Point(63, 67);
            this.MemoRk.Name = "MemoRk";
            this.MemoRk.Size = new System.Drawing.Size(265, 71);
            this.MemoRk.StyleController = this.layoutControl1;
            this.MemoRk.TabIndex = 6;
            // 
            // TxtYncnt
            // 
            this.TxtYncnt.EnterMoveNextControl = true;
            this.TxtYncnt.Location = new System.Drawing.Point(223, 41);
            this.TxtYncnt.Name = "TxtYncnt";
            this.TxtYncnt.Properties.Appearance.Options.UseTextOptions = true;
            this.TxtYncnt.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.TxtYncnt.Properties.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.TxtYncnt.Size = new System.Drawing.Size(105, 22);
            this.TxtYncnt.StyleController = this.layoutControl1;
            this.TxtYncnt.TabIndex = 5;
            // 
            // DateBasdt
            // 
            this.DateBasdt.EditValue = null;
            this.DateBasdt.EnterMoveNextControl = true;
            this.DateBasdt.Location = new System.Drawing.Point(63, 41);
            this.DateBasdt.Name = "DateBasdt";
            this.DateBasdt.Properties.AllowNullInput = DevExpress.Utils.DefaultBoolean.False;
            this.DateBasdt.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.DateBasdt.Properties.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.DateBasdt.Properties.DisplayFormat.FormatString = "yyyy-MM-dd";
            this.DateBasdt.Properties.DisplayFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.DateBasdt.Properties.EditFormat.FormatString = "yyyy-MM-dd";
            this.DateBasdt.Properties.EditFormat.FormatType = DevExpress.Utils.FormatType.DateTime;
            this.DateBasdt.Properties.Mask.EditMask = "yyyy-MM-dd";
            this.DateBasdt.Properties.Mask.MaskType = DevExpress.XtraEditors.Mask.MaskType.DateTimeAdvancingCaret;
            this.DateBasdt.Properties.MaxValue = new System.DateTime(9999, 12, 31, 23, 59, 59, 0);
            this.DateBasdt.Size = new System.Drawing.Size(105, 22);
            this.DateBasdt.StyleController = this.layoutControl1;
            this.DateBasdt.TabIndex = 4;
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.emptySpaceItem1,
            this.layoutControlItem2,
            this.layoutControlItem3,
            this.layoutControlItem4,
            this.layoutControlItem5,
            this.layoutControlItem6,
            this.layoutControlItem7,
            this.layoutControlItem1});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(340, 184);
            this.Root.TextVisible = false;
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(0, 130);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(43, 34);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.AppearanceItemCaption.Options.UseTextOptions = true;
            this.layoutControlItem2.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.layoutControlItem2.AppearanceItemCaption.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.layoutControlItem2.Control = this.TxtYncnt;
            this.layoutControlItem2.Location = new System.Drawing.Point(160, 29);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(160, 26);
            this.layoutControlItem2.Text = "발생수량";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(48, 15);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.AppearanceItemCaption.Options.UseTextOptions = true;
            this.layoutControlItem3.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.layoutControlItem3.AppearanceItemCaption.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.layoutControlItem3.Control = this.MemoRk;
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 55);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(320, 75);
            this.layoutControlItem3.Text = "비고";
            this.layoutControlItem3.TextSize = new System.Drawing.Size(48, 15);
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.Control = this.BtnSave;
            this.layoutControlItem4.Location = new System.Drawing.Point(142, 130);
            this.layoutControlItem4.MaxSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem4.MinSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(89, 34);
            this.layoutControlItem4.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem4.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem4.TextVisible = false;
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.Control = this.BtnClose;
            this.layoutControlItem5.Location = new System.Drawing.Point(231, 130);
            this.layoutControlItem5.MaxSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem5.MinSize = new System.Drawing.Size(89, 34);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(89, 34);
            this.layoutControlItem5.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem5.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem5.TextVisible = false;
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.BtnSaveMt;
            this.layoutControlItem6.Location = new System.Drawing.Point(43, 130);
            this.layoutControlItem6.MaxSize = new System.Drawing.Size(99, 34);
            this.layoutControlItem6.MinSize = new System.Drawing.Size(99, 34);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(99, 34);
            this.layoutControlItem6.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem6.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem6.TextVisible = false;
            // 
            // layoutControlItem7
            // 
            this.layoutControlItem7.AppearanceItemCaption.Options.UseTextOptions = true;
            this.layoutControlItem7.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.layoutControlItem7.AppearanceItemCaption.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.layoutControlItem7.Control = this.RadiYNGUB;
            this.layoutControlItem7.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem7.Name = "layoutControlItem7";
            this.layoutControlItem7.Size = new System.Drawing.Size(320, 29);
            this.layoutControlItem7.Text = "구분";
            this.layoutControlItem7.TextSize = new System.Drawing.Size(48, 15);
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.AppearanceItemCaption.Options.UseTextOptions = true;
            this.layoutControlItem1.AppearanceItemCaption.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.layoutControlItem1.AppearanceItemCaption.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.layoutControlItem1.Control = this.DateBasdt;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 29);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(160, 26);
            this.layoutControlItem1.Text = "발생일자";
            this.layoutControlItem1.TextSize = new System.Drawing.Size(48, 15);
            // 
            // GE003F01
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(340, 184);
            this.Controls.Add(this.layoutControl1);
            this.KeyPreview = true;
            this.Name = "GE003F01";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "연차발생관리";
            this.Load += new System.EventHandler(this.GE003F01_Load);
            this.Shown += new System.EventHandler(this.GE003F01_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.GE003F01_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.RadiYNGUB.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.MemoRk.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtYncnt.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DateBasdt.Properties.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DateBasdt.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem7)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraEditors.SimpleButton BtnSaveMt;
        private DevExpress.XtraEditors.SimpleButton BtnClose;
        private DevExpress.XtraEditors.SimpleButton BtnSave;
        private DevExpress.XtraEditors.MemoEdit MemoRk;
        private DevExpress.XtraEditors.TextEdit TxtYncnt;
        private DevExpress.XtraEditors.DateEdit DateBasdt;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraEditors.RadioGroup RadiYNGUB;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem7;
    }
}