namespace Daeji_UPDATE
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.LblProgress = new DevExpress.XtraEditors.LabelControl();
            this.BtnUpdate = new DevExpress.XtraEditors.SimpleButton();
            this.BtnClose = new DevExpress.XtraEditors.SimpleButton();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.LayoutUpdate = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LayoutUpdate)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.AutoScroll = false;
            this.layoutControl1.Controls.Add(this.LblProgress);
            this.layoutControl1.Controls.Add(this.BtnUpdate);
            this.layoutControl1.Controls.Add(this.BtnClose);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(448, 77);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // LblProgress
            // 
            this.LblProgress.Appearance.ForeColor = System.Drawing.Color.Black;
            this.LblProgress.Appearance.Options.UseFont = true;
            this.LblProgress.Appearance.Options.UseForeColor = true;
            this.LblProgress.Location = new System.Drawing.Point(12, 11);
            this.LblProgress.Name = "LblProgress";
            this.LblProgress.Size = new System.Drawing.Size(7, 14);
            this.LblProgress.StyleController = this.layoutControl1;
            this.LblProgress.TabIndex = 6;
            this.LblProgress.Text = "1";
            // 
            // BtnUpdate
            // 
            this.BtnUpdate.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold);
            this.BtnUpdate.Appearance.ForeColor = System.Drawing.Color.Black;
            this.BtnUpdate.Appearance.Options.UseFont = true;
            this.BtnUpdate.Appearance.Options.UseForeColor = true;
            this.BtnUpdate.AppearanceDisabled.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold);
            this.BtnUpdate.AppearanceDisabled.ForeColor = System.Drawing.Color.Black;
            this.BtnUpdate.AppearanceDisabled.Options.UseFont = true;
            this.BtnUpdate.AppearanceDisabled.Options.UseForeColor = true;
            this.BtnUpdate.AppearanceHovered.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold);
            this.BtnUpdate.AppearanceHovered.ForeColor = System.Drawing.Color.Black;
            this.BtnUpdate.AppearanceHovered.Options.UseFont = true;
            this.BtnUpdate.AppearanceHovered.Options.UseForeColor = true;
            this.BtnUpdate.AppearancePressed.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold);
            this.BtnUpdate.AppearancePressed.ForeColor = System.Drawing.Color.Black;
            this.BtnUpdate.AppearancePressed.Options.UseFont = true;
            this.BtnUpdate.AppearancePressed.Options.UseForeColor = true;
            this.BtnUpdate.Location = new System.Drawing.Point(272, 29);
            this.BtnUpdate.Name = "BtnUpdate";
            this.BtnUpdate.Size = new System.Drawing.Size(80, 35);
            this.BtnUpdate.StyleController = this.layoutControl1;
            this.BtnUpdate.TabIndex = 4;
            this.BtnUpdate.Text = "업데이트";
            this.BtnUpdate.Click += new System.EventHandler(this.BtnUpdate_Click);
            // 
            // BtnClose
            // 
            this.BtnClose.Appearance.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold);
            this.BtnClose.Appearance.ForeColor = System.Drawing.Color.Black;
            this.BtnClose.Appearance.Options.UseFont = true;
            this.BtnClose.Appearance.Options.UseForeColor = true;
            this.BtnClose.AppearanceDisabled.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold);
            this.BtnClose.AppearanceDisabled.ForeColor = System.Drawing.Color.Black;
            this.BtnClose.AppearanceDisabled.Options.UseFont = true;
            this.BtnClose.AppearanceDisabled.Options.UseForeColor = true;
            this.BtnClose.AppearanceHovered.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold);
            this.BtnClose.AppearanceHovered.ForeColor = System.Drawing.Color.Black;
            this.BtnClose.AppearanceHovered.Options.UseFont = true;
            this.BtnClose.AppearanceHovered.Options.UseForeColor = true;
            this.BtnClose.AppearancePressed.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold);
            this.BtnClose.AppearancePressed.ForeColor = System.Drawing.Color.Black;
            this.BtnClose.AppearancePressed.Options.UseFont = true;
            this.BtnClose.AppearancePressed.Options.UseForeColor = true;
            this.BtnClose.Location = new System.Drawing.Point(356, 29);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(80, 35);
            this.BtnClose.StyleController = this.layoutControl1;
            this.BtnClose.TabIndex = 4;
            this.BtnClose.Text = "닫기";
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.LayoutUpdate,
            this.layoutControlItem2,
            this.emptySpaceItem1});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(448, 77);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.LblProgress;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(428, 18);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // LayoutUpdate
            // 
            this.LayoutUpdate.Control = this.BtnUpdate;
            this.LayoutUpdate.CustomizationFormText = "LayoutUpdate";
            this.LayoutUpdate.Location = new System.Drawing.Point(260, 18);
            this.LayoutUpdate.MaxSize = new System.Drawing.Size(84, 39);
            this.LayoutUpdate.MinSize = new System.Drawing.Size(84, 36);
            this.LayoutUpdate.Name = "LayoutUpdate";
            this.LayoutUpdate.OptionsPrint.AppearanceItemCaption.Options.UseFont = true;
            this.LayoutUpdate.Size = new System.Drawing.Size(84, 41);
            this.LayoutUpdate.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.LayoutUpdate.TextSize = new System.Drawing.Size(0, 0);
            this.LayoutUpdate.TextVisible = false;
            this.LayoutUpdate.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.BtnClose;
            this.layoutControlItem2.CustomizationFormText = "layoutControlItem1";
            this.layoutControlItem2.Location = new System.Drawing.Point(344, 18);
            this.layoutControlItem2.MaxSize = new System.Drawing.Size(84, 39);
            this.layoutControlItem2.MinSize = new System.Drawing.Size(84, 36);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.OptionsPrint.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem2.Size = new System.Drawing.Size(84, 41);
            this.layoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem2.Text = "layoutControlItem1";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(0, 18);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(260, 41);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(448, 77);
            this.Controls.Add(this.layoutControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "대지에스텍 업데이트 체크";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LayoutUpdate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraEditors.LabelControl LblProgress;
        private DevExpress.XtraEditors.SimpleButton BtnUpdate;
        private DevExpress.XtraEditors.SimpleButton BtnClose;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem LayoutUpdate;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
    }
}

