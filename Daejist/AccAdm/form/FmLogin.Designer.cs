namespace AccAdm
{
    partial class FmLogin
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FmLogin));
            this.TxtId = new DevExpress.XtraEditors.TextEdit();
            this.TxtPw = new DevExpress.XtraEditors.TextEdit();
            this.ChkSave = new DevExpress.XtraEditors.CheckEdit();
            this.BtnLogin = new DevExpress.XtraEditors.LabelControl();
            this.BtnExit = new DevExpress.XtraEditors.LabelControl();
            this.LblVersion = new DevExpress.XtraEditors.LabelControl();
            ((System.ComponentModel.ISupportInitialize)(this.TxtId.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtPw.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ChkSave.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // TxtId
            // 
            this.TxtId.EnterMoveNextControl = true;
            this.TxtId.Location = new System.Drawing.Point(323, 134);
            this.TxtId.Name = "TxtId";
            this.TxtId.Properties.Appearance.BorderColor = System.Drawing.Color.White;
            this.TxtId.Properties.Appearance.Font = new System.Drawing.Font("맑은 고딕", 11F);
            this.TxtId.Properties.Appearance.Options.UseBorderColor = true;
            this.TxtId.Properties.Appearance.Options.UseFont = true;
            this.TxtId.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.TxtId.Size = new System.Drawing.Size(138, 24);
            this.TxtId.TabIndex = 0;
            // 
            // TxtPw
            // 
            this.TxtPw.EnterMoveNextControl = true;
            this.TxtPw.Location = new System.Drawing.Point(324, 194);
            this.TxtPw.Name = "TxtPw";
            this.TxtPw.Properties.Appearance.Font = new System.Drawing.Font("맑은 고딕", 11F);
            this.TxtPw.Properties.Appearance.Options.UseFont = true;
            this.TxtPw.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.TxtPw.Properties.UseSystemPasswordChar = true;
            this.TxtPw.Size = new System.Drawing.Size(138, 24);
            this.TxtPw.TabIndex = 1;
            this.TxtPw.KeyDown += new System.Windows.Forms.KeyEventHandler(this.TxtPw_KeyDown);
            // 
            // ChkSave
            // 
            this.ChkSave.Location = new System.Drawing.Point(478, 138);
            this.ChkSave.Name = "ChkSave";
            this.ChkSave.Properties.Appearance.Options.UseFont = true;
            this.ChkSave.Properties.Caption = "";
            this.ChkSave.Size = new System.Drawing.Size(93, 19);
            this.ChkSave.TabIndex = 7;
            // 
            // BtnLogin
            // 
            this.BtnLogin.Location = new System.Drawing.Point(427, 266);
            this.BtnLogin.MinimumSize = new System.Drawing.Size(65, 18);
            this.BtnLogin.Name = "BtnLogin";
            this.BtnLogin.Size = new System.Drawing.Size(65, 18);
            this.BtnLogin.TabIndex = 8;
            this.BtnLogin.Click += new System.EventHandler(this.BtnLogin_Click);
            // 
            // BtnExit
            // 
            this.BtnExit.Location = new System.Drawing.Point(501, 267);
            this.BtnExit.MinimumSize = new System.Drawing.Size(57, 18);
            this.BtnExit.Name = "BtnExit";
            this.BtnExit.Size = new System.Drawing.Size(57, 18);
            this.BtnExit.TabIndex = 9;
            this.BtnExit.Click += new System.EventHandler(this.BtnExit_Click);
            // 
            // LblVersion
            // 
            this.LblVersion.Location = new System.Drawing.Point(247, 289);
            this.LblVersion.Name = "LblVersion";
            this.LblVersion.Size = new System.Drawing.Size(21, 15);
            this.LblVersion.TabIndex = 10;
            this.LblVersion.Text = "111";
            // 
            // FmLogin
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImageLayoutStore = System.Windows.Forms.ImageLayout.Stretch;
            this.BackgroundImageStore = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImageStore")));
            this.ClientSize = new System.Drawing.Size(607, 308);
            this.Controls.Add(this.LblVersion);
            this.Controls.Add(this.BtnExit);
            this.Controls.Add(this.BtnLogin);
            this.Controls.Add(this.ChkSave);
            this.Controls.Add(this.TxtPw);
            this.Controls.Add(this.TxtId);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(617, 340);
            this.MinimumSize = new System.Drawing.Size(617, 340);
            this.Name = "FmLogin";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "로그인";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FmLogin_FormClosed);
            this.Load += new System.EventHandler(this.FmLogin_Load);
            ((System.ComponentModel.ISupportInitialize)(this.TxtId.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtPw.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ChkSave.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraEditors.TextEdit TxtId;
        private DevExpress.XtraEditors.TextEdit TxtPw;
        private DevExpress.XtraEditors.CheckEdit ChkSave;
        private DevExpress.XtraEditors.LabelControl BtnLogin;
        private DevExpress.XtraEditors.LabelControl BtnExit;
        private DevExpress.XtraEditors.LabelControl LblVersion;
    }
}