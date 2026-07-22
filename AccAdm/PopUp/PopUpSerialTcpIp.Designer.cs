namespace AccAdm
{
    partial class PopUpSerialTcpIp
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
            this.components = new System.ComponentModel.Container();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.TxtPortNum = new DevExpress.XtraEditors.TextEdit();
            this.TxtTcpIP = new DevExpress.XtraEditors.TextEdit();
            this.BtnPortOpen = new DevExpress.XtraEditors.SimpleButton();
            this.CboComPort = new DevExpress.XtraEditors.ComboBoxEdit();
            this.CboBaudRate = new DevExpress.XtraEditors.ComboBoxEdit();
            this.CboDataBits = new DevExpress.XtraEditors.ComboBoxEdit();
            this.CboParity = new DevExpress.XtraEditors.ComboBoxEdit();
            this.CboStopBits = new DevExpress.XtraEditors.ComboBoxEdit();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem1 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem4 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem5 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem4 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem6 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem5 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem7 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.emptySpaceItem8 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlGroup2 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem8 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem2 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem9 = new DevExpress.XtraLayout.LayoutControlItem();
            this.emptySpaceItem3 = new DevExpress.XtraLayout.EmptySpaceItem();
            this.layoutControlItem6 = new DevExpress.XtraLayout.LayoutControlItem();
            this.serialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.emptySpaceItem9 = new DevExpress.XtraLayout.EmptySpaceItem();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TxtPortNum.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtTcpIP.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboComPort.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboBaudRate.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboDataBits.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboParity.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboStopBits.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem7)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem8)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem8)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem9)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem9)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.TxtPortNum);
            this.layoutControl1.Controls.Add(this.TxtTcpIP);
            this.layoutControl1.Controls.Add(this.BtnPortOpen);
            this.layoutControl1.Controls.Add(this.CboComPort);
            this.layoutControl1.Controls.Add(this.CboBaudRate);
            this.layoutControl1.Controls.Add(this.CboDataBits);
            this.layoutControl1.Controls.Add(this.CboParity);
            this.layoutControl1.Controls.Add(this.CboStopBits);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(432, 260);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // TxtPortNum
            // 
            this.TxtPortNum.EditValue = "8080";
            this.TxtPortNum.Location = new System.Drawing.Point(299, 80);
            this.TxtPortNum.Name = "TxtPortNum";
            this.TxtPortNum.Properties.Appearance.Options.UseFont = true;
            this.TxtPortNum.Size = new System.Drawing.Size(109, 22);
            this.TxtPortNum.StyleController = this.layoutControl1;
            this.TxtPortNum.TabIndex = 11;
            // 
            // TxtTcpIP
            // 
            this.TxtTcpIP.EditValue = "192.168.43.241";
            this.TxtTcpIP.Location = new System.Drawing.Point(299, 44);
            this.TxtTcpIP.Name = "TxtTcpIP";
            this.TxtTcpIP.Properties.Appearance.Options.UseFont = true;
            this.TxtTcpIP.Size = new System.Drawing.Size(109, 22);
            this.TxtTcpIP.StyleController = this.layoutControl1;
            this.TxtTcpIP.TabIndex = 10;
            // 
            // BtnPortOpen
            // 
            this.BtnPortOpen.Appearance.Options.UseFont = true;
            this.BtnPortOpen.Location = new System.Drawing.Point(232, 173);
            this.BtnPortOpen.Name = "BtnPortOpen";
            this.BtnPortOpen.Size = new System.Drawing.Size(176, 63);
            this.BtnPortOpen.StyleController = this.layoutControl1;
            this.BtnPortOpen.TabIndex = 8;
            this.BtnPortOpen.Text = "Port Open";
            this.BtnPortOpen.Click += new System.EventHandler(this.BtnPortOpen_Click);
            // 
            // CboComPort
            // 
            this.CboComPort.Location = new System.Drawing.Point(91, 44);
            this.CboComPort.Name = "CboComPort";
            this.CboComPort.Properties.Appearance.Options.UseFont = true;
            this.CboComPort.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.CboComPort.Size = new System.Drawing.Size(113, 22);
            this.CboComPort.StyleController = this.layoutControl1;
            this.CboComPort.TabIndex = 4;
            this.CboComPort.SelectedIndexChanged += new System.EventHandler(this.CboComPort_SelectedIndexChanged);
            // 
            // CboBaudRate
            // 
            this.CboBaudRate.EditValue = "9600";
            this.CboBaudRate.Location = new System.Drawing.Point(91, 81);
            this.CboBaudRate.Name = "CboBaudRate";
            this.CboBaudRate.Properties.Appearance.Options.UseFont = true;
            this.CboBaudRate.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.CboBaudRate.Properties.Items.AddRange(new object[] {
            "9600",
            "14400",
            "19200",
            "38400",
            "57600",
            "115200",
            "128000"});
            this.CboBaudRate.Size = new System.Drawing.Size(113, 22);
            this.CboBaudRate.StyleController = this.layoutControl1;
            this.CboBaudRate.TabIndex = 4;
            this.CboBaudRate.SelectedIndexChanged += new System.EventHandler(this.CboBaudRate_SelectedIndexChanged);
            // 
            // CboDataBits
            // 
            this.CboDataBits.EditValue = "7";
            this.CboDataBits.Location = new System.Drawing.Point(91, 118);
            this.CboDataBits.Name = "CboDataBits";
            this.CboDataBits.Properties.Appearance.Options.UseFont = true;
            this.CboDataBits.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.CboDataBits.Properties.Items.AddRange(new object[] {
            "8",
            "7"});
            this.CboDataBits.Size = new System.Drawing.Size(113, 22);
            this.CboDataBits.StyleController = this.layoutControl1;
            this.CboDataBits.TabIndex = 5;
            this.CboDataBits.SelectedIndexChanged += new System.EventHandler(this.CboDataBits_SelectedIndexChanged);
            // 
            // CboParity
            // 
            this.CboParity.EditValue = "None";
            this.CboParity.Location = new System.Drawing.Point(91, 155);
            this.CboParity.Name = "CboParity";
            this.CboParity.Properties.Appearance.Options.UseFont = true;
            this.CboParity.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.CboParity.Properties.Items.AddRange(new object[] {
            "Even",
            "Mark",
            "None",
            "Odd",
            "Space",
            "None"});
            this.CboParity.Size = new System.Drawing.Size(113, 22);
            this.CboParity.StyleController = this.layoutControl1;
            this.CboParity.TabIndex = 6;
            this.CboParity.SelectedIndexChanged += new System.EventHandler(this.CboParity_SelectedIndexChanged);
            // 
            // CboStopBits
            // 
            this.CboStopBits.EditValue = "1";
            this.CboStopBits.Location = new System.Drawing.Point(91, 191);
            this.CboStopBits.Name = "CboStopBits";
            this.CboStopBits.Properties.Appearance.Options.UseFont = true;
            this.CboStopBits.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.CboStopBits.Properties.Items.AddRange(new object[] {
            "None",
            "1",
            "1.5",
            "2"});
            this.CboStopBits.Size = new System.Drawing.Size(113, 22);
            this.CboStopBits.StyleController = this.layoutControl1;
            this.CboStopBits.TabIndex = 7;
            this.CboStopBits.SelectedIndexChanged += new System.EventHandler(this.CboStopBits_SelectedIndexChanged);
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlGroup1,
            this.layoutControlGroup2});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(432, 260);
            this.Root.TextVisible = false;
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.AppearanceGroup.Options.UseFont = true;
            this.layoutControlGroup1.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.emptySpaceItem1,
            this.layoutControlItem2,
            this.emptySpaceItem4,
            this.layoutControlItem3,
            this.emptySpaceItem5,
            this.layoutControlItem4,
            this.emptySpaceItem6,
            this.layoutControlItem5,
            this.emptySpaceItem7,
            this.emptySpaceItem8});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(208, 240);
            this.layoutControlGroup1.Text = "시리얼 포트 설정";
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem1.Control = this.CboComPort;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.MinSize = new System.Drawing.Size(50, 27);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(184, 27);
            this.layoutControlItem1.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem1.Text = "Com Port";
            this.layoutControlItem1.TextSize = new System.Drawing.Size(64, 15);
            // 
            // emptySpaceItem1
            // 
            this.emptySpaceItem1.AllowHotTrack = false;
            this.emptySpaceItem1.Location = new System.Drawing.Point(0, 27);
            this.emptySpaceItem1.Name = "emptySpaceItem1";
            this.emptySpaceItem1.Size = new System.Drawing.Size(184, 10);
            this.emptySpaceItem1.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem2.Control = this.CboBaudRate;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 37);
            this.layoutControlItem2.MinSize = new System.Drawing.Size(50, 27);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(184, 27);
            this.layoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem2.Text = "Baud Rate";
            this.layoutControlItem2.TextSize = new System.Drawing.Size(64, 15);
            // 
            // emptySpaceItem4
            // 
            this.emptySpaceItem4.AllowHotTrack = false;
            this.emptySpaceItem4.Location = new System.Drawing.Point(0, 64);
            this.emptySpaceItem4.Name = "emptySpaceItem4";
            this.emptySpaceItem4.Size = new System.Drawing.Size(184, 10);
            this.emptySpaceItem4.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem3.Control = this.CboDataBits;
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 74);
            this.layoutControlItem3.MinSize = new System.Drawing.Size(50, 27);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(184, 27);
            this.layoutControlItem3.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem3.Text = "DataBits";
            this.layoutControlItem3.TextSize = new System.Drawing.Size(64, 15);
            // 
            // emptySpaceItem5
            // 
            this.emptySpaceItem5.AllowHotTrack = false;
            this.emptySpaceItem5.Location = new System.Drawing.Point(0, 101);
            this.emptySpaceItem5.Name = "emptySpaceItem5";
            this.emptySpaceItem5.Size = new System.Drawing.Size(184, 10);
            this.emptySpaceItem5.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem4
            // 
            this.layoutControlItem4.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem4.Control = this.CboParity;
            this.layoutControlItem4.Location = new System.Drawing.Point(0, 111);
            this.layoutControlItem4.Name = "layoutControlItem4";
            this.layoutControlItem4.Size = new System.Drawing.Size(184, 26);
            this.layoutControlItem4.Text = "Parity";
            this.layoutControlItem4.TextSize = new System.Drawing.Size(64, 15);
            // 
            // emptySpaceItem6
            // 
            this.emptySpaceItem6.AllowHotTrack = false;
            this.emptySpaceItem6.Location = new System.Drawing.Point(0, 137);
            this.emptySpaceItem6.Name = "emptySpaceItem6";
            this.emptySpaceItem6.Size = new System.Drawing.Size(184, 10);
            this.emptySpaceItem6.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem5
            // 
            this.layoutControlItem5.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem5.Control = this.CboStopBits;
            this.layoutControlItem5.Location = new System.Drawing.Point(0, 147);
            this.layoutControlItem5.MaxSize = new System.Drawing.Size(184, 26);
            this.layoutControlItem5.MinSize = new System.Drawing.Size(184, 26);
            this.layoutControlItem5.Name = "layoutControlItem5";
            this.layoutControlItem5.Size = new System.Drawing.Size(184, 26);
            this.layoutControlItem5.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem5.Text = "Stop Bits";
            this.layoutControlItem5.TextSize = new System.Drawing.Size(64, 15);
            // 
            // emptySpaceItem7
            // 
            this.emptySpaceItem7.AllowHotTrack = false;
            this.emptySpaceItem7.Location = new System.Drawing.Point(0, 173);
            this.emptySpaceItem7.Name = "emptySpaceItem7";
            this.emptySpaceItem7.Size = new System.Drawing.Size(184, 10);
            this.emptySpaceItem7.TextSize = new System.Drawing.Size(0, 0);
            // 
            // emptySpaceItem8
            // 
            this.emptySpaceItem8.AllowHotTrack = false;
            this.emptySpaceItem8.Location = new System.Drawing.Point(0, 183);
            this.emptySpaceItem8.Name = "emptySpaceItem8";
            this.emptySpaceItem8.Size = new System.Drawing.Size(184, 13);
            this.emptySpaceItem8.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlGroup2
            // 
            this.layoutControlGroup2.AppearanceGroup.Options.UseFont = true;
            this.layoutControlGroup2.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlGroup2.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem8,
            this.emptySpaceItem2,
            this.layoutControlItem9,
            this.emptySpaceItem3,
            this.layoutControlItem6});
            this.layoutControlGroup2.Location = new System.Drawing.Point(208, 0);
            this.layoutControlGroup2.Name = "layoutControlGroup2";
            this.layoutControlGroup2.Size = new System.Drawing.Size(204, 240);
            this.layoutControlGroup2.Text = "TCP/IP 설정";
            // 
            // layoutControlItem8
            // 
            this.layoutControlItem8.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem8.Control = this.TxtTcpIP;
            this.layoutControlItem8.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem8.Name = "layoutControlItem8";
            this.layoutControlItem8.Size = new System.Drawing.Size(180, 26);
            this.layoutControlItem8.Text = "TCP/IP 설정";
            this.layoutControlItem8.TextSize = new System.Drawing.Size(64, 15);
            // 
            // emptySpaceItem2
            // 
            this.emptySpaceItem2.AllowHotTrack = false;
            this.emptySpaceItem2.Location = new System.Drawing.Point(0, 26);
            this.emptySpaceItem2.Name = "emptySpaceItem2";
            this.emptySpaceItem2.Size = new System.Drawing.Size(180, 10);
            this.emptySpaceItem2.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem9
            // 
            this.layoutControlItem9.AppearanceItemCaption.Options.UseFont = true;
            this.layoutControlItem9.Control = this.TxtPortNum;
            this.layoutControlItem9.Location = new System.Drawing.Point(0, 36);
            this.layoutControlItem9.Name = "layoutControlItem9";
            this.layoutControlItem9.Size = new System.Drawing.Size(180, 26);
            this.layoutControlItem9.Text = "포트번호";
            this.layoutControlItem9.TextSize = new System.Drawing.Size(64, 15);
            // 
            // emptySpaceItem3
            // 
            this.emptySpaceItem3.AllowHotTrack = false;
            this.emptySpaceItem3.Location = new System.Drawing.Point(0, 62);
            this.emptySpaceItem3.Name = "emptySpaceItem3";
            this.emptySpaceItem3.Size = new System.Drawing.Size(180, 67);
            this.emptySpaceItem3.TextSize = new System.Drawing.Size(0, 0);
            // 
            // layoutControlItem6
            // 
            this.layoutControlItem6.Control = this.BtnPortOpen;
            this.layoutControlItem6.Location = new System.Drawing.Point(0, 129);
            this.layoutControlItem6.MinSize = new System.Drawing.Size(70, 26);
            this.layoutControlItem6.Name = "layoutControlItem6";
            this.layoutControlItem6.Size = new System.Drawing.Size(180, 67);
            this.layoutControlItem6.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom;
            this.layoutControlItem6.Text = "Port Open";
            this.layoutControlItem6.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem6.TextVisible = false;
            // 
            // emptySpaceItem9
            // 
            this.emptySpaceItem9.AllowHotTrack = false;
            this.emptySpaceItem9.Location = new System.Drawing.Point(85, 204);
            this.emptySpaceItem9.Name = "emptySpaceItem9";
            this.emptySpaceItem9.Size = new System.Drawing.Size(99, 46);
            this.emptySpaceItem9.TextSize = new System.Drawing.Size(0, 0);
            // 
            // PopUpSerialTcpIp
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(432, 260);
            this.Controls.Add(this.layoutControl1);
            this.Name = "PopUpSerialTcpIp";
            this.Text = "시리얼포트 및 TCP/IP 설정";
            this.Load += new System.EventHandler(this.PopUpSerialTcpIp_Load);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.TxtPortNum.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TxtTcpIP.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboComPort.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboBaudRate.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboDataBits.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboParity.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CboStopBits.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem7)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem8)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem8)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem9)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.emptySpaceItem9)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem4;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem5;
        private DevExpress.XtraEditors.SimpleButton BtnPortOpen;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem4;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem6;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem5;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem7;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem6;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem8;
        private DevExpress.XtraEditors.TextEdit TxtPortNum;
        private DevExpress.XtraEditors.TextEdit TxtTcpIP;
        private DevExpress.XtraEditors.ComboBoxEdit CboComPort;
        private DevExpress.XtraEditors.ComboBoxEdit CboBaudRate;
        private DevExpress.XtraEditors.ComboBoxEdit CboDataBits;
        private DevExpress.XtraEditors.ComboBoxEdit CboParity;
        private DevExpress.XtraEditors.ComboBoxEdit CboStopBits;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem8;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem9;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem3;
        private System.IO.Ports.SerialPort serialPort1;
        private DevExpress.XtraLayout.EmptySpaceItem emptySpaceItem9;
    }
}