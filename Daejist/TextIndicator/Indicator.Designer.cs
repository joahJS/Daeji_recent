namespace TextIndicator
{
    partial class Indicator
    {
        /// <summary> 
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary> 
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마십시오.
        /// </summary>
        private void InitializeComponent()
        {
            this.pictureBoxStable = new System.Windows.Forms.PictureBox();
            this.pictureBoxConnect = new System.Windows.Forms.PictureBox();
            this.labelWeight = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStable)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxConnect)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBoxStable
            // 
            this.pictureBoxStable.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxStable.BackgroundImage = global::TextIndicator.Properties.Resources.steady_false;
            this.pictureBoxStable.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBoxStable.Location = new System.Drawing.Point(3, 72);
            this.pictureBoxStable.Name = "pictureBoxStable";
            this.pictureBoxStable.Size = new System.Drawing.Size(46, 50);
            this.pictureBoxStable.TabIndex = 4;
            this.pictureBoxStable.TabStop = false;
            // 
            // pictureBoxConnect
            // 
            this.pictureBoxConnect.BackColor = System.Drawing.Color.Transparent;
            this.pictureBoxConnect.BackgroundImage = global::TextIndicator.Properties.Resources.connect_false;
            this.pictureBoxConnect.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBoxConnect.Location = new System.Drawing.Point(3, 3);
            this.pictureBoxConnect.Name = "pictureBoxConnect";
            this.pictureBoxConnect.Size = new System.Drawing.Size(46, 50);
            this.pictureBoxConnect.TabIndex = 3;
            this.pictureBoxConnect.TabStop = false;
            // 
            // labelWeight
            // 
            this.labelWeight.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelWeight.BackColor = System.Drawing.Color.Transparent;
            this.labelWeight.Font = new System.Drawing.Font("HY견고딕", 27.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.labelWeight.ForeColor = System.Drawing.Color.White;
            this.labelWeight.Location = new System.Drawing.Point(56, 3);
            this.labelWeight.Name = "labelWeight";
            this.labelWeight.Size = new System.Drawing.Size(375, 117);
            this.labelWeight.TabIndex = 5;
            this.labelWeight.Text = "0kg";
            this.labelWeight.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Indicator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.Controls.Add(this.labelWeight);
            this.Controls.Add(this.pictureBoxStable);
            this.Controls.Add(this.pictureBoxConnect);
            this.MinimumSize = new System.Drawing.Size(229, 100);
            this.Name = "Indicator";
            this.Size = new System.Drawing.Size(457, 125);
            this.Load += new System.EventHandler(this.Indicator_Load);
            this.SizeChanged += new System.EventHandler(this.Indicator_SizeChanged);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxStable)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxConnect)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBoxStable;
        private System.Windows.Forms.PictureBox pictureBoxConnect;
        private System.Windows.Forms.Label labelWeight;
    }
}
