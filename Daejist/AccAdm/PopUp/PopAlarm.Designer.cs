namespace AccAdm
{
    partial class PopAlarm
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
            DevExpress.XtraEditors.Controls.EditorButtonImageOptions editorButtonImageOptions1 = new DevExpress.XtraEditors.Controls.EditorButtonImageOptions();
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject1 = new DevExpress.Utils.SerializableAppearanceObject();
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject2 = new DevExpress.Utils.SerializableAppearanceObject();
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject3 = new DevExpress.Utils.SerializableAppearanceObject();
            DevExpress.Utils.SerializableAppearanceObject serializableAppearanceObject4 = new DevExpress.Utils.SerializableAppearanceObject();
            DevExpress.Utils.ContextButton contextButton1 = new DevExpress.Utils.ContextButton();
            DevExpress.XtraEditors.TableLayout.TableColumnDefinition tableColumnDefinition1 = new DevExpress.XtraEditors.TableLayout.TableColumnDefinition();
            DevExpress.XtraEditors.TableLayout.TableRowDefinition tableRowDefinition1 = new DevExpress.XtraEditors.TableLayout.TableRowDefinition();
            DevExpress.XtraEditors.TableLayout.TableRowDefinition tableRowDefinition2 = new DevExpress.XtraEditors.TableLayout.TableRowDefinition();
            DevExpress.XtraEditors.TableLayout.TableRowDefinition tableRowDefinition3 = new DevExpress.XtraEditors.TableLayout.TableRowDefinition();
            DevExpress.XtraGrid.Views.Tile.TileViewItemElement tileViewItemElement1 = new DevExpress.XtraGrid.Views.Tile.TileViewItemElement();
            DevExpress.XtraGrid.Views.Tile.TileViewItemElement tileViewItemElement2 = new DevExpress.XtraGrid.Views.Tile.TileViewItemElement();
            DevExpress.XtraGrid.Views.Tile.TileViewItemElement tileViewItemElement3 = new DevExpress.XtraGrid.Views.Tile.TileViewItemElement();
            this.GridColPGMNM = new DevExpress.XtraGrid.Columns.TileViewColumn();
            this.GridColALMSG = new DevExpress.XtraGrid.Columns.TileViewColumn();
            this.GridColCDATE = new DevExpress.XtraGrid.Columns.TileViewColumn();
            this.repositoryItemButtonEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.button1 = new System.Windows.Forms.Button();
            this.GridRetr = new DevExpress.XtraGrid.GridControl();
            this.GridViewTile = new DevExpress.XtraGrid.Views.Tile.TileView();
            this.Root = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemButtonEdit1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.GridRetr)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewTile)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            this.SuspendLayout();
            // 
            // GridColPGMNM
            // 
            this.GridColPGMNM.AppearanceCell.Options.UseTextOptions = true;
            this.GridColPGMNM.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColPGMNM.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColPGMNM.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColPGMNM.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColPGMNM.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColPGMNM.Caption = "프로그램명";
            this.GridColPGMNM.FieldName = "PGMNM";
            this.GridColPGMNM.Name = "GridColPGMNM";
            this.GridColPGMNM.OptionsColumn.AllowEdit = false;
            this.GridColPGMNM.Visible = true;
            this.GridColPGMNM.VisibleIndex = 1;
            // 
            // GridColALMSG
            // 
            this.GridColALMSG.AppearanceCell.Options.UseTextOptions = true;
            this.GridColALMSG.AppearanceCell.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near;
            this.GridColALMSG.AppearanceCell.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColALMSG.AppearanceHeader.Options.UseTextOptions = true;
            this.GridColALMSG.AppearanceHeader.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
            this.GridColALMSG.AppearanceHeader.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
            this.GridColALMSG.Caption = "메세지";
            this.GridColALMSG.FieldName = "ALMSG";
            this.GridColALMSG.Name = "GridColALMSG";
            this.GridColALMSG.OptionsColumn.AllowEdit = false;
            this.GridColALMSG.Visible = true;
            this.GridColALMSG.VisibleIndex = 0;
            // 
            // GridColCDATE
            // 
            this.GridColCDATE.Caption = "알람일시";
            this.GridColCDATE.FieldName = "CDATE";
            this.GridColCDATE.Name = "GridColCDATE";
            this.GridColCDATE.OptionsColumn.AllowEdit = false;
            this.GridColCDATE.Visible = true;
            this.GridColCDATE.VisibleIndex = 2;
            // 
            // repositoryItemButtonEdit1
            // 
            this.repositoryItemButtonEdit1.AutoHeight = false;
            editorButtonImageOptions1.Image = global::AccAdm.Properties.Resources.close;
            this.repositoryItemButtonEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Close, "", -1, true, true, false, editorButtonImageOptions1, new DevExpress.Utils.KeyShortcut(System.Windows.Forms.Keys.None), serializableAppearanceObject1, serializableAppearanceObject2, serializableAppearanceObject3, serializableAppearanceObject4, "", null, null, DevExpress.Utils.ToolTipAnchor.Default)});
            this.repositoryItemButtonEdit1.Name = "repositoryItemButtonEdit1";
            this.repositoryItemButtonEdit1.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.HideTextEditor;
            // 
            // layoutControl1
            // 
            this.layoutControl1.Controls.Add(this.button1);
            this.layoutControl1.Controls.Add(this.GridRetr);
            this.layoutControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.layoutControl1.Location = new System.Drawing.Point(0, 0);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.Root;
            this.layoutControl1.Size = new System.Drawing.Size(400, 659);
            this.layoutControl1.TabIndex = 0;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(50)))), ((int)(((byte)(50)))), ((int)(((byte)(50)))));
            this.button1.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(66)))), ((int)(((byte)(66)))), ((int)(((byte)(66)))));
            this.button1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(66)))), ((int)(((byte)(66)))), ((int)(((byte)(66)))));
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Gray;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(12, 12);
            this.button1.Margin = new System.Windows.Forms.Padding(0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(376, 22);
            this.button1.TabIndex = 5;
            this.button1.Text = "Check all alarms";
            this.button1.UseCompatibleTextRendering = true;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // GridRetr
            // 
            this.GridRetr.Location = new System.Drawing.Point(12, 38);
            this.GridRetr.MainView = this.GridViewTile;
            this.GridRetr.Name = "GridRetr";
            this.GridRetr.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemButtonEdit1});
            this.GridRetr.Size = new System.Drawing.Size(376, 609);
            this.GridRetr.TabIndex = 0;
            this.GridRetr.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.GridViewTile});
            // 
            // GridViewTile
            // 
            this.GridViewTile.Appearance.EmptySpace.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.GridViewTile.Appearance.EmptySpace.Options.UseBackColor = true;
            this.GridViewTile.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.GridColALMSG,
            this.GridColPGMNM,
            this.GridColCDATE});
            contextButton1.AlignmentOptions.Position = DevExpress.Utils.ContextItemPosition.Far;
            contextButton1.Id = new System.Guid("4e067021-169e-4cc6-9eb2-720b0c5b5f8d");
            contextButton1.ImageOptionsCollection.ItemNormal.Image = global::AccAdm.Properties.Resources.close;
            contextButton1.Name = "BtnReadY";
            contextButton1.Visibility = DevExpress.Utils.ContextItemVisibility.Visible;
            this.GridViewTile.ContextButtons.Add(contextButton1);
            this.GridViewTile.GridControl = this.GridRetr;
            this.GridViewTile.Name = "GridViewTile";
            this.GridViewTile.OptionsTiles.GroupTextPadding = new System.Windows.Forms.Padding(4, 8, 12, 8);
            this.GridViewTile.OptionsTiles.IndentBetweenGroups = 0;
            this.GridViewTile.OptionsTiles.IndentBetweenItems = 0;
            this.GridViewTile.OptionsTiles.ItemPadding = new System.Windows.Forms.Padding(6);
            this.GridViewTile.OptionsTiles.ItemSize = new System.Drawing.Size(248, 119);
            this.GridViewTile.OptionsTiles.LayoutMode = DevExpress.XtraGrid.Views.Tile.TileViewLayoutMode.List;
            this.GridViewTile.OptionsTiles.Orientation = System.Windows.Forms.Orientation.Vertical;
            this.GridViewTile.OptionsTiles.Padding = new System.Windows.Forms.Padding(1);
            this.GridViewTile.OptionsTiles.RowCount = 0;
            this.GridViewTile.OptionsTiles.ScrollMode = DevExpress.XtraEditors.TileControlScrollMode.TouchScrollBar;
            tableColumnDefinition1.PaddingLeft = 5;
            tableColumnDefinition1.PaddingRight = 5;
            this.GridViewTile.TileColumns.Add(tableColumnDefinition1);
            tableRowDefinition1.Length.Value = 29D;
            tableRowDefinition1.PaddingBottom = 5;
            tableRowDefinition1.PaddingTop = 5;
            tableRowDefinition2.Length.Value = 51D;
            tableRowDefinition2.PaddingBottom = 5;
            tableRowDefinition2.PaddingTop = 5;
            tableRowDefinition3.Length.Value = 27D;
            tableRowDefinition3.PaddingBottom = 5;
            tableRowDefinition3.PaddingTop = 5;
            this.GridViewTile.TileRows.Add(tableRowDefinition1);
            this.GridViewTile.TileRows.Add(tableRowDefinition2);
            this.GridViewTile.TileRows.Add(tableRowDefinition3);
            tileViewItemElement1.Appearance.Normal.Font = new System.Drawing.Font("맑은 고딕", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            tileViewItemElement1.Appearance.Normal.ForeColor = System.Drawing.Color.White;
            tileViewItemElement1.Appearance.Normal.Options.UseFont = true;
            tileViewItemElement1.Appearance.Normal.Options.UseForeColor = true;
            tileViewItemElement1.Column = this.GridColPGMNM;
            tileViewItemElement1.ImageOptions.ImageAlignment = DevExpress.XtraEditors.TileItemContentAlignment.MiddleCenter;
            tileViewItemElement1.ImageOptions.ImageScaleMode = DevExpress.XtraEditors.TileItemImageScaleMode.ZoomInside;
            tileViewItemElement1.Text = "GridColPGMNM";
            tileViewItemElement1.TextAlignment = DevExpress.XtraEditors.TileItemContentAlignment.MiddleLeft;
            tileViewItemElement2.Appearance.Normal.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            tileViewItemElement2.Appearance.Normal.Options.UseForeColor = true;
            tileViewItemElement2.Column = this.GridColALMSG;
            tileViewItemElement2.ImageOptions.ImageAlignment = DevExpress.XtraEditors.TileItemContentAlignment.MiddleCenter;
            tileViewItemElement2.ImageOptions.ImageScaleMode = DevExpress.XtraEditors.TileItemImageScaleMode.ZoomInside;
            tileViewItemElement2.RowIndex = 1;
            tileViewItemElement2.Text = "GridColALMSG";
            tileViewItemElement2.TextAlignment = DevExpress.XtraEditors.TileItemContentAlignment.MiddleLeft;
            tileViewItemElement3.Appearance.Normal.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            tileViewItemElement3.Appearance.Normal.Options.UseForeColor = true;
            tileViewItemElement3.Column = this.GridColCDATE;
            tileViewItemElement3.ImageOptions.ImageAlignment = DevExpress.XtraEditors.TileItemContentAlignment.MiddleCenter;
            tileViewItemElement3.ImageOptions.ImageScaleMode = DevExpress.XtraEditors.TileItemImageScaleMode.ZoomInside;
            tileViewItemElement3.RowIndex = 2;
            tileViewItemElement3.Text = "GridColCDATE";
            tileViewItemElement3.TextAlignment = DevExpress.XtraEditors.TileItemContentAlignment.BottomRight;
            this.GridViewTile.TileTemplate.Add(tileViewItemElement1);
            this.GridViewTile.TileTemplate.Add(tileViewItemElement2);
            this.GridViewTile.TileTemplate.Add(tileViewItemElement3);
            this.GridViewTile.ItemClick += new DevExpress.XtraGrid.Views.Tile.TileViewItemClickEventHandler(this.GridViewTile_ItemClick);
            this.GridViewTile.ContextButtonClick += new DevExpress.Utils.ContextItemClickEventHandler(this.GridViewTile_ContextButtonClick);
            this.GridViewTile.ItemCustomize += new DevExpress.XtraGrid.Views.Tile.TileViewItemCustomizeEventHandler(this.GridViewTile_ItemCustomize);
            // 
            // Root
            // 
            this.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.Root.GroupBordersVisible = false;
            this.Root.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem3});
            this.Root.Name = "Root";
            this.Root.Size = new System.Drawing.Size(400, 659);
            this.Root.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.GridRetr;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 26);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(380, 613);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.button1;
            this.layoutControlItem3.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(380, 26);
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            this.layoutControlItem3.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            // 
            // PopAlarm
            // 
            this.Appearance.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.Appearance.Options.UseBackColor = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 659);
            this.Controls.Add(this.layoutControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.KeyPreview = true;
            this.Name = "PopAlarm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "알림창";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Form_Load);
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemButtonEdit1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.GridRetr)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GridViewTile)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Root)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup Root;
        private DevExpress.XtraGrid.GridControl GridRetr;
        private DevExpress.XtraGrid.Views.Tile.TileView GridViewTile;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraGrid.Columns.TileViewColumn GridColALMSG;
        private DevExpress.XtraGrid.Columns.TileViewColumn GridColPGMNM;
        private DevExpress.XtraGrid.Columns.TileViewColumn GridColCDATE;
        private DevExpress.XtraEditors.Repository.RepositoryItemButtonEdit repositoryItemButtonEdit1;
        private System.Windows.Forms.Button button1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
    }
}