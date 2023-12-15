namespace DeliveryDocumentsCleaner.App
{
    partial class XlColorPanel
    {
        /// <summary> 
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        #region コンポーネント デザイナーで生成されたコード

        /// <summary> 
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を 
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.panelOfSelectedXlColor = new System.Windows.Forms.Panel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.labelOfSelectedXlColor = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonToSelectXlColor = new System.Windows.Forms.Button();
            this.flowLayoutPanel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelOfSelectedXlColor
            // 
            this.panelOfSelectedXlColor.BackColor = System.Drawing.SystemColors.Control;
            this.panelOfSelectedXlColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelOfSelectedXlColor.Location = new System.Drawing.Point(3, 3);
            this.panelOfSelectedXlColor.Name = "panelOfSelectedXlColor";
            this.panelOfSelectedXlColor.Size = new System.Drawing.Size(40, 40);
            this.panelOfSelectedXlColor.TabIndex = 0;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.AutoScroll = true;
            this.flowLayoutPanel1.Controls.Add(this.panelOfSelectedXlColor);
            this.flowLayoutPanel1.Controls.Add(this.panel2);
            this.flowLayoutPanel1.Controls.Add(this.buttonToSelectXlColor);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(207, 47);
            this.flowLayoutPanel1.TabIndex = 1;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.labelOfSelectedXlColor);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Location = new System.Drawing.Point(49, 3);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(120, 40);
            this.panel2.TabIndex = 1;
            // 
            // labelOfSelectedXlColor
            // 
            this.labelOfSelectedXlColor.Dock = System.Windows.Forms.DockStyle.Fill;
            this.labelOfSelectedXlColor.Location = new System.Drawing.Point(0, 20);
            this.labelOfSelectedXlColor.Name = "labelOfSelectedXlColor";
            this.labelOfSelectedXlColor.Size = new System.Drawing.Size(120, 20);
            this.labelOfSelectedXlColor.TabIndex = 1;
            this.labelOfSelectedXlColor.Text = "---";
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Top;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "選択した塗りつぶし色";
            // 
            // buttonToSelectXlColor
            // 
            this.buttonToSelectXlColor.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.buttonToSelectXlColor.BackColor = System.Drawing.Color.Bisque;
            this.buttonToSelectXlColor.Location = new System.Drawing.Point(175, 11);
            this.buttonToSelectXlColor.Name = "buttonToSelectXlColor";
            this.buttonToSelectXlColor.Size = new System.Drawing.Size(25, 23);
            this.buttonToSelectXlColor.TabIndex = 2;
            this.buttonToSelectXlColor.Text = "...";
            this.buttonToSelectXlColor.UseVisualStyleBackColor = false;
            this.buttonToSelectXlColor.Click += new System.EventHandler(this.ButtonToSelectXlColor_Click);
            // 
            // XlColorPanel
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.BackColor = System.Drawing.Color.White;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.flowLayoutPanel1);
            this.Name = "XlColorPanel";
            this.Size = new System.Drawing.Size(207, 47);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private Panel panelOfSelectedXlColor;
        private FlowLayoutPanel flowLayoutPanel1;
        private Panel panel2;
        private Label labelOfSelectedXlColor;
        private Label label1;
        private Button buttonToSelectXlColor;
    }
}
