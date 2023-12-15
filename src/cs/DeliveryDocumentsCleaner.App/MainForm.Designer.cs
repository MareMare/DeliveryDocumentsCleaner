namespace DeliveryDocumentsCleaner.App
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.statusStrip1 = new StatusStrip();
            this.toolStripStatusLabel1 = new ToolStripStatusLabel();
            this.toolStripProgressBar1 = new ToolStripProgressBar();
            this.panelOfFolder = new Panel();
            this.textBoxOfFolder = new TextBox();
            this.buttonToSelectFolder = new Button();
            this.label1 = new Label();
            this.panelOfFiles = new Panel();
            this.listBoxOfFiles = new ListBox();
            this.panelOfFilesOperation = new Panel();
            this.buttonToMoveDown = new Button();
            this.buttonToMoveUp = new Button();
            this.tabControl1 = new TabControl();
            this.tabPage1 = new TabPage();
            this.buttonToExecuteClear = new Button();
            this.groupBox1 = new GroupBox();
            this.xlColorPanel1 = new XlColorPanel();
            this.groupBoxOfWdColor = new GroupBox();
            this.wdColorPanel1 = new WdColorPanel();
            this.tabPage2 = new TabPage();
            this.tableLayoutPanel1 = new TableLayoutPanel();
            this.statusStrip1.SuspendLayout();
            this.panelOfFolder.SuspendLayout();
            this.panelOfFiles.SuspendLayout();
            this.panelOfFilesOperation.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBoxOfWdColor.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new ToolStripItem[] { this.toolStripStatusLabel1, this.toolStripProgressBar1 });
            this.statusStrip1.Location = new Point(5, 457);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new Size(643, 22);
            this.statusStrip1.TabIndex = 0;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new Size(526, 17);
            this.toolStripStatusLabel1.Spring = true;
            this.toolStripStatusLabel1.Text = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Alignment = ToolStripItemAlignment.Right;
            this.toolStripProgressBar1.AutoSize = false;
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Overflow = ToolStripItemOverflow.Never;
            this.toolStripProgressBar1.Size = new Size(100, 16);
            this.toolStripProgressBar1.Step = 5;
            this.toolStripProgressBar1.Style = ProgressBarStyle.Continuous;
            // 
            // panelOfFolder
            // 
            this.panelOfFolder.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.panelOfFolder.Controls.Add(this.textBoxOfFolder);
            this.panelOfFolder.Controls.Add(this.buttonToSelectFolder);
            this.panelOfFolder.Controls.Add(this.label1);
            this.panelOfFolder.Location = new Point(3, 3);
            this.panelOfFolder.Name = "panelOfFolder";
            this.panelOfFolder.Size = new Size(381, 46);
            this.panelOfFolder.TabIndex = 1;
            // 
            // textBoxOfFolder
            // 
            this.textBoxOfFolder.BackColor = SystemColors.Info;
            this.textBoxOfFolder.Dock = DockStyle.Fill;
            this.textBoxOfFolder.Location = new Point(0, 23);
            this.textBoxOfFolder.Name = "textBoxOfFolder";
            this.textBoxOfFolder.ReadOnly = true;
            this.textBoxOfFolder.Size = new Size(356, 23);
            this.textBoxOfFolder.TabIndex = 0;
            // 
            // buttonToSelectFolder
            // 
            this.buttonToSelectFolder.BackColor = Color.Bisque;
            this.buttonToSelectFolder.Dock = DockStyle.Right;
            this.buttonToSelectFolder.Location = new Point(356, 23);
            this.buttonToSelectFolder.Name = "buttonToSelectFolder";
            this.buttonToSelectFolder.Size = new Size(25, 23);
            this.buttonToSelectFolder.TabIndex = 1;
            this.buttonToSelectFolder.Text = "...";
            this.buttonToSelectFolder.UseVisualStyleBackColor = false;
            this.buttonToSelectFolder.Click += this.ButtonToSelectFolder_Click;
            // 
            // label1
            // 
            this.label1.Dock = DockStyle.Top;
            this.label1.Location = new Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new Size(381, 23);
            this.label1.TabIndex = 2;
            this.label1.Text = "フォルダの選択";
            this.label1.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // panelOfFiles
            // 
            this.panelOfFiles.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.panelOfFiles.Controls.Add(this.listBoxOfFiles);
            this.panelOfFiles.Controls.Add(this.panelOfFilesOperation);
            this.panelOfFiles.Location = new Point(3, 55);
            this.panelOfFiles.Name = "panelOfFiles";
            this.panelOfFiles.Size = new Size(381, 344);
            this.panelOfFiles.TabIndex = 2;
            // 
            // listBoxOfFiles
            // 
            this.listBoxOfFiles.Dock = DockStyle.Fill;
            this.listBoxOfFiles.FormattingEnabled = true;
            this.listBoxOfFiles.ItemHeight = 15;
            this.listBoxOfFiles.Location = new Point(0, 0);
            this.listBoxOfFiles.Name = "listBoxOfFiles";
            this.listBoxOfFiles.SelectionMode = SelectionMode.MultiSimple;
            this.listBoxOfFiles.Size = new Size(356, 344);
            this.listBoxOfFiles.TabIndex = 0;
            // 
            // panelOfFilesOperation
            // 
            this.panelOfFilesOperation.Controls.Add(this.buttonToMoveDown);
            this.panelOfFilesOperation.Controls.Add(this.buttonToMoveUp);
            this.panelOfFilesOperation.Dock = DockStyle.Right;
            this.panelOfFilesOperation.Location = new Point(356, 0);
            this.panelOfFilesOperation.Name = "panelOfFilesOperation";
            this.panelOfFilesOperation.Size = new Size(25, 344);
            this.panelOfFilesOperation.TabIndex = 3;
            // 
            // buttonToMoveDown
            // 
            this.buttonToMoveDown.Location = new Point(0, 161);
            this.buttonToMoveDown.Name = "buttonToMoveDown";
            this.buttonToMoveDown.Size = new Size(25, 23);
            this.buttonToMoveDown.TabIndex = 3;
            this.buttonToMoveDown.Text = "↓";
            this.buttonToMoveDown.UseVisualStyleBackColor = true;
            // 
            // buttonToMoveUp
            // 
            this.buttonToMoveUp.Location = new Point(0, 132);
            this.buttonToMoveUp.Name = "buttonToMoveUp";
            this.buttonToMoveUp.Size = new Size(25, 23);
            this.buttonToMoveUp.TabIndex = 2;
            this.buttonToMoveUp.Text = "↑";
            this.buttonToMoveUp.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new Point(390, 3);
            this.tabControl1.Name = "tabControl1";
            this.tableLayoutPanel1.SetRowSpan(this.tabControl1, 2);
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new Size(250, 396);
            this.tabControl1.TabIndex = 3;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.buttonToExecuteClear);
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.groupBoxOfWdColor);
            this.tabPage1.Location = new Point(4, 24);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new Padding(3);
            this.tabPage1.Size = new Size(242, 368);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "tabPage1";
            // 
            // buttonToExecuteClear
            // 
            this.buttonToExecuteClear.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this.buttonToExecuteClear.BackColor = Color.Bisque;
            this.buttonToExecuteClear.Location = new Point(161, 322);
            this.buttonToExecuteClear.Name = "buttonToExecuteClear";
            this.buttonToExecuteClear.Size = new Size(75, 40);
            this.buttonToExecuteClear.TabIndex = 2;
            this.buttonToExecuteClear.Text = "クリア実行";
            this.buttonToExecuteClear.UseVisualStyleBackColor = false;
            this.buttonToExecuteClear.Click += this.ButtonToExecuteClear_Click;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.xlColorPanel1);
            this.groupBox1.Dock = DockStyle.Top;
            this.groupBox1.Location = new Point(3, 180);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new Size(236, 72);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "塗りつぶしの色（Excelのみ）";
            // 
            // xlColorPanel1
            // 
            this.xlColorPanel1.BackColor = Color.White;
            this.xlColorPanel1.BorderStyle = BorderStyle.FixedSingle;
            this.xlColorPanel1.Dock = DockStyle.Fill;
            this.xlColorPanel1.Location = new Point(3, 19);
            this.xlColorPanel1.Name = "xlColorPanel1";
            this.xlColorPanel1.Size = new Size(230, 50);
            this.xlColorPanel1.TabIndex = 0;
            // 
            // groupBoxOfWdColor
            // 
            this.groupBoxOfWdColor.Controls.Add(this.wdColorPanel1);
            this.groupBoxOfWdColor.Dock = DockStyle.Top;
            this.groupBoxOfWdColor.Location = new Point(3, 3);
            this.groupBoxOfWdColor.Name = "groupBoxOfWdColor";
            this.groupBoxOfWdColor.Size = new Size(236, 177);
            this.groupBoxOfWdColor.TabIndex = 0;
            this.groupBoxOfWdColor.TabStop = false;
            this.groupBoxOfWdColor.Text = "蛍光ペン（Wordのみ）";
            // 
            // wdColorPanel1
            // 
            this.wdColorPanel1.Dock = DockStyle.Fill;
            this.wdColorPanel1.Location = new Point(3, 19);
            this.wdColorPanel1.Name = "wdColorPanel1";
            this.wdColorPanel1.Size = new Size(230, 155);
            this.wdColorPanel1.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new Point(4, 24);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new Padding(3);
            this.tabPage2.Size = new Size(242, 368);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 256F));
            this.tableLayoutPanel1.Controls.Add(this.panelOfFolder, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.panelOfFiles, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.tabControl1, 1, 0);
            this.tableLayoutPanel1.Dock = DockStyle.Fill;
            this.tableLayoutPanel1.Location = new Point(5, 5);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            this.tableLayoutPanel1.Size = new Size(643, 452);
            this.tableLayoutPanel1.TabIndex = 3;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new SizeF(7F, 15F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new Size(653, 484);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.statusStrip1);
            this.Name = "MainForm";
            this.Padding = new Padding(5);
            this.Text = "納品ドキュメントの調整";
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.panelOfFolder.ResumeLayout(false);
            this.panelOfFolder.PerformLayout();
            this.panelOfFiles.ResumeLayout(false);
            this.panelOfFilesOperation.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBoxOfWdColor.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private StatusStrip statusStrip1;
        private ToolStripStatusLabel toolStripStatusLabel1;
        private ToolStripProgressBar toolStripProgressBar1;
        private Panel panelOfFolder;
        private TextBox textBoxOfFolder;
        private Button buttonToSelectFolder;
        private Label label1;
        private Panel panelOfFiles;
        private ListBox listBoxOfFiles;
        private Panel panelOfFilesOperation;
        private Button buttonToMoveDown;
        private Button buttonToMoveUp;
        private TabControl tabControl1;
        private TabPage tabPage2;
        private TabPage tabPage1;
        private TableLayoutPanel tableLayoutPanel1;
        private GroupBox groupBoxOfWdColor;
        private GroupBox groupBox1;
        private XlColorPanel xlColorPanel1;
        private WdColorPanel wdColorPanel1;
        private Button buttonToExecuteClear;
    }
}
