namespace XlCodeExtractor
{
    partial class CodeExtractor
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
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("VBASolution");
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fileSourceToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openSourceMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ListeSourceCodeMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ExportSourceCodesMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeSourceMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fileDestinationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openDestinationMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ListDestinationCodeMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ExportDestinationCodesMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeDestinationMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tvSourceListCodes = new System.Windows.Forms.TreeView();
            this.txtSourceCode = new System.Windows.Forms.TextBox();
            this.lblSource = new System.Windows.Forms.Label();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.txtDestinationCode = new System.Windows.Forms.TextBox();
            this.tvDestinationListCodes = new System.Windows.Forms.TreeView();
            this.lblSourceFilePath = new System.Windows.Forms.Label();
            this.lblDestinationFilePath = new System.Windows.Forms.Label();
            this.lblDestination = new System.Windows.Forms.Label();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.tvVBASolution = new System.Windows.Forms.TreeView();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.fileSourceToolStripMenuItem,
            this.fileDestinationToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1322, 24);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(35, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(92, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // fileSourceToolStripMenuItem
            // 
            this.fileSourceToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openSourceMenuItem,
            this.ListeSourceCodeMenuItem,
            this.ExportSourceCodesMenuItem,
            this.closeSourceMenuItem});
            this.fileSourceToolStripMenuItem.Name = "fileSourceToolStripMenuItem";
            this.fileSourceToolStripMenuItem.Size = new System.Drawing.Size(71, 20);
            this.fileSourceToolStripMenuItem.Text = "File Source";
            // 
            // openSourceMenuItem
            // 
            this.openSourceMenuItem.Name = "openSourceMenuItem";
            this.openSourceMenuItem.Size = new System.Drawing.Size(151, 22);
            this.openSourceMenuItem.Text = "Open Source ...";
            this.openSourceMenuItem.Click += new System.EventHandler(this.openSourceMenuItem_Click);
            // 
            // ListeSourceCodeMenuItem
            // 
            this.ListeSourceCodeMenuItem.Name = "ListeSourceCodeMenuItem";
            this.ListeSourceCodeMenuItem.Size = new System.Drawing.Size(151, 22);
            this.ListeSourceCodeMenuItem.Text = "List Code";
            this.ListeSourceCodeMenuItem.Click += new System.EventHandler(this.ListeSourceCodeMenuItem_Click);
            // 
            // ExportSourceCodesMenuItem
            // 
            this.ExportSourceCodesMenuItem.Name = "ExportSourceCodesMenuItem";
            this.ExportSourceCodesMenuItem.Size = new System.Drawing.Size(151, 22);
            this.ExportSourceCodesMenuItem.Text = "Export Codes";
            this.ExportSourceCodesMenuItem.Click += new System.EventHandler(this.ExportSourceCodesMenuItem_Click);
            // 
            // closeSourceMenuItem
            // 
            this.closeSourceMenuItem.Name = "closeSourceMenuItem";
            this.closeSourceMenuItem.Size = new System.Drawing.Size(151, 22);
            this.closeSourceMenuItem.Text = "Close";
            this.closeSourceMenuItem.Click += new System.EventHandler(this.closeSourceMenuItem_Click);
            // 
            // fileDestinationToolStripMenuItem
            // 
            this.fileDestinationToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openDestinationMenuItem,
            this.ListDestinationCodeMenuItem,
            this.ExportDestinationCodesMenuItem,
            this.closeDestinationMenuItem});
            this.fileDestinationToolStripMenuItem.Name = "fileDestinationToolStripMenuItem";
            this.fileDestinationToolStripMenuItem.Size = new System.Drawing.Size(92, 20);
            this.fileDestinationToolStripMenuItem.Text = "File Destination";
            // 
            // openDestinationMenuItem
            // 
            this.openDestinationMenuItem.Name = "openDestinationMenuItem";
            this.openDestinationMenuItem.Size = new System.Drawing.Size(172, 22);
            this.openDestinationMenuItem.Text = "Open Destination ...";
            this.openDestinationMenuItem.Click += new System.EventHandler(this.openDestinationMenuItem_Click);
            // 
            // ListDestinationCodeMenuItem
            // 
            this.ListDestinationCodeMenuItem.Name = "ListDestinationCodeMenuItem";
            this.ListDestinationCodeMenuItem.Size = new System.Drawing.Size(172, 22);
            this.ListDestinationCodeMenuItem.Text = "List Code";
            this.ListDestinationCodeMenuItem.Click += new System.EventHandler(this.ListDestinationMenuItem_Click);
            // 
            // ExportDestinationCodesMenuItem
            // 
            this.ExportDestinationCodesMenuItem.Name = "ExportDestinationCodesMenuItem";
            this.ExportDestinationCodesMenuItem.Size = new System.Drawing.Size(172, 22);
            this.ExportDestinationCodesMenuItem.Text = "Export Codes";
            this.ExportDestinationCodesMenuItem.Click += new System.EventHandler(this.ExportDestinationCodesMenuItem_Click);
            // 
            // closeDestinationMenuItem
            // 
            this.closeDestinationMenuItem.Name = "closeDestinationMenuItem";
            this.closeDestinationMenuItem.Size = new System.Drawing.Size(172, 22);
            this.closeDestinationMenuItem.Text = "Close";
            this.closeDestinationMenuItem.Click += new System.EventHandler(this.closeDestinationMenuItem_Click);
            // 
            // tvSourceListCodes
            // 
            this.tvSourceListCodes.CheckBoxes = true;
            this.tvSourceListCodes.Location = new System.Drawing.Point(0, 0);
            this.tvSourceListCodes.Name = "tvSourceListCodes";
            this.tvSourceListCodes.Size = new System.Drawing.Size(486, 151);
            this.tvSourceListCodes.TabIndex = 5;
            this.tvSourceListCodes.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvListCodes_AfterSelect);
            // 
            // txtSourceCode
            // 
            this.txtSourceCode.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSourceCode.BackColor = System.Drawing.SystemColors.Info;
            this.txtSourceCode.Location = new System.Drawing.Point(0, 157);
            this.txtSourceCode.Multiline = true;
            this.txtSourceCode.Name = "txtSourceCode";
            this.txtSourceCode.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtSourceCode.Size = new System.Drawing.Size(485, 192);
            this.txtSourceCode.TabIndex = 6;
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.Location = new System.Drawing.Point(12, 45);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(41, 13);
            this.lblSource.TabIndex = 7;
            this.lblSource.Text = "Source";
            // 
            // splitContainer1
            // 
            this.splitContainer1.Location = new System.Drawing.Point(12, 61);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.tvSourceListCodes);
            this.splitContainer1.Panel1.Controls.Add(this.txtSourceCode);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.txtDestinationCode);
            this.splitContainer1.Panel2.Controls.Add(this.tvDestinationListCodes);
            this.splitContainer1.Size = new System.Drawing.Size(979, 352);
            this.splitContainer1.SplitterDistance = 485;
            this.splitContainer1.TabIndex = 9;
            // 
            // txtDestinationCode
            // 
            this.txtDestinationCode.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDestinationCode.BackColor = System.Drawing.SystemColors.Info;
            this.txtDestinationCode.Location = new System.Drawing.Point(3, 157);
            this.txtDestinationCode.Multiline = true;
            this.txtDestinationCode.Name = "txtDestinationCode";
            this.txtDestinationCode.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtDestinationCode.Size = new System.Drawing.Size(484, 192);
            this.txtDestinationCode.TabIndex = 12;
            // 
            // tvDestinationListCodes
            // 
            this.tvDestinationListCodes.Location = new System.Drawing.Point(3, 0);
            this.tvDestinationListCodes.Name = "tvDestinationListCodes";
            this.tvDestinationListCodes.Size = new System.Drawing.Size(484, 151);
            this.tvDestinationListCodes.TabIndex = 10;
            // 
            // lblSourceFilePath
            // 
            this.lblSourceFilePath.AutoSize = true;
            this.lblSourceFilePath.Location = new System.Drawing.Point(12, 416);
            this.lblSourceFilePath.Name = "lblSourceFilePath";
            this.lblSourceFilePath.Size = new System.Drawing.Size(44, 13);
            this.lblSourceFilePath.TabIndex = 10;
            this.lblSourceFilePath.Text = "Source:";
            // 
            // lblDestinationFilePath
            // 
            this.lblDestinationFilePath.AutoSize = true;
            this.lblDestinationFilePath.Location = new System.Drawing.Point(501, 416);
            this.lblDestinationFilePath.Name = "lblDestinationFilePath";
            this.lblDestinationFilePath.Size = new System.Drawing.Size(63, 13);
            this.lblDestinationFilePath.TabIndex = 11;
            this.lblDestinationFilePath.Text = "Destination:";
            // 
            // lblDestination
            // 
            this.lblDestination.AutoSize = true;
            this.lblDestination.Location = new System.Drawing.Point(501, 45);
            this.lblDestination.Name = "lblDestination";
            this.lblDestination.Size = new System.Drawing.Size(60, 13);
            this.lblDestination.TabIndex = 12;
            this.lblDestination.Text = "Destination";
            // 
            // splitContainer2
            // 
            this.splitContainer2.Location = new System.Drawing.Point(1019, 27);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.tvVBASolution);
            this.splitContainer2.Size = new System.Drawing.Size(291, 386);
            this.splitContainer2.SplitterDistance = 137;
            this.splitContainer2.TabIndex = 13;
            // 
            // tvVBASolution
            // 
            this.tvVBASolution.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tvVBASolution.Location = new System.Drawing.Point(0, 0);
            this.tvVBASolution.Name = "tvVBASolution";
            treeNode1.Name = "nRootVBASolution";
            treeNode1.Text = "VBASolution";
            this.tvVBASolution.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1});
            this.tvVBASolution.Size = new System.Drawing.Size(150, 386);
            this.tvVBASolution.TabIndex = 0;
            // 
            // CodeExtractor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1322, 469);
            this.Controls.Add(this.splitContainer2);
            this.Controls.Add(this.lblDestination);
            this.Controls.Add(this.lblDestinationFilePath);
            this.Controls.Add(this.lblSourceFilePath);
            this.Controls.Add(this.lblSource);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "CodeExtractor";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "XL Code Extractor";
            this.Load += new System.EventHandler(this.CodeExtractor_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.TreeView tvSourceListCodes;
        private System.Windows.Forms.TextBox txtSourceCode;
        private System.Windows.Forms.ToolStripMenuItem fileSourceToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openSourceMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ListeSourceCodeMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ExportSourceCodesMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fileDestinationToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openDestinationMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ListDestinationCodeMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ExportDestinationCodesMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeSourceMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeDestinationMenuItem;
        private System.Windows.Forms.Label lblSource;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TreeView tvDestinationListCodes;
        private System.Windows.Forms.Label lblSourceFilePath;
        private System.Windows.Forms.Label lblDestinationFilePath;
        private System.Windows.Forms.TextBox txtDestinationCode;
        private System.Windows.Forms.Label lblDestination;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.TreeView tvVBASolution;
    }
}

