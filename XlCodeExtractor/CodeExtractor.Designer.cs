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
            this.rTxtSelectedModuleCode = new System.Windows.Forms.RichTextBox();
            this.tvVBASolution = new System.Windows.Forms.TreeView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.SaveCodesDestinationMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveCodesSourceMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnSaveCurrentCode = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            this.panel1.SuspendLayout();
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
            this.menuStrip1.Size = new System.Drawing.Size(1367, 24);
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
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // fileSourceToolStripMenuItem
            // 
            this.fileSourceToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openSourceMenuItem,
            this.toolStripSeparator1,
            this.ListeSourceCodeMenuItem,
            this.ExportSourceCodesMenuItem,
            this.saveCodesSourceMenuItem,
            this.toolStripSeparator2,
            this.closeSourceMenuItem});
            this.fileSourceToolStripMenuItem.Name = "fileSourceToolStripMenuItem";
            this.fileSourceToolStripMenuItem.Size = new System.Drawing.Size(71, 20);
            this.fileSourceToolStripMenuItem.Text = "File Source";
            // 
            // openSourceMenuItem
            // 
            this.openSourceMenuItem.Name = "openSourceMenuItem";
            this.openSourceMenuItem.Size = new System.Drawing.Size(152, 22);
            this.openSourceMenuItem.Text = "Open Source ...";
            this.openSourceMenuItem.Click += new System.EventHandler(this.openSourceMenuItem_Click);
            // 
            // ListeSourceCodeMenuItem
            // 
            this.ListeSourceCodeMenuItem.Name = "ListeSourceCodeMenuItem";
            this.ListeSourceCodeMenuItem.Size = new System.Drawing.Size(152, 22);
            this.ListeSourceCodeMenuItem.Text = "List Code";
            this.ListeSourceCodeMenuItem.Click += new System.EventHandler(this.ListeSourceCodeMenuItem_Click);
            // 
            // ExportSourceCodesMenuItem
            // 
            this.ExportSourceCodesMenuItem.Name = "ExportSourceCodesMenuItem";
            this.ExportSourceCodesMenuItem.Size = new System.Drawing.Size(152, 22);
            this.ExportSourceCodesMenuItem.Text = "Export Codes";
            this.ExportSourceCodesMenuItem.Click += new System.EventHandler(this.ExportSourceCodesMenuItem_Click);
            // 
            // closeSourceMenuItem
            // 
            this.closeSourceMenuItem.Name = "closeSourceMenuItem";
            this.closeSourceMenuItem.Size = new System.Drawing.Size(152, 22);
            this.closeSourceMenuItem.Text = "Close";
            this.closeSourceMenuItem.Click += new System.EventHandler(this.closeSourceMenuItem_Click);
            // 
            // fileDestinationToolStripMenuItem
            // 
            this.fileDestinationToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openDestinationMenuItem,
            this.toolStripSeparator3,
            this.ListDestinationCodeMenuItem,
            this.ExportDestinationCodesMenuItem,
            this.SaveCodesDestinationMenuItem,
            this.toolStripSeparator4,
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
            this.txtSourceCode.Size = new System.Drawing.Size(258, 192);
            this.txtSourceCode.TabIndex = 6;
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.Location = new System.Drawing.Point(5, 9);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(41, 13);
            this.lblSource.TabIndex = 7;
            this.lblSource.Text = "Source";
            // 
            // splitContainer1
            // 
            this.splitContainer1.Location = new System.Drawing.Point(5, 25);
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
            this.splitContainer1.Size = new System.Drawing.Size(522, 352);
            this.splitContainer1.SplitterDistance = 258;
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
            this.txtDestinationCode.Size = new System.Drawing.Size(254, 192);
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
            this.lblSourceFilePath.Location = new System.Drawing.Point(2, 30);
            this.lblSourceFilePath.Name = "lblSourceFilePath";
            this.lblSourceFilePath.Size = new System.Drawing.Size(44, 13);
            this.lblSourceFilePath.TabIndex = 10;
            this.lblSourceFilePath.Text = "Source:";
            // 
            // lblDestinationFilePath
            // 
            this.lblDestinationFilePath.AutoSize = true;
            this.lblDestinationFilePath.Location = new System.Drawing.Point(826, 30);
            this.lblDestinationFilePath.Name = "lblDestinationFilePath";
            this.lblDestinationFilePath.Size = new System.Drawing.Size(63, 13);
            this.lblDestinationFilePath.TabIndex = 11;
            this.lblDestinationFilePath.Text = "Destination:";
            // 
            // lblDestination
            // 
            this.lblDestination.AutoSize = true;
            this.lblDestination.Location = new System.Drawing.Point(270, 9);
            this.lblDestination.Name = "lblDestination";
            this.lblDestination.Size = new System.Drawing.Size(60, 13);
            this.lblDestination.TabIndex = 12;
            this.lblDestination.Text = "Destination";
            // 
            // splitContainer2
            // 
            this.splitContainer2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer2.Location = new System.Drawing.Point(5, 46);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.tvVBASolution);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.rTxtSelectedModuleCode);
            this.splitContainer2.Size = new System.Drawing.Size(818, 619);
            this.splitContainer2.SplitterDistance = 205;
            this.splitContainer2.TabIndex = 13;
            // 
            // rTxtSelectedModuleCode
            // 
            this.rTxtSelectedModuleCode.BackColor = System.Drawing.SystemColors.Info;
            this.rTxtSelectedModuleCode.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rTxtSelectedModuleCode.Location = new System.Drawing.Point(0, 0);
            this.rTxtSelectedModuleCode.Name = "rTxtSelectedModuleCode";
            this.rTxtSelectedModuleCode.Size = new System.Drawing.Size(609, 619);
            this.rTxtSelectedModuleCode.TabIndex = 1;
            this.rTxtSelectedModuleCode.Text = "";
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
            this.tvVBASolution.Size = new System.Drawing.Size(205, 619);
            this.tvVBASolution.TabIndex = 0;
            this.tvVBASolution.DoubleClick += new System.EventHandler(this.tvVBASolution_DoubleClick);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnSaveCurrentCode);
            this.panel1.Controls.Add(this.splitContainer1);
            this.panel1.Controls.Add(this.lblSource);
            this.panel1.Controls.Add(this.lblDestination);
            this.panel1.Location = new System.Drawing.Point(829, 46);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(534, 421);
            this.panel1.TabIndex = 14;
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(149, 6);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(149, 6);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(169, 6);
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(169, 6);
            // 
            // SaveCodesDestinationMenuItem
            // 
            this.SaveCodesDestinationMenuItem.Name = "SaveCodesDestinationMenuItem";
            this.SaveCodesDestinationMenuItem.Size = new System.Drawing.Size(172, 22);
            this.SaveCodesDestinationMenuItem.Text = "Save Codes";
            this.SaveCodesDestinationMenuItem.Click += new System.EventHandler(this.saveCodesDestinationMenuItem_Click);
            // 
            // saveCodesSourceMenuItem
            // 
            this.saveCodesSourceMenuItem.Name = "saveCodesSourceMenuItem";
            this.saveCodesSourceMenuItem.Size = new System.Drawing.Size(152, 22);
            this.saveCodesSourceMenuItem.Text = "Save Codes";
            this.saveCodesSourceMenuItem.Click += new System.EventHandler(this.saveCodesSourceMenuItem_Click);
            // 
            // btnSaveCurrentCode
            // 
            this.btnSaveCurrentCode.Location = new System.Drawing.Point(272, 381);
            this.btnSaveCurrentCode.Name = "btnSaveCurrentCode";
            this.btnSaveCurrentCode.Size = new System.Drawing.Size(234, 23);
            this.btnSaveCurrentCode.TabIndex = 13;
            this.btnSaveCurrentCode.Text = "Save Current Code";
            this.btnSaveCurrentCode.UseVisualStyleBackColor = true;
            // 
            // CodeExtractor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1367, 668);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.splitContainer2);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.lblDestinationFilePath);
            this.Controls.Add(this.lblSourceFilePath);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
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
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
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
        private System.Windows.Forms.RichTextBox rTxtSelectedModuleCode;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripMenuItem SaveCodesDestinationMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripMenuItem saveCodesSourceMenuItem;
        private System.Windows.Forms.Button btnSaveCurrentCode;
    }
}

