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
            this.btnExport = new System.Windows.Forms.Button();
            this.listViewModulesDestination = new System.Windows.Forms.ListView();
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
            this.lblFilePath = new System.Windows.Forms.Label();
            this.tvListCodes = new System.Windows.Forms.TreeView();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.lblSource = new System.Windows.Forms.Label();
            this.lblDestination = new System.Windows.Forms.Label();
            this.groupBoxListCodes = new System.Windows.Forms.GroupBox();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(783, 135);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 1;
            this.btnExport.Text = "=>";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // listViewModulesDestination
            // 
            this.listViewModulesDestination.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewModulesDestination.Location = new System.Drawing.Point(0, 0);
            this.listViewModulesDestination.Name = "listViewModulesDestination";
            this.listViewModulesDestination.Size = new System.Drawing.Size(357, 320);
            this.listViewModulesDestination.TabIndex = 2;
            this.listViewModulesDestination.UseCompatibleStateImageBehavior = false;
            this.listViewModulesDestination.View = System.Windows.Forms.View.List;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.fileSourceToolStripMenuItem,
            this.fileDestinationToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1003, 24);
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
            // 
            // lblFilePath
            // 
            this.lblFilePath.AutoSize = true;
            this.lblFilePath.Location = new System.Drawing.Point(0, 28);
            this.lblFilePath.Name = "lblFilePath";
            this.lblFilePath.Size = new System.Drawing.Size(0, 13);
            this.lblFilePath.TabIndex = 4;
            // 
            // tvListCodes
            // 
            this.tvListCodes.CheckBoxes = true;
            this.tvListCodes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tvListCodes.Location = new System.Drawing.Point(0, 0);
            this.tvListCodes.Name = "tvListCodes";
            this.tvListCodes.Size = new System.Drawing.Size(311, 320);
            this.tvListCodes.TabIndex = 5;
            this.tvListCodes.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvListCodes_AfterSelect);
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Info;
            this.textBox1.Location = new System.Drawing.Point(33, 418);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox1.Size = new System.Drawing.Size(524, 249);
            this.textBox1.TabIndex = 6;
            // 
            // lblSource
            // 
            this.lblSource.AutoSize = true;
            this.lblSource.Location = new System.Drawing.Point(12, 40);
            this.lblSource.Name = "lblSource";
            this.lblSource.Size = new System.Drawing.Size(41, 13);
            this.lblSource.TabIndex = 7;
            this.lblSource.Text = "Source";
            // 
            // lblDestination
            // 
            this.lblDestination.AutoSize = true;
            this.lblDestination.Location = new System.Drawing.Point(330, 40);
            this.lblDestination.Name = "lblDestination";
            this.lblDestination.Size = new System.Drawing.Size(60, 13);
            this.lblDestination.TabIndex = 8;
            this.lblDestination.Text = "Destination";
            // 
            // groupBoxListCodes
            // 
            this.groupBoxListCodes.Location = new System.Drawing.Point(610, 447);
            this.groupBoxListCodes.Name = "groupBoxListCodes";
            this.groupBoxListCodes.Size = new System.Drawing.Size(331, 186);
            this.groupBoxListCodes.TabIndex = 9;
            this.groupBoxListCodes.TabStop = false;
            this.groupBoxListCodes.Text = "List Codes";
            // 
            // splitContainer1
            // 
            this.splitContainer1.Location = new System.Drawing.Point(12, 56);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.tvListCodes);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.listViewModulesDestination);
            this.splitContainer1.Size = new System.Drawing.Size(672, 320);
            this.splitContainer1.SplitterDistance = 311;
            this.splitContainer1.TabIndex = 9;
            // 
            // CodeExtractor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1003, 679);
            this.Controls.Add(this.lblDestination);
            this.Controls.Add(this.lblSource);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.groupBoxListCodes);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.lblFilePath);
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
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.ListView listViewModulesDestination;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.Label lblFilePath;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.TreeView tvListCodes;
        private System.Windows.Forms.TextBox textBox1;
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
        private System.Windows.Forms.Label lblDestination;
        private System.Windows.Forms.GroupBox groupBoxListCodes;
        private System.Windows.Forms.SplitContainer splitContainer1;
    }
}

