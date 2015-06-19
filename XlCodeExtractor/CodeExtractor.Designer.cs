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
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.listCodeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lblFilePath = new System.Windows.Forms.Label();
            this.tvListCodes = new System.Windows.Forms.TreeView();
            this.exportCodesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(237, 176);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 1;
            this.btnExport.Text = "=>";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // listViewModulesDestination
            // 
            this.listViewModulesDestination.Location = new System.Drawing.Point(318, 78);
            this.listViewModulesDestination.Name = "listViewModulesDestination";
            this.listViewModulesDestination.Size = new System.Drawing.Size(219, 218);
            this.listViewModulesDestination.TabIndex = 2;
            this.listViewModulesDestination.UseCompatibleStateImageBehavior = false;
            this.listViewModulesDestination.View = System.Windows.Forms.View.List;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(562, 24);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.listCodeToolStripMenuItem,
            this.exportCodesToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(35, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.openToolStripMenuItem.Text = "Open...";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // listCodeToolStripMenuItem
            // 
            this.listCodeToolStripMenuItem.Name = "listCodeToolStripMenuItem";
            this.listCodeToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.listCodeToolStripMenuItem.Text = "List Code";
            this.listCodeToolStripMenuItem.Click += new System.EventHandler(this.listCodeToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
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
            this.tvListCodes.Location = new System.Drawing.Point(12, 78);
            this.tvListCodes.Name = "tvListCodes";
            this.tvListCodes.Size = new System.Drawing.Size(219, 218);
            this.tvListCodes.TabIndex = 5;
            // 
            // exportCodesToolStripMenuItem
            // 
            this.exportCodesToolStripMenuItem.Name = "exportCodesToolStripMenuItem";
            this.exportCodesToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.exportCodesToolStripMenuItem.Text = "Export Codes";
            this.exportCodesToolStripMenuItem.Click += new System.EventHandler(this.exportCodesToolStripMenuItem_Click);
            // 
            // CodeExtractor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(562, 476);
            this.Controls.Add(this.tvListCodes);
            this.Controls.Add(this.lblFilePath);
            this.Controls.Add(this.listViewModulesDestination);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "CodeExtractor";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "XL Code Extractor";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
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
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem listCodeToolStripMenuItem;
        private System.Windows.Forms.TreeView tvListCodes;
        private System.Windows.Forms.ToolStripMenuItem exportCodesToolStripMenuItem;
    }
}

