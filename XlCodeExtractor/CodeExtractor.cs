using System;
using System.Windows.Forms;
using System.Threading;

using System.Reflection;
using System.Diagnostics;
//using InteropExcel = Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
//using ToolsExcel = Microsoft.Office.Tools.Excel ;
using Microsoft.Vbe.Interop;
using System.IO;
using IniFile;
using XlCodeExtractor.Xl;

namespace XlCodeExtractor
{
    enum WorkbookType
    {
        Source=1,
        Destination
    };

    public partial class CodeExtractor : Form
    {
        Excel.Application ExcelApp = null;
        Excel.Workbooks oWorkbooks;

        private Excel.Workbook oActiveWorkbook { set; get; } // = null;
        Int32 ThisWorkbookCount = 0;
        String excelFileName = string.Empty;// @"S:\PRODUCTS\hyperclose\hypercloseV Bugs\JTY hypercloseV 2015-06-11 Period Marche Pas.xlsb";//String.Empty;
        string localAppPath = String.Empty;

        //Excel Automation variables:
        //Excel.Application xlApp;

        //Excel.Worksheet xlSheet1, xlSheet2, xlSheet3;
        
        //Excel event delegate variables:
        Excel.AppEvents_WorkbookBeforeCloseEventHandler EventDel_BeforeBookClose;
        //Excel.DocEvents_ChangeEventHandler EventDel_CellsChange;

        //Excel.AppEvents_WorkbookOpenEventHandler Events_WorkbookOpenEventHandler;


        XlSource oXlSource = new XlSource();
        XlDestination oXlDestination = new XlDestination();



        public enum MsoAutomationSecurity
        {
            msoAutomationSecurityLow = 1,
            msoAutomationSecurityByUI = 2,
            msoAutomationSecurityForceDisable = 3,
        }

        public CodeExtractor()
        {
            InitializeComponent();

            localAppPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            IniProfile ini = new IniProfile(string.Format("{0}testIni.ini", localAppPath));
            ini.IniWriteValue("Password", "ABC", "123456");
            ini.IniWriteValue("Password", "ABCapp@micronium", "B?[L0gTgV");

            //ini.Write("cfoapp@soltimum", "BfWLK?[0gTgc");

            ListeSourceCodeMenuItem.Enabled = false;
            ExportSourceCodesMenuItem.Enabled = false;
            closeSourceMenuItem.Enabled = false;
            saveCodesSourceMenuItem.Enabled = false;

            ListDestinationCodeMenuItem.Enabled = false;
            ExportDestinationCodesMenuItem.Enabled = false;
            closeDestinationMenuItem.Enabled = false;
            SaveCodesDestinationMenuItem.Enabled = false;

            ExcelApp = new Excel.Application() { Visible = false, DisplayAlerts = false, EnableEvents = false };



            //xlApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(app_WorkbookOpen);
            ExcelApp.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(app_WorkbookActivate);

            //Add an event handler for the WorkbookBeforeClose Event of the
            //Application object.
            EventDel_BeforeBookClose = new Excel.AppEvents_WorkbookBeforeCloseEventHandler(BeforeBookClose);
            ExcelApp.WorkbookBeforeClose += EventDel_BeforeBookClose;

            //Add an event handler for the Change event of both worksheet objects.
            //EventDel_CellsChange = new Excel.DocEvents_ChangeEventHandler(CellsChange);

            //xlSheet1.Change += EventDel_CellsChange;
            //xlSheet2.Change += EventDel_CellsChange;
            //xlSheet3.Change += EventDel_CellsChange;

            //Make Excel visible and give the user control.
            //ExcelApp.Visible = true;
            ExcelApp.UserControl = true;

        }
        private void BeforeBookClose(Excel.Workbook Wb, ref bool Cancel)
        {
            //This is called when you choose to close the workbook in Excel.
            //The event handlers are removed, and then the workbook is closed 
            //without saving the changes.
            Wb.Saved = true;
            Debug.WriteLine("Delegate: Closing the workbook and removing event handlers.");
            //xlSheet1.Change -= EventDel_CellsChange;
            //xlSheet2.Change -= EventDel_CellsChange;
            //xlSheet3.Change -= EventDel_CellsChange;
            ExcelApp.WorkbookBeforeClose -= EventDel_BeforeBookClose;
        }
        private void CellsChange(Excel.Range Target)
        {
            //This is called when any cell on a worksheet is changed.
            Debug.WriteLine("Delegate: You Changed Cells " +
               Target.get_Address(Missing.Value, Missing.Value,
               Excel.XlReferenceStyle.xlA1, Missing.Value, Missing.Value) +
               " on " + Target.Worksheet.Name);
        }

        void app_WorkbookActivate(Excel.Workbook Wb)
        {
            MessageBox.Show(Wb.FullName);
        }

        void app_WorkbookOpen(Excel.Workbook Wb)
        {
            MessageBox.Show(Wb.FullName);
        }


        private Excel.Workbook GetActiveWorkbook()
        {
            oWorkbooks = this.ExcelApp.Workbooks;
            localAppPath = Path.Combine(localAppPath, Path.GetFileNameWithoutExtension(@excelFileName));
            if (Directory.Exists(path: localAppPath))
            { Directory.Delete(path: localAppPath, recursive: true); }
            Directory.CreateDirectory(path: localAppPath);
            return oWorkbooks.Open(@excelFileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        }


        private void btnExport_Click(object sender, EventArgs e)
        {
            //@excelFileName = ExcelApp.GetOpenFilename(FileFilter: "Excel File |*.xlsb;*,xlsb", Title: "Open Excel File to check!");
            try
            {
                ThisWorkbookCount = 0;
            }
            catch (Exception Ex) { MessageBox.Show(String.Format("{0}\r\n{1}\r\n\r\n{2}\r\n\r\n{3}", Ex.Message, Ex.Source, Ex.StackTrace, "Please try launching again!"), "Loading Workbook Codes"); }
        }

        private void ViewCode(WorkbookType type, string codes)
        {
            switch(type)
            {
                case WorkbookType.Source:
                    txtSourceCode.Text = codes;
                    rTxtSelectedModuleCode.Text = codes;
                    break;
                case WorkbookType.Destination:
                    txtDestinationCode.Text = codes;
                    break;
            }
        }

        private void ClearViewCode(WorkbookType type)
        {
            switch (type)
            {
                case WorkbookType.Source:
                    txtSourceCode.Text = string.Empty;
                    rTxtSelectedModuleCode.Text = string.Empty;
                    tvSourceListCodes.Nodes.Clear();
                    tvVBASolution.Nodes.Clear();
                    break;
                case WorkbookType.Destination:
                    txtDestinationCode.Text = string.Empty;
                    tvDestinationListCodes.Nodes.Clear();
                    break;
            }
        }

        private void DesableMenu(WorkbookType type)
        {
            switch (type)
            {
                case WorkbookType.Source:
                    ListeSourceCodeMenuItem.Enabled = false;
                    ExportSourceCodesMenuItem.Enabled = false;
                    closeSourceMenuItem.Enabled = false;
                    break;
                case WorkbookType.Destination:
                    ListDestinationCodeMenuItem.Enabled = false;
                    ExportDestinationCodesMenuItem.Enabled = false;
                    closeDestinationMenuItem.Enabled = false;
                    break;
            }
        }

        private void LoadWorkbookCodes(Excel.Workbook Wb, XlCodeExtractor.WorkbookType WbType, Boolean exportCode = false) //
        {
            oActiveWorkbook = Wb;
            ExcelApp.Visible = exportCode;
            oActiveWorkbook.CheckCompatibility = false;

            Thread.Sleep(TimeSpan.FromMilliseconds(1000));
            Boolean AreVBProjectProtected = (((dynamic)oActiveWorkbook.VBProject).Protection == 1);
            
            try
            {
                if (AreVBProjectProtected)
                {
                    //string message = "Unprotecting Workbook project.";
                    ExcelApp.Visible = true;
                    ExcelApp.WindowState = Excel.XlWindowState.xlMinimized;
                    System.Threading.Thread.Sleep(1000);
                    ExcelApp.Visible = false;
                }
            }
            catch (Exception Ex) { MessageBox.Show(String.Format("{0}\r\n{1}\r\n\r\n{2}\r\n\r\n{3}", Ex.Message, Ex.Source, Ex.StackTrace, "Please try launching again!"), "Loading Workbook Codes"); }

            VBComponents oModules = oActiveWorkbook.VBProject.VBComponents;
            CodeModule oCodeModule = null;
            if (oActiveWorkbook.VBProject.Name.StartsWith("VBAProject")) { oActiveWorkbook.VBProject.Name = oActiveWorkbook.VBProject.Name.Replace("VBAProject", "HypercloseExcelWorkbook"); }
            //String[] ressKey;
            String CodeName = String.Empty;
            // Resource Ref.
            String FilePath = String.Empty;
            String FileCode = String.Empty;
            String FileName = String.Empty;
            Int32 CountOfLines = 0;
            TreeNode tnModule = new TreeNode(text: "Modules");
            TreeNode tnExcelObjects = new TreeNode(text: "Microsoft_Excel_Objects");
            TreeNode tnForms = new TreeNode(text: "Forms");
            TreeNode tnClassModules = new TreeNode(text: "Class_Modules");


            if (WbType.CompareTo(WorkbookType.Source) == 0)
            {
                oXlSource.oModules = oModules;
                tvVBASolution.Nodes.Clear();
                tvVBASolution.Nodes.Add(tnModule);
                tvVBASolution.Nodes.Add(tnExcelObjects);
                tvVBASolution.Nodes.Add(tnForms);
                tvVBASolution.Nodes.Add(tnClassModules);

                tvSourceListCodes.Nodes.Clear();
                tvSourceListCodes.Nodes.Add(tnModule.Clone() as TreeNode );
                tvSourceListCodes.Nodes.Add(tnExcelObjects.Clone() as TreeNode);
                tvSourceListCodes.Nodes.Add(tnForms.Clone() as TreeNode);
                tvSourceListCodes.Nodes.Add(tnClassModules.Clone() as TreeNode);
            }
            else
            {
                oXlDestination.oModules = oModules;

                tvDestinationListCodes.Nodes.Clear();
                tvDestinationListCodes.Nodes.Add(tnModule.Clone() as TreeNode);
                tvDestinationListCodes.Nodes.Add(tnExcelObjects.Clone() as TreeNode);
                tvDestinationListCodes.Nodes.Add(tnForms.Clone() as TreeNode);
                tvDestinationListCodes.Nodes.Add(tnClassModules.Clone() as TreeNode);
            }


            try
            {

                foreach (VBComponent module in oModules)
                {
                    //listViewModulesSource.Items.Add(module.Name);
                    switch (module.Type)
                    {
                        case vbext_ComponentType.vbext_ct_ClassModule:
                            
                            if (!AreWorksheetExist(module.Name) )
                            {
                                //lblMessage.Text = String.Format("Removing Module '{0}'", module.Name);
                                //lblMessage.Refresh();
                                //oModules.Remove(VBComponent: module);
                                oCodeModule = module.CodeModule;
                                CountOfLines = oCodeModule.CountOfLines;
                                if (CountOfLines > 0)
                                {
                                    tnClassModules.Nodes.Add(module.Name);
                                    FileCode = oCodeModule.get_Lines(StartLine: 1, Count: CountOfLines); //oCodeModule.DeleteLines(StartLine: 1, Count: oCodeModule.CountOfLines);
                                    if (exportCode)
                                    {
                                        FilePath = "Class_Modules";
                                        FilePath = Path.Combine(localAppPath, FilePath);
                                        if (!Directory.Exists(path: FilePath))
                                        { Directory.CreateDirectory(path: FilePath); }
                                        module.Export(FileName: String.Concat(Path.Combine(FilePath, module.Name), ".cls"));
                                        CodeExport(filePath: FilePath, fileName: module.Name, fileExt: ".cls", text: FileCode);
                                    }
                                    /*
                                    Ligne 05: Attribute VB_Name = "CFOConstant"
                                    Ligne 11: Const CURRENT_VERSION As String = "10.315.04.13388"
                                    */
                                    Int32 CountOfDeclarationLines = module.CodeModule.CountOfDeclarationLines;
                                    string CFOConstantVersion = String.Empty;
                                    if (String.Equals(module.Name, "CFOConstant"))
                                    {
                                        for (int i = 1; i < module.CodeModule.CountOfDeclarationLines; i++)
                                        {
                                            CFOConstantVersion = module.CodeModule.get_Lines(StartLine: i, Count: 1).Split('=')[0].Replace("\"", "").Replace(".", "_").Trim();
                                            if (String.Equals(CFOConstantVersion, "Const CURRENT_VERSION As String"))
                                            {
                                                CFOConstantVersion = module.CodeModule.get_Lines(StartLine: i, Count: 1).Split('=')[1].Replace("\"", "").Replace(".", "_").Trim();
                                                Directory.CreateDirectory(Path.Combine(localAppPath, CFOConstantVersion));
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            break;
                        case vbext_ComponentType.vbext_ct_Document: //"microsoft_excel_objects":
                            
                            if (module.Name == "ThisWorkbook")
                            { ThisWorkbookCount++; }
                            if (ThisWorkbookCount > 1)
                            {
                                MessageBox.Show(text: "Workbook Corrupted!", caption: "Workbook Consistancy", buttons: MessageBoxButtons.OK, icon: MessageBoxIcon.Warning);
                                return;
                            }

                            //lblMessage.Text = "Removing Module Codes from 'ThisWorkbook'";
                            //lblMessage.Refresh();
                            oCodeModule = module.CodeModule;
                            CountOfLines = oCodeModule.CountOfLines;
                            if (CountOfLines > 0 )
                            {
                                tnExcelObjects.Nodes.Add(module.Name);
                                FileCode = oCodeModule.get_Lines(StartLine: 1, Count: CountOfLines); //oCodeModule.DeleteLines(StartLine: 1, Count: oCodeModule.CountOfLines);
                                ViewCode(type: WbType, codes: FileCode);
                                if (exportCode)
                                {

                                    FilePath = "Microsoft_Excel_Objects";
                                    FilePath = Path.Combine(localAppPath, FilePath);
                                    if (!Directory.Exists(path: FilePath))
                                    { Directory.CreateDirectory(path: FilePath); }
                                    module.Export(FileName: String.Concat(Path.Combine(FilePath, module.Name), ".cls"));
                                    //CodeExport(filePath: String.Concat(Path.Combine(FilePath, module.Name), ".cls"), text: FileCode);
                                    CodeExport(filePath: FilePath, fileName: module.Name, fileExt: ".cls", text: FileCode);
                                }
                            }
                            //FileCode = CFO.Excel.VBA.CFOExcelVBA.GetStringResourceFile(ressFilePath: "Microsoft_Excel_Objects", ressFileName: "ThisWorkbook.cls", CodeOnly: true);
                            //oCodeModule.AddFromString(String: FileCode);
                            break;
                        case vbext_ComponentType.vbext_ct_MSForm: //"forms":
                            
                            CountOfLines = oCodeModule.CountOfLines;
                            if (CountOfLines > 0 )
                            {
                                tnForms.Nodes.Add(module.Name);
                                FileCode = oCodeModule.get_Lines(StartLine: 1, Count: CountOfLines); //oCodeModule.DeleteLines(StartLine: 1, Count: oCodeModule.CountOfLines);
                                if (exportCode)
                                {

                                    FilePath = "Forms";
                                    FilePath = Path.Combine(localAppPath, FilePath);
                                    if (!Directory.Exists(path: FilePath))
                                    { Directory.CreateDirectory(path: FilePath); }
                                    module.Export(FileName: String.Concat(Path.Combine(FilePath, module.Name), ".frm"));
                                    //CodeExport(filePath: FilePath, fileName: module.Name, fileExt: ".frm", text: FileCode);
                                }
                            }
                            break;
                        case vbext_ComponentType.vbext_ct_StdModule: //"modules": "class_modules":
                            
                            //lblMessage.Text = String.Format("Removing Module '{0}'", module.Name);
                            //lblMessage.Refresh();
                            //oModules.Remove(VBComponent: module);
                            CountOfLines = oCodeModule.CountOfLines;
                            if (CountOfLines > 0 )
                            {
                                tnModule.Nodes.Add(module.Name);
                                FileCode = oCodeModule.get_Lines(StartLine: 1, Count: CountOfLines); //oCodeModule.DeleteLines(StartLine: 1, Count: oCodeModule.CountOfLines);
                                if (exportCode)
                                {

                                    FilePath = "Modules";
                                    FilePath = Path.Combine(localAppPath, FilePath);
                                    if (!Directory.Exists(path: FilePath))
                                    { Directory.CreateDirectory(path: FilePath); }
                                    module.Export(FileName: String.Concat(Path.Combine(FilePath, module.Name), ".bas"));
                                    //CodeExport(filePath: FilePath, fileName: module.Name, fileExt: ".bas", text: FileCode);
                                }
                            }
                            break;
                    }

                }

            }
            catch (Exception Ex) { MessageBox.Show(String.Format("{0}\r\n{1}", Ex.Message, Ex.StackTrace), "LoadWorkbookCodes(Searching Module"); }

        }

        /// <summary>
        /// AreWorksheetExist
        /// </summary>
        /// <param name="WorksheetName"></param>
        /// <returns></returns>
        private Boolean AreWorksheetExist(string WorksheetName)
        {
            Boolean WorksheetExist = false;
            foreach (Excel.Worksheet oWs in oActiveWorkbook.Worksheets)
            {
                if (String.Equals(oWs.Name, WorksheetName)) { WorksheetExist = true; break; }
            }

            return WorksheetExist;
        }

        private void CodeExport(string filePath, string fileName, string fileExt, string text)
        {
            //string[] VB_Name = oCodeModule.get_Lines(StartLine: 5, Count: 1).Split('=');
            //string moduleName = VB_Name[1];
            string[] lines = File.ReadAllLines(path: Path.Combine(filePath, String.Concat(fileName, fileExt)));
            string lineFileName = lines[4].Split('=')[1].Replace("\"", "").Trim();
            if (String.Equals(lineFileName, "CFOConstant")) { string lineVersion = lines[10].Split('=')[1].Replace("\"", "").Trim().Replace(".", "_"); }
            File.Move(Path.Combine(filePath, String.Concat(fileName, fileExt)), destFileName: Path.Combine(filePath, String.Concat(lineFileName, fileExt)));
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!Object.Equals(oActiveWorkbook, null))
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oActiveWorkbook);
                oActiveWorkbook = null;
            }

            if (!object.Equals(oXlSource.oWorkbook, null))
            {
                oXlSource.oWorkbook.Close(SaveChanges: false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlSource.oWorkbook);
                oXlSource.oWorkbook = null;
            }

            if (!object.Equals(oXlDestination.oWorkbook, null))
            {
                oXlDestination.oWorkbook.Close(SaveChanges: false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlDestination.oWorkbook);
                oXlDestination.oWorkbook = null;
            }

            if (!object.Equals(ExcelApp, null))
            {
                ExcelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
                ExcelApp = null;
            }

            System.Windows.Forms.Application.Exit();
        }

        //private void listCodeToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    LoadWorkbookCodes(exportCode: false, WbType: WorkbookType.Source  );
        //}

        //private void exportCodesToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    ThisWorkbookCount = 0;
        //    LoadWorkbookCodes(exportCode: true, WbType: WorkbookType.Destination );
        //}

        private void tvListCodes_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void ListeSourceCodeMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void ListDestinationMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void ExportSourceCodesMenuItem_Click(object sender, EventArgs e)
        {
            //@excelFileName = ExcelApp.GetOpenFilename(FileFilter: "Excel File |*.xlsb;*,xlsb", Title: "Open Excel File to check!");
            try
            {
                ThisWorkbookCount = 0;
                LoadWorkbookCodes(oXlSource.oWorkbook, WbType: WorkbookType.Source);
            }
            catch (Exception Ex) { MessageBox.Show(String.Format("{0}\r\n{1}\r\n\r\n{2}\r\n\r\n{3}", Ex.Message, Ex.Source, Ex.StackTrace, "Please try launching again!"), "Loading Workbook Codes"); }
        }

        private void ExportDestinationCodesMenuItem_Click(object sender, EventArgs e)
        {
            //@excelFileName = ExcelApp.GetOpenFilename(FileFilter: "Excel File |*.xlsb;*,xlsb", Title: "Open Excel File to check!");
            try
            {
                ThisWorkbookCount = 0;
                LoadWorkbookCodes(oXlDestination.oWorkbook, WbType: WorkbookType.Destination );
            }
            catch (Exception Ex) { MessageBox.Show(String.Format("{0}\r\n{1}\r\n\r\n{2}\r\n\r\n{3}", Ex.Message, Ex.Source, Ex.StackTrace, "Please try launching again!"), "Loading Workbook Codes"); }
        }

        private string OpenWorkbook()
        {
            @excelFileName = ExcelApp.GetOpenFilename(FileFilter: "Macro Excel File |*.xlsb;*,xlsm",
                                                      Title: "Open Excel File to check!",
                                                      MultiSelect: false);
            ExcelApp.ScreenUpdating = false;
            ExcelApp.EnableEvents = false;
            return @excelFileName;
        }

        private void openSourceMenuItem_Click(object sender, EventArgs e)
        {
            oXlSource.filePath = OpenWorkbook();
            lblSourceFilePath.Text = oXlSource.filePath;
            oXlSource.oWorkbook =  GetActiveWorkbook();
            
            ThisWorkbookCount = 0;
            LoadWorkbookCodes(oXlSource.oWorkbook, WbType: WorkbookType.Source );

            ListeSourceCodeMenuItem.Enabled = true;
            ExportSourceCodesMenuItem.Enabled = true;
            closeSourceMenuItem.Enabled = true;
       }

        private void openDestinationMenuItem_Click(object sender, EventArgs e)
        {
            oXlDestination.filePath = OpenWorkbook();
            lblDestinationFilePath.Text = oXlDestination.filePath;
            oXlDestination.oWorkbook = GetActiveWorkbook();

            ThisWorkbookCount = 0;
            LoadWorkbookCodes(oXlDestination.oWorkbook, WbType: WorkbookType.Destination );

            ListDestinationCodeMenuItem.Enabled = true;
            ExportDestinationCodesMenuItem.Enabled = true;
            closeDestinationMenuItem.Enabled = true;
        }

        private void CodeExtractor_Load(object sender, EventArgs e)
        {

        }

        private void closeSourceMenuItem_Click(object sender, EventArgs e)
        {
            //ExcelApp.Workbooks.Item[oXlSource.oWorkbook.Name].Close();
            oXlSource.oWorkbook.Close(SaveChanges: true, Filename: String.Format("_{0}", oXlSource.oWorkbook.Name));
            ClearViewCode(WorkbookType.Source);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlSource.oWorkbook);
            oXlSource.oWorkbook = null;
        }

        private void closeDestinationMenuItem_Click(object sender, EventArgs e)
        {
            //ExcelApp.Workbooks.Item[oXlDestination.oWorkbook.Name].Close(); 
            oXlDestination.oWorkbook.Close(SaveChanges: true, Filename: String.Format("_{0}",oXlDestination.oWorkbook.Name));
            ClearViewCode(WorkbookType.Destination );
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlDestination.oWorkbook);
            oXlDestination.oWorkbook = null;
        }

        private void tvVBASolution_DoubleClick(object sender, EventArgs e)
        {
            TreeNode selectedNode = tvVBASolution.SelectedNode;
            foreach (VBComponent module in oXlSource.oModules)
            {
                if (string.Equals(module.Name, selectedNode.Text ))
                {
                    string fileCode = module.CodeModule.get_Lines(StartLine: 1, Count: module.CodeModule.CountOfLines);
                    ViewCode(type: WorkbookType.Source, codes: fileCode);
                    break;
                }
            }
        }

        private void saveCodesSourceMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void saveCodesDestinationMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}

//OpenFileDialog ofd = new OpenFileDialog();
//ofd.Filter = "Excel File |*.xlsb;*,xlsb";
//if (ofd.ShowDialog() == DialogResult.OK)
//{
//    string extn = Path.GetExtension(ofd.FileName);
//    if (extn.Equals(".xlsb"))
//    {
//        excelFileName = ofd.FileName;

//        if (excelFileName != "")
//        {
//            try
//            {
//                string xlFileName = Path.GetFileName(excelFileName);
//            }
//            catch (Exception ew)
//            {
//                MessageBox.Show("Errror:" + ew.ToString());
//            }
//        }
//    }
//}
