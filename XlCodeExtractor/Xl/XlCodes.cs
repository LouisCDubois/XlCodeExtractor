using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace XlCodeExtractor.Xl
{
    public abstract class XlCodes
    {
        public string fileName{ get; set; }
        public string filePath { get; set; }
        public Excel.Workbook oWorkbook { set; get; }

        public List<string> XlObjects = new List<string>();
        public List<string> ClassObjects = new List<string>();
        public List<string> ModulesObjects = new List<string>();

        //List<List<string>> XlModulesObjects = new List<List<string>>();


        public XlCodes() 
        {
            //XlModulesObjects.Add(new List<string>());

            XlObjects.Clear();
            ClassObjects.Clear();
            ModulesObjects.Clear();

        }
    }
}
