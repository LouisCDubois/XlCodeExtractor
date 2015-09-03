using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlCodeExtractor.Xl
{
    public class Xls
    {
        string fileName{ get; set; }
        string filePath { get; set; }

        List<string> XlObjects = new List<string>();
        List<string> ClassObjects = new List<string>();
        List<string> ModulesObjects = new List<string>();

        public Xls() 
        {
            XlObjects.Clear();
            ClassObjects.Clear();
            ModulesObjects.Clear();

        }
    }
}
