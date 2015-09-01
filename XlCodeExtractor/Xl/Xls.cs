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

        List<string> XlObjects;
        List<string> ClassObjects;
        List<string> ModulesObjects;

        public Xls() 
        {
            XlObjects.Clear();
            ClassObjects.Clear();
            ModulesObjects.Clear();
        }
    }
}
