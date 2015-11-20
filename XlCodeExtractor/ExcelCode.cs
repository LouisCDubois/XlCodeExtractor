using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlCodeExtractor
{
    public class ExcelCode
    {

        vbext_ComponentType componentType { set; get; }
        string ligneCodes { set; get; }
        int ligneCodesCount { set; get; }
        VBComponent module { set; get; }

        public ExcelCode()
        {
            ;
        }
    }
}
