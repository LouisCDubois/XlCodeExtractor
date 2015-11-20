using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlCodeExtractor
{
    public class TestParser
    {
        //public static void Main()
        public static void Test_Parser()
        {
            IniParser parser = new IniParser(@"C:\test.ini");

            String newMessage;

            newMessage = parser.GetSetting("appsettings", "msgpart1");
            newMessage += parser.GetSetting("appsettings", "msgpart2");
            newMessage += parser.GetSetting("punctuation", "ex");

            //Returns "Hello World!"
            Console.WriteLine(newMessage);
            Console.ReadLine();
        }
    }
}
