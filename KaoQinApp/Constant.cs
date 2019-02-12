using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace KaoQinApp
{
    public class Constant
    {
        public static Dictionary<string, string> NameDictionary
        {
            get
            {
                var dd = new Dictionary<string, string>();
                var source = ConfigurationManager.AppSettings;

                return source.Cast<string>().ToDictionary(s => s, s => source[s]);
            }
        }
        

        public static Dictionary<int, string> MonthDictionary = new Dictionary<int, string>{
            {1,"一" },
            {2,"二" },
            {3,"三" },
            {4,"四" },
            {5,"五" },
            {6,"六" },
            {7,"七" },
            {8,"八" },
            {9,"九" },
            {10,"十" },
            {11,"十一" },
            {12,"十二" },
        };
                
    }
}
