using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Entities.Misc
{
    public class ValidResultConfig
    {
        //XLSXConfig
        public String targetPath;
        public String tempatePath;
        public String targetSheet;
        public String headersPath;


        //Generic
        public String resultValMappingPath;
        public String resultMappingPath;
        public String IAFormatFields;

        //SQLConfig
        public String sqlQueryTemplate;      
        public String sqlPath;
        public String sqlTemplatePath;


    }
}
