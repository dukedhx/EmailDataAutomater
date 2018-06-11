using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ConsoleApp1.Entities.Control
{
   public class EmailResponseConfig
    {
        public String[] replyRecipients, ccRecipients;
        public FileInfo template;
        public Dictionary<String, String> ResultMap;
        public IEnumerable<Tuple<String, String, Boolean?>> rows;
        public Outlook.Application oApp;
        public String defaultMessage, sentonbehalf,savesentfolder;
        
    }
}
