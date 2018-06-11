using ClassLibrary1.Utils.Persistence;
using System;
using System.Collections;
using System.Collections.Generic;


namespace ConsoleApp1.Utils
{

    
    internal static class ConfigHelper
    {
        public static readonly string Email_Date = "Email Date";
        public static readonly string Sender_Email_Address = "Sender Email Address";
        public static readonly string Email_Subject = "Email Subject";
        public static readonly string Filename = "Filename";
        public static readonly string SN = "SN";

        

      


    }

   

    public sealed class DATARSTHeader : SealedNameList
    {
        public DATARSTHeader(string name, string ID) : base(name, ID)
        {
        }

        public DATARSTHeader(string name, string ID, int value) : base(name, ID, value)
        {
        }

        private void init()
        {
            //if (eheaders == null)
            //    eheaders = new SealedNameList[]
            //    {
            //ConfigHelper. Email_Date ,  ConfigHelper. Sender_Email_Address , ConfigHelper. Email_Subject, ConfigHelper.Filename, ConfigHelper.SN
            //    };
        }
      

    }

    
}
