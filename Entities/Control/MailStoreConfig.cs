using ClassLibrary1.Utils.Persistence;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ConsoleApp1.Entities.Control
{
    public class MailStoreConfig
    {
        public string storename, infolder, sucfolder, rejfolder, dpath, sentonbehalf, restricter, retfolder,savemailpath;
        public String[] sucEmails;
public FileInfo sucTemplate;
        public IEnumerable<SealedNameList> headers;
        public  IEnumerable<EmailValidatorConfigCol> validColsCol;
    }
}
