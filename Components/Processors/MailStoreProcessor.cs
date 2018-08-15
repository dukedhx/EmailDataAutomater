using ConsoleApp1.Components.Interfaces;
using ConsoleApp1.Entities.Control;
using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Linq;
using ClassLibrary1.Utils;
using System.Runtime.InteropServices;

namespace ConsoleApp1.Components.Contollers
{
    public class MailStoreProcessor : IProcessor<MailStoreConfig, Object>
    {
        public ProperNameController pnc; //Adapter to fill in proper nouns in string, e.g. {year},{today},{guid} ...
        /*
            Log on to the running Outlook instance
            This processor requires a running Outlook instance with sufficient access to all relevant Exchange resources (e.g. mailstores, folders and their permissions)
             */
        public static Outlook.NameSpace logOn(Outlook.Application oApp)
        {
            
            Outlook.NameSpace oNS = oApp.GetNamespace("MAPI");
            oNS.Logon("Outlook", Type.Missing, false, true);
            return oNS;
        }
        /*
            Process mail items in a given mailstore and its designated folders, see class MailStorConfig for config details
             */
        public bool? process(MailStoreConfig subject, Object controller, IDictionary<string, string> configs = null)
        {
            Outlook.Application oApp = null;

            Outlook.NameSpace oNS = null;
            Outlook.MAPIFolder rdestFolder = null;
            Outlook.MAPIFolder destFolder = null;
            Outlook.MAPIFolder rootf = null;
        
            Outlook.Items items = null;
      
            Outlook.MailItem mail = null;
            try
            {
                configs = configs ?? new Dictionary<string, string>();
                //Interop with Outlook and log on
                oNS = logOn(oApp = new Outlook.Application());
                var store = oApp.Session.Stores[subject.storename];
                rootf = store.GetRootFolder();
                configs[EPconfigsEnum.storename.ToString()] = subject.storename;

             //   configs[XCDconfigsEnum.dpath.ToString()] = subject.dpath;
                configs[EPconfigsEnum.saveMailPath.ToString()] = pnc?.execute (subject.savemailpath,null,null)?? subject.savemailpath;

                configs[EPconfigsEnum.retfolder.ToString()] = String.IsNullOrWhiteSpace( subject.retfolder) ? "Inbox" : subject.retfolder;
   
                configs[EPconfigsEnum.rdestFolder.ToString()] = subject.rejfolder;
                configs[EPconfigsEnum.sucTemplate.ToString()] = subject.sucTemplate?.FullName;
                configs[EPconfigsEnum.sucFolder.ToString()] = subject.sucfolder;
                configs[EPconfigsEnum.sentonbehalf.ToString()] = subject.sentonbehalf;
                
                destFolder = rootf.Folders[ configs[EPconfigsEnum.destFolder.ToString()] = subject.infolder];

                destFolder.Session.SendAndReceive(false);
                items = String.IsNullOrWhiteSpace(subject.restricter) ? destFolder.Items : destFolder.Items.Restrict(subject.restricter);
        

                //Proceed to iterate mails per defined filters
                if ((mail = items.GetLast() as Outlook.MailItem) != null)
                {
                  


                    return new EmailProcessor() { oApp = oApp, sendResponse = true, saveChanges = true, moveFolder=true, validColsCols=subject.validColsCol.SelectMany(vc=>vc), vconfigCol=subject.validColsCol,pnc=new ProperNameController() }.process(mail,null,configs);
                }
                return true;

            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                return false;
            }
            finally {
      

                if (mail != null)
                    Marshal.ReleaseComObject(mail);


                if (items != null) Marshal.ReleaseComObject(items);

                if (rdestFolder != null) Marshal.ReleaseComObject(rdestFolder);
           
                if (rootf != null) Marshal.ReleaseComObject(rootf);
                if (oNS != null)
                {
                    oNS.Logoff();
                    Marshal.ReleaseComObject(oNS);

                }
                if (oApp != null) Marshal.ReleaseComObject(oApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();

            }
        }
    }
}
