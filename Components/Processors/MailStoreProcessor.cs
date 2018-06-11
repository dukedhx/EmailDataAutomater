using ConsoleApp1.Components.Interfaces;
using ConsoleApp1.Entities.Control;
using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Linq;
using ClassLibrary1.Utils;
using System.Runtime.InteropServices;
using ConsoleApp1.Components.Consolidators;

namespace ConsoleApp1.Components.Processors
{
    public class MailStoreProcessor : IProcessor<MailStoreConfig, Object>
    {
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
                oApp = new Outlook.Application();
                oNS = oApp.GetNamespace("MAPI");
                oNS.Logon("Outlook", Type.Missing, false, true);
                var store = oApp.Session.Stores[subject.storename];
                rootf = store.GetRootFolder();
                configs[EPconfigsEnum.storename.ToString()] = subject.storename;

                configs[XCDconfigsEnum.dpath.ToString()] = subject.dpath;
                configs[EPconfigsEnum.saveMailPath.ToString()] = subject.savemailpath;

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
                  


                    return new EmailProcessor() { oApp = oApp, sendResponse = true, saveChanges = true, moveFolder=true, validColsCols=subject.validColsCol.SelectMany(vc=>vc), vconfigCol=subject.validColsCol }.process(mail,null,configs);
                }
                return true;

            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                return false;
            }
            finally {
      ;

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
