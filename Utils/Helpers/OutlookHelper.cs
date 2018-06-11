using System;

using System.Linq;

using Outlook = Microsoft.Office.Interop.Outlook;

using System.Text.RegularExpressions;


namespace ConsoleApp1.Utils
{
    public static class OutlookHelper
    {
        public static Boolean Running;
        public static int errcount = 0;

        public static readonly Regex EmailAddrRegex = new Regex(@"^([\w\.\-]+)@([\w\-]+)((\.(\w){2,3})+)$");
     
        //public static Boolean SendResponse(Outlook.Application oApp, Outlook.MailItem mail, String[] sucEmails, String[] rejEmails,FileInfo sucTemplate, FileInfo rejTemplate,IEnumerable<Tuple<String,String,Boolean?>> values,String sentonbehalf=null, Dictionary<String,String> ResultMap=null, Outlook.MAPIFolder savesentfolder=null )
        //{
        //    Outlook.MailItem reply=null;
        //    reply = mail.Reply();
        //    var sbc = mail.SenderEmailAddress;

        //    if (rejEmails == null)
        //    {
        //        reply = mail.Reply();
        //        try
        //        {
        //            var ignccs = ConfigurationManager.Configuration["IgnoredCCAddress"]?.Split(";");
        //            for (int i = 1; i <= mail.Recipients.Count; i++)
        //            {
        //                var recip = mail.Recipients[i];
        //                if (recip.Type == (int)Outlook.OlMailRecipientType.olCC && ignccs?.Contains(recip.Address, StringComparer.CurrentCultureIgnoreCase) != true)

        //                    reply.CC += (reply.CC?.Length > 0 ? ";" : "") + recip.Address;
                        
        //            }

        //        }catch(Exception ex)
        //        {
        //            Logger.Log(ex);
        //        }
        //    }
        //    else
        //    {
        //        reply = mail.Forward();
        //        foreach (var e in rejEmails.Where(e => EmailAddrRegex.IsMatch(e)))
        //            reply.Recipients.Add(e);
        //    }
        //    if (sucEmails != null)
        //        reply.BCC = GetEmailString(sucEmails);
        //    Logger.WriteToConsole($"Prepared validation response for [{reply.Subject}]");



        //    String htmlbody = null, mmsg = values == null ? "Success" : $"<table border='1'>{String.Join(Environment.NewLine, values.Select(r => $"<tr><td>{r.Item1}</td><td>{r.Item2}</td></tr>"))}</table>";
        //    FileInfo mtpath = values==null ? sucTemplate : rejTemplate;

        //    if (mtpath?.Exists == true)
        //    {
               
        //        var tpath =Path.Combine(Tools.GetExecutingPath,$"temp\\{System.Guid.NewGuid()}{Path.GetExtension(mtpath.Name)}");
        //        mtpath.CopyTo(tpath);
        //        var template = (Outlook.MailItem)oApp.Session.OpenSharedItem(tpath);
        //        htmlbody = template.HTMLBody;
        //        template.Close(Outlook.OlInspectorClose.olDiscard);
        //        Marshal.ReleaseComObject(template);
        //        template = null;
        //        File.Delete(tpath);
        //    }

        //    if (ResultMap != null)
        //        foreach (var amp in ResultMap)
        //            mmsg = mmsg.Replace(amp.Key,amp.Value);
            
        //    reply.HTMLBody =  htmlbody?.Replace("{{message}}", mmsg)??mmsg;


        //    var subj = reply.Subject;
        //    reply.SentOnBehalfOfName = sentonbehalf;

        //    if (savesentfolder != null)
        //        (reply.Copy() as Outlook.MailItem).Move(savesentfolder);

        //    if (!reply.Recipients.ResolveAll()) reply.CC = "";
        //    reply.Send();
        //    Logger.WriteToConsole($"Sent [{subj}]");
        //    Marshal.ReleaseComObject(reply);
        //    reply = null;
        //    return true;
        //}

        public static String GetSenderEmailAddr(Outlook.MailItem mail)
        {
            if (mail.Sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry || mail.Sender.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
              return   mail.Sender.GetExchangeUser()?.PrimarySmtpAddress;
            else
               return  mail.SenderEmailAddress;
        }

        //public static int ConsolidateEmail(String storename, String infolder, String sfolder,String retfolder, String dpath, IEnumerable<EmailValidatorConfigCol> validColsCol, String sentonbehalf,IEnumerable<SealedNameList> headers,FileInfo sucTemplate, String restricter,String[] sucEmails=null)
        //{
        //    Running = true;

        //    Outlook.Application oApp = null;

        //    Outlook.NameSpace oNS = null;
        //    Outlook.MAPIFolder oInbox = null;
        //    Outlook.MAPIFolder rdestFolder = null;
        //    Outlook.MAPIFolder destFolder = null;
        //    Outlook.MAPIFolder sucFolder = null;
        //    Outlook.MAPIFolder rootf = null;
        //    Outlook.MAPIFolder saveSentFolder = null;

        //    Outlook.MailItem reply = null;
        //    Outlook.Items items = null;
        //    Outlook.Attachments attachments = null;
        //    Outlook.Attachment attachment = null;
        //    Outlook.MailItem template = null;
        //    Outlook.MailItem mail = null;
        //    int n = 0;
         
        //    try
        //    {

        //        //Interop with Outlook and log on
        //        oApp = new Outlook.Application();               
        //        oNS = oApp.GetNamespace("MAPI");
        //        oNS.Logon("Outlook", Type.Missing, false, true);
        //        var store = oApp.Session.Stores[storename];
        //         rootf = store.GetRootFolder();
        //        oInbox = rootf.Folders[String.IsNullOrWhiteSpace(retfolder) ? "Inbox" : retfolder];
        //        saveSentFolder = String.IsNullOrWhiteSpace(ConfigurationManager.Configuration["saveSentFolder"])?null: rootf.Folders[ConfigurationManager.Configuration["saveSentFolder"]];
        //        rdestFolder = rootf.Folders[ConfigurationManager.Configuration["FADataRejFolder"]];
        //        sucFolder = String.IsNullOrWhiteSpace(sfolder)?null: rootf.Folders[sfolder];
        //        destFolder = rootf.Folders[infolder];
                
        //        destFolder.Session.SendAndReceive(false);
        //        items = String.IsNullOrWhiteSpace(restricter)?destFolder.Items: destFolder.Items.Restrict(restricter);
        //        Outlook.MAPIFolder moveto = null;
        //        var validColsCols = validColsCol.SelectMany(vc=>vc);
                

        //        //Proceed to iterate mails per defined filters
        //        if ((mail = items.GetLast() as Outlook.MailItem) != null)
        //        {


        //            Dictionary<EmailValidatorConfig, IEnumerable<Tuple<String, String, Boolean?>>> avalCols = new Dictionary<EmailValidatorConfig, IEnumerable<Tuple<String, String, Boolean?>>>();
        //            bool success = true, error=false,clearCat = false, movetoinbox=false;

        //            Logger.Log($"Start to process [{mail.Subject}]");
        //            attachments = mail.Attachments;
        //            String rhtml = null;
          
        //            if (attachments.Count > 0)
        //            {
        //                var SN = ConsolidateHelper.TimeLapsedID;


        //                var evals = new Dictionary<String, String>() {
        //                     { ConfigHelper.Email_Date.name, mail.ReceivedTime.ToString("MM/dd/yyyy HH:mm:ss")},
        //                    { ConfigHelper.Email_Subject.name,mail.Subject},
        //                    { ConfigHelper.Sender_Email_Address.name,GetSenderEmailAddr(mail)},
        //                    { ConfigHelper.SN.name,SN}


        //                };
        //                //Logger.WriteToConsole(evals);
        //                var d = Directory.CreateDirectory($@"temp\{SN}");
        //                Dictionary<String, ValidResults> resultsCol = new Dictionary<String, ValidResults>();
        //                var stopOnReject = false;


        //                //Proceed to iterate attachments
        //                for (int i = 1; i <= attachments.Count; i++)
        //                {
        //                    if (stopOnReject|| error) break;
                         
        //                    try
        //                    {
        //                        attachment = attachments[i];
        //                        if (attachment.FileName.Split(".").Last().ToLower().Equals("xlsm"))
        //                        {

        //                            var path = Tools.GetUniqueFileName( Path.Combine(d.FullName, attachment.FileName));
        //                            attachment.SaveAsFile(path);
        //                            Logger.Log($"{attachment.FileName} saved to {path} ...");
        //                            ValidResults vr = new ValidResults();

        //                            foreach (var validCols in validColsCol)
        //                            {
        //                                var aresultscol = ConsolidateHelper.validateWS(validCols.sheetName, validCols.version, validCols.macroSnippet, new FileInfo(path), validCols.Where(vc => vc.fileInfo != null && !vc.useCustomValidation).Select(vc => vc.fileInfo), validCols.UnitValidator as CellValidator);

        //                                if (aresultscol == null)
        //                                    continue;
        //                                Logger.WriteToConsole($"Processing results [{attachment.FileName}]");
        //                                foreach (var aresults in aresultscol)
        //                                {
        //                                    var t = validCols.FirstOrDefault(vc => vc.fileInfo.FullName.Equals(aresults.Key.FullName));
        //                                    var results = aresults.Value;
        //                                    if (results.Any(r => r.Value != true))
        //                                    {

        //                                        if (!t.continueOnReject)




        //                                            avalCols.Clear();


        //                                        if (t.rejectOnInvalid) success = false;
        //                                        var rsts = results.Select(r => Tuple.Create(attachment.FileName, r.Key, r.Value));
        //                                        if (avalCols.ContainsKey(t))
        //                                            avalCols[t].Concat(rsts);
        //                                        else avalCols.Add(t, new LinkedList<Tuple<String, String, Boolean?>>(rsts));



        //                                        Logger.WriteToConsole($"Logged to DB for [{path}], result:{DBHelper.InsertOptAttachmentReceived(new OptAttachmentReceived() { AttachmentFileName = attachment.FileName, ExecutionTime = DateTime.Now, ImportErrorMessage = String.Join(",", results.Select(r => $"[{r.Key}][{r.Value}]")), Subject = mail.Subject, SenderEmailAddress = evals[ConfigHelper.Sender_Email_Address.name], RecievedTime = mail.ReceivedTime, _ImportResult = success })                  }                 ");



        //                                        if (!t.continueOnReject)
        //                                        {
        //                                            stopOnReject = true;
        //                                            break;
        //                                        }
        //                                        }
        //                                    if (success )



        //                                        vr.AddAll(results);


        //                                    if (t.sucTemplate != null && !avalCols.ContainsKey(t))
        //                                        avalCols.Add(t, null);
                                            
        //                                }

        //                                if (stopOnReject) break;
        //                            }
        //                            if (success)
        //                            {
                                        
        //                                if (!String.IsNullOrWhiteSpace(vr.sheetName))
        //                                resultsCol.Add(path, vr);
        //                                //  rhtml.Add((vr.ID = attachment.FileName), "Success");
        //                                Logger.WriteToConsole($"Successfully validated {attachment.FileName}");

        //                            }
                                   

        //                        }

        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        error = true;errcount++;

        //                        Logger.Log(ex);
        //                        rhtml = ex.StackTrace.ToString();
        //                        success = false;
        //                        break;
        //                    }
        //                    finally
        //                    {
        //                        Marshal.ReleaseComObject(attachment);
        //                        attachment = null;
        //                    }
        //                }


        //                //Proceed to load data results

        //                if (success && !error && resultsCol.Any())
        //                {

        //                    long ticks = TimeSpan.FromMinutes(5).Ticks;
        //                    var lockobj = ConfigurationManager.Locks[dpath];
        //                    try
        //                    {
        //                        while (ticks-- > 0)
        //                        {
        //                            if (Monitor.TryEnter(lockobj))
        //                            {


        //                                // var password = ConfigurationManager.Configuration["DataResultPassword"];
        //                                Logger.WriteToConsole($"Saving DataResults to [{dpath}]");
        //                                var tdpath = $"temp\\{System.Guid.NewGuid()}.xlsx";
        //                                Tools.FileOverWriteCopy(dpath, tdpath);

        //                                using (var p = new ExcelPackage(new FileInfo(tdpath)))
        //                                {

        //                                    ExcelWorksheet worksheet = ConsolidateHelper.GetWorksheetOrAdd(p.Workbook, Path.GetFileNameWithoutExtension(dpath));
        //                                    foreach (var results in resultsCol)
        //                                        success = ConsolidateHelper.writeDataResult(results.Key, results.Value.sheetName, worksheet, evals.Concat(results.Value.Select(e => new KeyValuePair<String, String>(e.Key, "")).Append(new KeyValuePair<string, string>(ConfigHelper.Filename.name,Path.GetFileName( results.Key)))).Concat(results.Value.vals), headers, results.Value.MatcherDict);
        //                                    if (!success) break;
        //                                    int r = worksheet.Dimension.End.Row + 1;

        //                                    var ccoli = headers.First(h => string.Equals(h.name, "count po", StringComparison.OrdinalIgnoreCase)).value;
        //                                    var pocol = ConsolidateHelper.GetColumnName(headers.First(h => string.Equals(h.name, "IA PO #", StringComparison.OrdinalIgnoreCase)).value - 1);

        //                                    worksheet.Cells[2, ccoli].CreateArrayFormula($"SUM(1/COUNTIF({pocol}2:{pocol}{r - 1},{pocol}2:{pocol}{r - 1}))");
        //                                    worksheet.Cells[2, ccoli].Calculate();
        //                                    int iac = 0;
        //                                    var tcell = worksheet.Cells[2, headers.First(h => string.Equals(h.name, "Total IA Count", StringComparison.OrdinalIgnoreCase)).value];
        //                                    Int32.TryParse(tcell.Value?.ToString(), out iac);
        //                                    tcell.Value = iac + 1;

        //                                    p.Save();
        //                                    Logger.WriteToConsole($"Successfully validated saved DataResults to [{tdpath}]");
        //                                }

        //                                Tools.FileOverWriteMove(tdpath, dpath);
        //                                Logger.WriteToConsole($"Successfully moved saved DataResults to [{dpath}]");
        //                                String spath = "";
        //                                try
        //                                {
        //                                    spath = Path.Combine(Directory.CreateDirectory(Path.GetDirectoryName(dpath)).FullName, string.Join("", mail.Subject.Split(Path.GetInvalidFileNameChars().Concat(new Char[] { }).ToArray()))) + ".msg";
        //                                    var tpath = Path.Combine(d.FullName, $"{System.Guid.NewGuid()}.msg");
        //                                    Logger.WriteToConsole($"Saving mail to {tpath}");

        //                                    mail.SaveAs(tpath, Outlook.OlSaveAsType.olMSG);
        //                                    Logger.WriteToConsole($"Moving mail to {(spath = Tools.GetUniqueFileName(spath))}");

        //                                    Tools.FileOverWriteMove(tpath, spath);
        //                                    Logger.WriteToConsole($"Mail saved to {spath}");

                                           
        //                                }
        //                                catch (Exception ex)
        //                                {
        //                                    Logger.Log(ex);
        //                                }

        //                                try
        //                                {
        //                                    var logrst = false;
        //                                    foreach (var fname in resultsCol)

        //                                        logrst = DBHelper.InsertOptAttachmentReceived(new OptAttachmentReceived()
        //                                        {
        //                                            AttachmentFileName = Path.GetFileName(fname.Key),
        //                                            ExecutionTime = DateTime.Now,
        //                                            SavedFileName = spath,
        //                                            Subject = mail.Subject,
        //                                            SenderEmailAddress = evals[ConfigHelper.Sender_Email_Address.name],
        //                                            RecievedTime = mail.ReceivedTime,
        //                                            _ImportResult = true
        //                                        });
        //                                    Logger.WriteToConsole($"Logged to DB for [{spath}], result:{logrst}");
        //                                }
        //                                catch (Exception ex)
        //                                {
        //                                    Logger.Log(ex);
        //                                }
        //                                break;





        //                            }
        //                            else
        //                            {
        //                                Thread.Sleep(1);
        //                                if (ticks == 0) error = true;
        //                            }

        //                        }
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        errcount++;

        //                        error = true;
        //                        Logger.Log(ex);
        //                        rhtml = ex.StackTrace.ToString();
        //                        success = false;
        //                        OA.Lasterror = ex;
        //                    }
        //                    finally { Monitor.Exit(lockobj); }
        //                    if (d.Exists)
        //                        d.Delete(true);
                            

        //                }


        //                //Proceed to send response
        //                if (!error)
        //                {
        //                     if (success)
        //                    {
                                
        //                        clearCat = true; 

        //                        Logger.WriteToConsole($"Cleared mark for [{mail.Subject}]");
        //                    }

        //                    if (avalCols.Any())

        //                        foreach (var avalcol in avalCols)
        //                        {

        //                            var t = validColsCols.First(vc => vc == avalcol.Key);


        //                                SendResponse(oApp, mail, t.sucEmails, t.rejEmails, t.sucTemplate, t.rejTemplate, avalcol.Value, sentonbehalf, avalcol.Key.ResultMap, saveSentFolder);

        //                            if (!success)
        //                            {
        //                                if (String.IsNullOrWhiteSpace(avalcol.Key.rejFolder)||(moveto = rootf.Folders[avalcol.Key.rejFolder])==null)      
        //                                clearCat = true;

        //                            }
        //                            Logger.WriteToConsole($"Sending response template [{(success ? t.sucTemplate : t.rejTemplate)}] for  [{mail.Subject}]");

        //                        }
        //                    else if (success && resultsCol.Any())
        //                    {


        //                        SendResponse(oApp, mail, sucEmails, null, sucTemplate, null, null, sentonbehalf, null, saveSentFolder);

        //                        moveto = sucFolder;
        //                        Logger.WriteToConsole($"Sending success template [{sucTemplate}] response [{mail.Subject}]");



        //                    }

        //                    else if (success)
        //                    {
        //                        movetoinbox = true;
        //                        Logger.WriteToConsole($"Moving [{oInbox?.Name}] due to having nothing to load...");

        //                    }
        //                    else Logger.WriteToConsole($"Unknown error occurred for [{mail.Subject}] ...");


                           

        //                }
        //            }
        //            else movetoinbox = true;

                   

        //            //Error notification
        //            if (error) {
                       

        //                    reply = mail.Forward();

        //                    reply.HTMLBody = rhtml ?? "Error occurred!";
        //                    foreach (var e in ConfigurationManager.Configuration["AdminEmail"]?.Split(";")?.Where(e => EmailAddrRegex.IsMatch(e)))
        //                        reply.Recipients.Add(e);
        //                     reply.Send();
        //                    moveto = rdestFolder;
        //                    Logger.WriteToConsole($"Prepared internal error response for [{reply.Subject}]");
                        
        //            }
                  

        //            //Proceed to post process the mail
        //            if (movetoinbox) { clearCat = true;moveto = oInbox; }

        //            if (clearCat) {mail.Categories = null; mail.Save(); }

        //            if (moveto != null && moveto.Name != destFolder.Name)
        //            {
        //                mail.Move(moveto);

        //                    Logger.WriteToConsole($"Moved [{mail.Subject}] to [{moveto.Name}]");
        //            }
        //            moveto = null;

        //            OA.LastEmailProcessed = mail.Subject;


        //            if (attachments != null)
        //                Marshal.ReleaseComObject(attachments);
        //            if (mail != null)
        //                Marshal.ReleaseComObject(mail);
        //            if (reply != null)
        //                Marshal.ReleaseComObject(reply);

        //            attachments = null;
        //            mail = null;
        //            reply = null;
        //            rdestFolder?.Session?.SendAndReceive(false);
        //            destFolder?.Session?.SendAndReceive(false);
        //                oInbox?.Session?.SendAndReceive(false);
        //            sucFolder?.Session?.SendAndReceive(false);
                    
        //            errcount = 0;
        //            n++;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        errcount++;
        //        try
        //        {
        //            if (mail != null) OA.LastEmailProcessed = mail.Subject;
        //            mail.Categories = null;
        //            mail.Save();
        //        }catch(Exception e)
        //        {
        //            Logger.Log(e);
        //        }
        //        OA.Lasterror = ex;
        //        Logger.Log(ex);
        //    }
        //    finally
        //    {


        //        //Important: release Interop COM objects to avoid memeory leak
        //        if (reply != null)
        //        {
        //            reply.Delete();
        //            //reply.Close(Outlook.OlInspectorClose.olDiscard);
        //            Marshal.ReleaseComObject(reply);
        //        }

        //        if (attachment != null) Marshal.ReleaseComObject(attachment);
        //        if (attachments != null) Marshal.ReleaseComObject(attachments);

        //        if (mail != null)             
        //            Marshal.ReleaseComObject(mail);

        //        if (template != null) Marshal.ReleaseComObject(template);
                
        //        if (items != null) Marshal.ReleaseComObject(items);

        //        if (oInbox != null) Marshal.ReleaseComObject(oInbox);
        //        if (rdestFolder != null) Marshal.ReleaseComObject(rdestFolder);
        //        if (sucFolder != null) Marshal.ReleaseComObject(sucFolder);
        //        if (destFolder != null) Marshal.ReleaseComObject(destFolder);
        //        if (rootf != null) Marshal.ReleaseComObject(rootf);
        //        if (oNS != null)
        //        {
        //            oNS.Logoff();
        //            Marshal.ReleaseComObject(oNS);
                    
        //        }
        //        if (oApp != null) Marshal.ReleaseComObject(oApp);

        //        GC.Collect();
        //        GC.WaitForPendingFinalizers();
        //        Logger.WriteToConsole($"Released OA objects ...");

        //    }
        //    Running = false;
        //    OA.LastRunFinished = DateTime.Now;
        //    return n;
        //}

        public static  String GetEmailString(String[] emails)
        {
            return String.Join(";", emails.Where(e => !String.IsNullOrWhiteSpace(e) && EmailAddrRegex.IsMatch(e)));
        }
        
    }
}
