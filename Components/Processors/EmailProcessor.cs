using ClassLibrary1.Utils;
using ConsoleApp1.Components.Controllers;
using ConsoleApp1.Components.Interfaces;
using ConsoleApp1.Entities.Control;
using ConsoleApp1.Entities.Misc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using ConsoleApp1.Utils;
using ConsoleApp1.Components.Consolidators;
using ClassLibrary1.Utils.Persistence;
using System.Threading;
//using OfficeOpenXml;
using YamlDotNet.RepresentationModel;
using ConsoleApp1.Components.Processors;
using System.Threading.Tasks;

public enum EPconfigsEnum { sucTemplate=1, sentonbehalf=3, saveSentFolder = 4 ,sucFolder=5, AdminEmail =6, rdestFolder=7,retfolder=8 , destFolder =9,storename=10, dataResultMappingPath =11, IAFormatFields =12, dataResultValMapping =13,saveMailPath=14,emailLogConnString=15, emailLogTemplate=16}

public class EmailProcessor: ValidatorConfigProcessor<Outlook.MailItem>
    {

    public Boolean error, stopOnReject;
  //  public int errcount;
  //  public String rhtml;
    public Outlook.Application oApp;
    public bool sendResponse, saveChanges, moveFolder;
    public IEnumerable<EmailValidatorConfig> validColsCols;
    protected String currentEmailFilePath;
    protected bool mergeTempResults(String tdpath, String dpath)
    {

        long ticks = TimeSpan.FromMinutes(5).Ticks;
        var lockobj = ConfigurationManager.Locks[dpath];
        try
        {
            while (ticks-- > 0)
            {

                if (Monitor.TryEnter(lockobj))
                {
                    Tools.FileOverWriteMove(tdpath, dpath);
                    Logger.Log($"Moved to {dpath}!");
                    return true;
                }

                Thread.Sleep(1);

            }
            throw new Exception($"Timed out when locking:{dpath}");
        }
        finally
        {
            Monitor.Exit(lockobj);
            try
            {
                if (File.Exists(tdpath)) File.Delete(tdpath);
            }
            catch (Exception ex) { Logger.Log(ex); }
        }

    }


    protected IDictionary<string, IEnumerable<KeyValuePair<string, string>>> getBindMaps(IDictionary<String,String> configs)
    {

        Dictionary<string, IEnumerable<KeyValuePair<string, string>>> bindmaps = new Dictionary<string, IEnumerable<KeyValuePair<string, string>>>();
        var yaml = new YamlStream();

        using (var sr = new FileInfo(configs[EPconfigsEnum.dataResultMappingPath.ToString()]).OpenText())

            yaml.Load(sr);

        bindmaps.Add (XCDconfigsEnum.resultMap.ToString(), ((YamlMappingNode)yaml.Documents[0].RootNode).Children.Select(entry => new KeyValuePair<string, string>( entry.Key.ToString(),  entry.Value.ToString() )));


        using (var sr = new FileInfo(configs[EPconfigsEnum.IAFormatFields.ToString()]).OpenText())

            yaml.Load(sr);

        bindmaps.Add(XCDconfigsEnum.IAFormatFields.ToString(),((YamlMappingNode)yaml.Documents[0].RootNode).Children.Select(entry => new KeyValuePair<string, string>( entry.Key.ToString(),  entry.Value.ToString() )));

        using (var sr = new FileInfo(configs[EPconfigsEnum.dataResultValMapping.ToString()]).OpenText())

            yaml.Load(sr);

        bindmaps.Add(XCDconfigsEnum.dataResultValMapping.ToString(), ((YamlMappingNode)yaml.Documents[0].RootNode).Children.Select(entry => new KeyValuePair<string, string>(entry.Key.ToString(), entry.Value.ToString())));

        return bindmaps;
    }

    protected bool saveMailtoLocal(Outlook.MailItem mail, String path)
    {
      
         
           var spath = Path.Combine(Directory.CreateDirectory(path).FullName, string.Join("", mail.Subject.Split(Path.GetInvalidFileNameChars().Concat(new Char[] { }).ToArray()))) + ".msg";
            var tpath = Path.Combine(Directory.CreateDirectory("temp").FullName, $"{Guid.NewGuid()}.msg");
            Logger.WriteToConsole($"Saving mail to {tpath}");

            mail.SaveAs(tpath, Outlook.OlSaveAsType.olMSG);
            Logger.WriteToConsole($"Moving mail to {(spath = Tools.GetUniqueFileName(spath))}");
        
            Tools.FileOverWriteMove(tpath, currentEmailFilePath= spath);
            return true;
       
    }

    public override bool? process(Outlook.MailItem subject, IControl<FileInfo, IValidationConfigCol<IValidator<string, string>, EmailValidatorConfig>, Dictionary<string,ValidResults>> controller, IDictionary<string, string> configs = null)
    {
            String moveto = null,rhtml=null;
        bool clearCat = false, movetoinbox = false;
        if (subject.Attachments?.Count>0) {
            Outlook.Attachment attachment = null;
            var attachments = subject.Attachments;



            Dictionary<EmailValidatorConfig, IEnumerable<Tuple<String, String, Boolean?>>> avalCols = new Dictionary<EmailValidatorConfig, IEnumerable<Tuple<String, String, Boolean?>>>();
            Dictionary<String, ValidResults> resultsCol = new Dictionary<String, ValidResults>(); ;
            try
            {
                for (int i = 1; i <= attachments.Count; i++)
                {
                    if (stopOnReject || error) break;


                    attachment = attachments[i];
                    if (attachment.FileName.Split(".").Last().ToLower().Equals("xlsm"))
                    {




                        var d = Directory.CreateDirectory($@"temp\{Guid.NewGuid()}");

                        var path = Tools.GetUniqueFileName(Path.Combine(d.FullName, attachment.FileName));




                        attachment.SaveAsFile(path);
                        Logger.Log($"{attachment.FileName} saved to {path} ...");


                        if (controller == null)
                            ExecuteValidationConfig(vconfigCol, avalCols, path, resultsCol);

                       else resultsCol.Concat(controller.execute(new FileInfo(path), null, configs));







                    }


                    Logger.WriteToConsole($"Successfully validated {attachment.FileName}");


                    //                                    Logger.WriteToConsole($"Logged to DB for [{path}], result:{DBHelper.InsertOptAttachmentReceived(new OptAttachmentReceived() { AttachmentFileName = attachment.FileName, ExecutionTime = DateTime.Now, ImportErrorMessage = String.Join(",", results.Select(r => $"[{r.Key}][{r.Value}]")), Subject = mail.Subject, SenderEmailAddress = evals[ConfigHelper.Sender_Email_Address.name], RecievedTime = mail.ReceivedTime, _ImportResult = success })}");
                }
                var guid = Guid.NewGuid().ToString();

                var pstpath = configs.ContainsKey(PSTconfigsEnum.pstPath.ToString()) && !String.IsNullOrWhiteSpace(configs[PSTconfigsEnum.pstPath.ToString()]) ? Directory.CreateDirectory(configs[PSTconfigsEnum.pstPath.ToString()]).FullName : "";

                var evals = new Dictionary<String, String>() {
                             { ConfigHelper.Email_Date, subject.ReceivedTime.ToString("MM/dd/yyyy HH:mm:ss")},
                            { ConfigHelper.Email_Subject, Tools.SafeSubstring(subject.Subject,0,199)},
                            { ConfigHelper.Sender_Email_Address,OutlookHelper.GetSenderEmailAddr(subject)},
                            { ConfigHelper.SN,ConsolidateHelper.TimeLapsedID}
                        };

                if (success&&resultsCol.Any())
                {

                    var tdpath = $"temp\\{guid}.xlsx";
                    var dpath = configs[XCDconfigsEnum.dpath.ToString()];
                  

                    if (File.Exists(dpath))
                        if (!Tools.FileOverWriteCopy(dpath, tdpath))

                            throw new Exception("Unable to create temp data result @:" + tdpath);
                    var bindmaps = getBindMaps(configs);
                  
                    configs[XCDconfigsEnum.cwwsn.ToString()] = Path.GetFileNameWithoutExtension(dpath);
                    var headers = DATARSTHeader.headers[configs[XCDconfigsEnum.headersID.ToString()]];
                    using (PersistenceSQLProcessor pp = new PersistenceSQLProcessor(String.IsNullOrWhiteSpace(pstpath)?"":Path.Combine(pstpath,guid+".data"), configs.ContainsKey(PSTconfigsEnum.pstConnString.ToString())&&!String.IsNullOrWhiteSpace(configs[PSTconfigsEnum.pstConnString.ToString()]) ? configs[PSTconfigsEnum.pstConnString.ToString()] : ""))
                    using (XlsxResultProcessor p = new XlsxResultProcessor(tdpath, configs[XCDconfigsEnum.cwwsn.ToString()], bindmaps, headers))
                        success = new XlsxConsolidator()
                        {

                            evals = evals,
                            headers = headers,
                            bindmaps = bindmaps

                        }.consolidate(new List<IProcessor<IDictionary<String, String>, Object>>() { configs.ContainsKey(PSTconfigsEnum.pstConnString.ToString()) && !String.IsNullOrWhiteSpace(configs[PSTconfigsEnum.pstConnString.ToString()]) ? pp : null, p }, resultsCol, configs);

                   


                    if (success && mergeTempResults(tdpath, dpath))
                    {
                        moveto = configs[EPconfigsEnum.sucFolder.ToString()];
                        moveto = configs[EPconfigsEnum.sucFolder.ToString()];
                        try
                        {
                            if (configs.ContainsKey(EPconfigsEnum.saveMailPath.ToString()))
                                saveMailtoLocal(subject, configs[EPconfigsEnum.saveMailPath.ToString()]);

                        }
                        catch (Exception ex)
                        {
                            Logger.Log(ex);
                        }
                    }
                    else
                    {
                        Logger.Log($"Attachment [{attachment.FileName}] consolidation filed...");

                        return false;
                    }
                   
                    
                }

                try
                {

                    foreach (var fname in resultsCol)



                        using (PersistenceSQLProcessor pp = new PersistenceSQLProcessor(Path.Combine(pstpath, guid + ".log"), configs[EPconfigsEnum.emailLogConnString.ToString()]))
                            Logger.WriteToConsole($"Logged to DB for [{fname.Key}], result:{pp.process(evals.Append(new KeyValuePair<string, string>("ImportResult", fname.Value.Any(v => v.Value == false) ? "FAIL" : "SUCCESS")).Append(new KeyValuePair<string, string>("ImportErrorMessage", Tools.SafeSubstring($"[{String.Join(",", fname.Value?.Select(v => $"{{\"{v.Key}\":\"{v.Value?.ToString() ?? "null"}\"}}"))}]", 0, 499))).Append(new KeyValuePair<string, string>("AttachmentFileName", Tools.SafeSubstring(Path.GetFileName(fname.Key), 0, 99))).Append(new KeyValuePair<string, string>("Filename", Tools.SafeSubstring(Path.GetFileName(currentEmailFilePath ?? ""), 0, 99))).Append(new KeyValuePair<string, string>("guid", guid)).ToDictionary(prop => prop.Key, prop => prop.Value), null, new Dictionary<String, String>() { { PSTconfigsEnum.template.ToString(), configs[EPconfigsEnum.emailLogTemplate.ToString()] } })}");

                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }

            }
            catch (Exception ex)
            {
                error = true;

                Logger.Log(ex);
                rhtml = ex.StackTrace.ToString();
                success = false;
             
            }
            finally
            {
                Marshal.ReleaseComObject(attachment);
                attachment = null;
            }
            


            if (error)
            {
                new EmailResponseController().execute(subject, new EmailResponseConfig() { oApp = oApp, defaultMessage = rhtml ?? "Error occurred!", replyRecipients = configs[EPconfigsEnum.AdminEmail.ToString()]?.Split(";") });


                moveto = configs[EPconfigsEnum.rdestFolder.ToString()];
                if (String.IsNullOrEmpty(moveto)|| moveto==configs[EPconfigsEnum.destFolder.ToString()]) clearCat = true;
                Logger.WriteToConsole($"Prepared internal error response for [{subject.Subject}]");
            }
            else
            {

                

                if (success)
                {

                    clearCat = true;

                    Logger.WriteToConsole($"Cleared mark for [{subject.Subject}]");
                }

                if (avalCols.Any())

                    foreach (var avalcol in avalCols)
                    {

                        var t = validColsCols.First(vc => vc == avalcol.Key);


                        new EmailResponseController().execute(subject, new EmailResponseConfig() { oApp = oApp, replyRecipients = avalcol.Value == null ? t.sucEmails : t.rejEmails, template = avalcol.Value == null ? t.sucTemplate : t.rejTemplate, rows = avalcol.Value, sentonbehalf = configs.ContainsKey(EPconfigsEnum.sentonbehalf.ToString()) ? configs[EPconfigsEnum.sentonbehalf.ToString()] : null, ResultMap = t.ResultMap });


                        if (!success)
                        {
                            moveto = t.rejFolder;
                            if (moveto == null)
                                clearCat = true;
                        }

                        Logger.WriteToConsole($"Sending response template [{(success ? t.sucTemplate : t.rejTemplate)}] for  [{subject.Subject}]");

                    }

                else if (resultsCol?.Any() == true)
                {
                    new EmailResponseController().execute(subject, new EmailResponseConfig() { oApp = oApp, template = configs.ContainsKey(EPconfigsEnum.sucTemplate.ToString()) ? new FileInfo(configs[EPconfigsEnum.sucTemplate.ToString()]) : null, sentonbehalf = configs.ContainsKey(EPconfigsEnum.sentonbehalf.ToString()) ? configs[EPconfigsEnum.sentonbehalf.ToString()] : null, savesentfolder = configs.ContainsKey(EPconfigsEnum.saveSentFolder.ToString()) ? configs[EPconfigsEnum.saveSentFolder.ToString()] : null });
                    Logger.WriteToConsole($"Sending response success template for  [{subject.Subject}]");
                }
                else if(success)
                {
                    movetoinbox = true;
                    Logger.WriteToConsole($"Moving to [{configs[EPconfigsEnum.retfolder.ToString()]}] due to having nothing to load...");

                }
                       else Logger.WriteToConsole($"Unknown error occurred for [{subject.Subject}] ...");
                


            }
        } else movetoinbox = true;

        if (saveChanges)
        {
            if (movetoinbox) { clearCat = true; moveto = configs[EPconfigsEnum.retfolder.ToString()]; }

            if (clearCat) { subject.Categories = null; subject.Save(); }

            if (moveto != null && moveto != configs[EPconfigsEnum.destFolder.ToString()])
            {
                var movetohandle = oApp.Session.Stores[configs[EPconfigsEnum.storename.ToString()]].GetRootFolder().Folders[moveto];
                subject.Move(movetohandle);
                Marshal.ReleaseComObject(movetohandle);
                movetohandle = null;
                Logger.WriteToConsole($"Moved [{subject.Subject}] to [{moveto}]");
            }

        }


        return success;
    }

    protected override string GetRelativeId(string id)
        {
            return Path.GetFileName(id);
        }
    }

