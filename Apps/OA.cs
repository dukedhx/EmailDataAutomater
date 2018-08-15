using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using ClassLibrary1.Utils;
using ClassLibrary1.Utils.Persistence;
using ConsoleApp1.Components.Consolidators;
using ConsoleApp1.Components.Controllers;
using ConsoleApp1.Components.Contollers;
using Outlook = Microsoft.Office.Interop.Outlook;
using ConsoleApp1.Entities.Control;
using ConsoleApp1.Utils;
using YamlDotNet.RepresentationModel;
using System.Runtime.InteropServices;
using ConsoleApp1.Utils.Helpers;
using ClassLibrary1.Utils.Misc;
using System.Threading.Tasks;
using ConsoleApp1.Util;

namespace ConsoleApp1.Apps
{
    public class OA
    {
   
        public static DateTime StartTime,LastRun,LastRunFinished;
        public static Exception Lasterror;
        public static String Currentdataresultpath,LastEmailProcessed,LastEmailLoaded,CurrentEmailSavePath;
        public static int LastStandardRun;
        public static int LastForcedRun;
        public static void sendEmail(Dictionary<String,String>args)
        {
            var oApp = new Outlook.Application();
            var oNS = MailStoreProcessor.logOn(oApp);
            IEnumerable<Tuple<string, string, bool?>> rows = null;
            switch (args.GetValueOrDefault("content"))
            {
                case "json":rows = DataTools.JsonToDict(args["text"]).Select(e => new Tuple<string, string, bool?>(e.Key,e.Value,null)); break;
                case "comma":rows = args["text"]?.Split(",")?.Select(e=>new Tuple<string,string,bool?>("",e,null)); break;
                
            }
            new EmailResponseController().execute(null, new EmailResponseConfig() {attachments=DataTools.getYamlArray( args.GetValueOrDefault("attachments")).ToArray(), oApp = oApp, template = args.ContainsKey("template")?new FileInfo(args["template"]):null, sentonbehalf = args.GetValueOrDefault("sentonbehalf")??ConfigurationManager.Configuration["DefaultSenderAddress"], emailSubject= args.GetValueOrDefault("subject"), ResultMap= args.ContainsKey("resultMap") ? ConfigTools.YamlToDict(args["resultMap"]) : null ,rows=rows,defaultMessage= args["text"] , replyRecipients=  args["recipients"].Split(";")  });
            Logger.WriteToConsole($"Sending response success template for  [{args["subject"]}]");

           

            if (oNS != null)
            {
                oNS.Logoff();
                Marshal.ReleaseComObject(oNS);

            }
            if (oApp != null) Marshal.ReleaseComObject(oApp);

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }


        public static void StartOAHelper(String path) {

            ProperNameController pnc = new ProperNameController();
            Task.WaitAll(
            Directory.GetFiles(Directory.CreateDirectory(path).FullName, "*.yaml").Select(e =>

                Task.Factory.StartNew(() => {
                    try
                    {



                        var configdict = YamlTools.getResultMapFromYaml(e);
                        foreach (String config in DataTools.getYamlArray(configdict.GetValueOrDefault("configs")))
                        {

                            var aconfigdict = YamlTools.getResultMapFromYaml(config);
                            aconfigdict.ToArray();
                            List<EmailValidatorConfigCol> emailValidatorConfigCols = new List<EmailValidatorConfigCol>();
                            foreach (String consolidation in DataTools.getYamlArray(aconfigdict.GetValueOrDefault("consolidationConfigs")))
                            {


                                var aconsolidationConfig = YamlTools.getNodeMapFromYaml(consolidation);
                                var aconsolidation = YamlTools.getResultMapFromYaml(consolidation);
                                foreach (var vaddic in aconsolidationConfig["validation"].AllNodes.Where(n => n.NodeType == YamlNodeType.Mapping).Select(n => ((YamlMappingNode)n).Children.Select(entry => new KeyValuePair<string, string>(entry.Key.ToString(), entry.Value.ToString())).ToDictionary(prop => prop.Key, prop => prop.Value)))
                                {

                                                       //vaddic.ToArray();

                                                       emailValidatorConfigCols.Add(new EmailValidatorConfigCol(vaddic["macroSnippet"], vaddic["version"], vaddic.GetValueOrDefault("sheetName"))
                                    {
                                                            new EmailValidatorConfig { ID = vaddic["path"], criteria =  YamlTools.getResultMapFromYaml(vaddic["path"]), rejFolder = aconsolidation.GetValueOrDefault("rejfolder")??aconfigdict.GetValueOrDefault("rejFolder"), rejTemplate =  new FileInfo(aconsolidation["rejTemplate"]), rejEmails=aconfigdict.GetValueOrDefault("rejRecipients")?.Split(";"), ResultMap = YamlTools.getResultMapFromYaml(aconfigdict.GetValueOrDefault("emailresponseMappingPath")),continueOnReject= aconsolidation.GetValueOrDefault("continueOnReject")=="true",rejectOnInvalid=aconsolidation.GetValueOrDefault("rejectOnInvalid")=="true",sucEmails=aconfigdict.GetValueOrDefault("sucRecipients")?.Split(";"),sucTemplate=String.IsNullOrWhiteSpace(aconsolidation.GetValueOrDefault("successTemplate"))?null:new FileInfo(aconsolidation["successTemplate"]),useCustomValidation=vaddic.GetValueOrDefault("useCustomValidation")=="true"}
                                    });
                                }

                            }

                            new MailStoreProcessor() { pnc = pnc }.process(new MailStoreConfig() { storename = configdict["mailStore"], rejfolder = aconfigdict["rejFolder"], savemailpath = aconfigdict.GetValueOrDefault("savemailpath") ?? configdict.GetValueOrDefault("savemailpath"), infolder = aconfigdict["inFolder"], sucfolder = aconfigdict["sucFolder"], validColsCol = emailValidatorConfigCols, sentonbehalf = aconfigdict.GetValueOrDefault("sentonbehalf") ?? configdict.GetValueOrDefault("sentonbehalf"), sucTemplate = new FileInfo(aconfigdict["successTemplate"]), retfolder = aconfigdict.GetValueOrDefault("returnFolder") ?? configdict.GetValueOrDefault("returnFolder"), restricter = aconfigdict["restricter"] ?? configdict.GetValueOrDefault("restricter") }, null, null);




                        }

                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"[{e}]{ex}");
                        Lasterror = ex;
                        Lasterror.HelpLink = e;
                    }

                })
            ).ToArray());

        }


        public static void newStartOAHelper()
        {
          //  try
          //  {
          //      DirectoryInfo di = new DirectoryInfo("temp");

          //      foreach (FileInfo file in di.GetFiles())
          //          try
          //          {
          //              file.Delete();
          //          }
          //          catch { }
          //      foreach (DirectoryInfo dir in di.GetDirectories())
          //          try
          //          {
          //              dir.Delete(true);
          //          }
          //          catch { }
          //  }
          //  catch (Exception ex)
          //  {
          //      Logger.Log("[Init Temp Cleaning Error]" + ex);
          //  }
          ////  DATARSTHeader.Reset();
          //  Logger.Log("Loading data mapping...");
          //  var yaml = new YamlStream();
          //  var headersID = Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["dataResultMappingPath"]);
          //  if (!DATARSTHeader.headers.ContainsKey(headersID))
          //  {
          //      using (var sr = new FileInfo(headersID).OpenText())

          //          yaml.Load(sr);
          //      (yaml.Documents[0].RootNode as YamlMappingNode).Children.Where(e => !ConsolidateHelper.isIgnoredHeader(e.Value?.ToString())).Select(e => new DATARSTHeader(e.Value.ToString(), headersID, DATARSTHeader.headers.GetValueOrDefault(headersID)?.FirstOrDefault(h => h.name == e.Value.ToString())?.value ?? 0)).ToArray();
          //  }
          //  var headers = DATARSTHeader.headers[headersID];
          //  Dictionary<String, String> ResultMap = null;
          //  using (var sr = new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["emailresponseMappingPath"])).OpenText())

          //      yaml.Load(sr);

          //  ResultMap = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.Select(entry => new { Key = entry.Key.ToString(), Value = entry.Value.ToString() }).ToDictionary(prop => prop.Key, prop => prop.Value);
          //  String vcpath = Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["IAVFValidationConfigPath"]);

          //  yaml.Load(new FileInfo(vcpath).OpenText());
          //  var vfvalidator = new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersion"], ConfigurationManager.Configuration["IAWorkSheetName"], new XlsxCellValidator()) {  new EmailValidatorConfig(new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["vfreminderMappingPath"]))) { continueOnReject = true, rejectOnInvalid = false, criteria= ((YamlMappingNode)yaml.Documents[0].RootNode).Children.ToDictionary(prop=>prop.Key.ToString(),prop=>prop.Value.ToString()), ID = vcpath, rejTemplate = new FileInfo(ConfigurationManager.Configuration["IA_Template_Success_ReminderPath"]) ,useCustomValidation=true} };

          //  var ovfvalidator = new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersionOLD"], ConfigurationManager.Configuration["IAWorkSheetName"], new XlsxCellValidator()) { new EmailValidatorConfig(new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["vfreminderMappingPath"]))) { continueOnReject = true, rejectOnInvalid = false, criteria = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.ToDictionary(prop => prop.Key.ToString(), prop => prop.Value.ToString()), ID = vcpath, rejTemplate = new FileInfo(ConfigurationManager.Configuration["IA_Template_Success_ReminderPath"]), useCustomValidation = true } };


          //  vcpath = Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["validation0509ConfigPath"]);
          //  yaml.Load(new FileInfo(vcpath).OpenText());

          //  var svalidator = new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersion"], ConfigurationManager.Configuration["IAWorkSheetName"]) { new EmailValidatorConfig { ID=vcpath, criteria = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.ToDictionary(prop => prop.Key.ToString(), prop => prop.Value.ToString()), rejFolder = ConfigurationManager.Configuration["FADataRejFolder"], rejTemplate = new FileInfo(ConfigurationManager.Configuration["IA_Template_RejectPath"]), ResultMap = ResultMap } };


          //  vcpath = Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["validationConfigPath"]);
          //  yaml.Load(new FileInfo(vcpath).OpenText());

          //  var old_svalidator = new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersionOLD"], ConfigurationManager.Configuration["IAWorkSheetName"]) { new EmailValidatorConfig {ID=vcpath, criteria = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.ToDictionary(prop => prop.Key.ToString(), prop => prop.Value.ToString()), rejFolder = ConfigurationManager.Configuration["FADataRejFolder"], rejTemplate = new FileInfo(ConfigurationManager.Configuration["IA_Template_RejectPath"]), ResultMap = ResultMap } };
          //  var fvalds = new EmailValidatorConfigCol[] { svalidator, old_svalidator };

          //  vcpath = Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["attention0509ConfigPath"]);
          //  yaml.Load(new FileInfo(vcpath).OpenText());

          //  var attn = new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersion"], ConfigurationManager.Configuration["IAWorkSheetName"]){ new EmailValidatorConfig { ID=vcpath, criteria = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.ToDictionary(prop => prop.Key.ToString(), prop => prop.Value.ToString()), rejFolder = ConfigurationManager.Configuration["FADataUrgentFolder"], rejEmails = ConfigurationManager.Configuration["IAurgentNotification"]?.Split(";"), rejTemplate = new FileInfo(ConfigurationManager.Configuration["IAurgentNotificationTemplate"]), ResultMap = ResultMap } };

          //  vcpath = Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["attentionConfigPath"]);
          //  yaml.Load(new FileInfo(vcpath).OpenText());

          //  var old_attn =  new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersionOLD"], ConfigurationManager.Configuration["IAWorkSheetName"]) { new EmailValidatorConfig { ID=vcpath, criteria = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.ToDictionary(prop => prop.Key.ToString(), prop => prop.Value.ToString()), rejFolder = ConfigurationManager.Configuration["FADataUrgentFolder"], rejEmails = ConfigurationManager.Configuration["IAurgentNotification"]?.Split(";"), rejTemplate = new FileInfo(ConfigurationManager.Configuration["IAurgentNotificationTemplate"]), ResultMap = ResultMap } };

          //  var svalds = new EmailValidatorConfigCol[] {  svalidator, old_svalidator,attn, old_attn };
          //  var today = DateTime.Now.ToString("s").Split("T")[0];
          //  var dparts = today.Split("-");
          //  Currentdataresultpath = Path.Combine(ConfigurationManager.Configuration["FADataResultPath"], dparts[0], dparts[1], dparts[2], $"{ today}.(OpenReadOnly).xlsx");

          //  CurrentEmailSavePath= Path.Combine(ConfigurationManager.Configuration["FADataResultEmailPath"], dparts[0], dparts[1], dparts[2]);
          //  try
          //  {

          //      //   LastStandardRun = OutlookHelper.ConsolidateEmail(ConfigurationManager.Configuration["FADataStoreName"], ConfigurationManager.Configuration["FADataInFolder"], ConfigurationManager.Configuration["FADataInFolder"], null, Currentdataresultpath, svalds, ConfigurationManager.Configuration["FADataSenderAddress"], DATARSTHeader.headers, new FileInfo(ConfigurationManager.Configuration["IA_Template_SuccessPath"]), ConfigurationManager.Configuration["IAOItemsRestricterString"]);


          //      new MailStoreProcessor().process(new MailStoreConfig() { storename = ConfigurationManager.Configuration["FADataStoreName"], validColsCol = fvalds,  infolder = ConfigurationManager.Configuration["FADataForceFolder"], savemailpath = CurrentEmailSavePath, sucfolder = ConfigurationManager.Configuration["FADataInFolder"], sentonbehalf = ConfigurationManager.Configuration["FADataSenderAddress"], sucTemplate = new FileInfo(ConfigurationManager.Configuration["IA_Template_SuccessPath"]), rejfolder = ConfigurationManager.Configuration["FADataRejFolder"], restricter = ConfigurationManager.Configuration["IAOItemsRestricterString"] }, null, new Dictionary<string, string>() { { EPconfigsEnum.dataResultMappingPath.ToString() , ConfigurationManager.Configuration["dataResultMappingPath"] },

          //          { EPconfigsEnum.dataResultValMapping.ToString() , ConfigurationManager.Configuration["dataResultValMapping"] },
          //          { EPconfigsEnum.AdminEmail.ToString() , ConfigurationManager.Configuration[EPconfigsEnum.AdminEmail.ToString()] },
          //          { PSTconfigsEnum.pstPath.ToString(), ConfigurationManager.Configuration["pstPath"] },
          //          { EPconfigsEnum.IAFormatFields.ToString() , ConfigurationManager.Configuration[EPconfigsEnum.IAFormatFields.ToString()] },
          //          { PSTconfigsEnum.template.ToString(), ConfigurationManager.Configuration["IAQueryTemplate"] },
          //             { EPconfigsEnum.emailLogTemplate.ToString(), ConfigurationManager.Configuration["optIAlogTemplate"] },
          //          { XCDconfigsEnum.headersID.ToString(),headersID},
          //          { EPconfigsEnum.emailLogConnString.ToString(), ConfigurationManager.Configuration["optlogdbconn"] },
          //          { PSTconfigsEnum.pstConnString.ToString(), ConfigurationManager.Configuration["iatdbconn"] }

          //      });
                
          //      new MailStoreProcessor() .process(new MailStoreConfig() {  storename= ConfigurationManager.Configuration["FADataStoreName"], validColsCol = svalds, infolder= ConfigurationManager.Configuration["FADataInFolder"] , savemailpath= CurrentEmailSavePath, sucfolder = ConfigurationManager.Configuration["FADataInFolder"] ,sentonbehalf = ConfigurationManager.Configuration["FADataSenderAddress"] ,sucTemplate= new FileInfo(ConfigurationManager.Configuration["IA_Template_SuccessPath"]),rejfolder= ConfigurationManager.Configuration["FADataRejFolder"], restricter= ConfigurationManager.Configuration["IAOItemsRestricterString"] },null,new Dictionary<string,string>() { { EPconfigsEnum.dataResultMappingPath.ToString() , ConfigurationManager.Configuration["dataResultMappingPath"] },

          //          { EPconfigsEnum.dataResultValMapping.ToString() , ConfigurationManager.Configuration["dataResultValMapping"] },
          //          { EPconfigsEnum.AdminEmail.ToString() , ConfigurationManager.Configuration[EPconfigsEnum.AdminEmail.ToString()] },
          //          { EPconfigsEnum.IAFormatFields.ToString() , ConfigurationManager.Configuration[EPconfigsEnum.IAFormatFields.ToString()] },
          //          { PSTconfigsEnum.pstPath.ToString(), ConfigurationManager.Configuration["pstPath"] },
          //          { PSTconfigsEnum.pstConnString.ToString(), ConfigurationManager.Configuration["iatdbconn"] },
          //          { EPconfigsEnum.emailLogTemplate.ToString(), ConfigurationManager.Configuration["optIAlogTemplate"] },
          //          { EPconfigsEnum.emailLogConnString.ToString(), ConfigurationManager.Configuration["optlogdbconn"] },
          //          { XCDconfigsEnum.headersID.ToString(),headersID},
          //          { PSTconfigsEnum.template.ToString(), ConfigurationManager.Configuration["IAQueryTemplate"] }

          //      });
          //  }
          //  catch(Exception ex)
          //  {
          //      ex.ToString();
          //  }

        }
        //    public static void StartOAHelper()
        //{
          
        //    try
        //    {
        //        DirectoryInfo di = new DirectoryInfo("temp");

        //        foreach (FileInfo file in di.GetFiles())
        //            try
        //            {
        //                file.Delete();
        //            }
        //            catch { }
        //        foreach (DirectoryInfo dir in di.GetDirectories())
        //            try
        //            {
        //                dir.Delete(true);
        //            }
        //            catch { }
        //    }catch(Exception ex)
        //    {
        //        Logger.Log("[Init Temp Cleaning Error]"+ex);
        //    }

            
        //        DATARSTHeader.Reset();
        //        Logger.Log("Loading data mapping...");
        //        var yaml = new YamlStream();

        //        using (var sr = new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["dataResultMappingPath"])).OpenText())

        //            yaml.Load(sr);
        //        (yaml.Documents[0].RootNode as YamlMappingNode).Children.Where(e => !ConsolidateHelper.isIgnoredHeader(e.Value?.ToString())).Select(e => new DATARSTHeader(e.Value.ToString())).ToArray();

        //    Dictionary<String, String> ResultMap = null;
        //    using (var sr = new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["emailresponseMappingPath"])).OpenText())

        //        yaml.Load(sr);

        //     ResultMap = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.Select(entry => new { Key = entry.Key.ToString(), Value = entry.Value.ToString() }).ToDictionary(prop => prop.Key, prop => prop.Value);


        //    var vfvalidator = new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersion"], ConfigurationManager.Configuration["IAWorkSheetName"], new XlsxCellValidator()) { new EmailValidatorConfig(new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["vfreminderMappingPath"]))) { continueOnReject = true, rejectOnInvalid = false, fileInfo = new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "",  ConfigurationManager.Configuration["IAVFValidationConfigPath"])), rejTemplate = new FileInfo(ConfigurationManager.Configuration["IA_Template_Success_ReminderPath"]) } } ;

        //    var svalidator = new EmailValidatorConfig { fileInfo = new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["validationConfigPath"])), rejFolder = ConfigurationManager.Configuration["FADataRejFolder"], rejTemplate = new FileInfo(ConfigurationManager.Configuration["IA_Template_RejectPath"]), ResultMap = ResultMap } ;
        //    var fvalds=new EmailValidatorConfigCol[] { new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersion"], ConfigurationManager.Configuration["IAWorkSheetName"]) { svalidator  }, vfvalidator };
        //    var svalds = new EmailValidatorConfigCol[] {  new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersion"], ConfigurationManager.Configuration["IAWorkSheetName"]) { svalidator, new EmailValidatorConfig { fileInfo = new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["attentionConfigPath"])), rejFolder = ConfigurationManager.Configuration["FADataUrgentFolder"], rejEmails =   ConfigurationManager.Configuration["IAurgentNotification"]?.Split(";")  , rejTemplate= new FileInfo(ConfigurationManager.Configuration["IAurgentNotificationTemplate"]), ResultMap = ResultMap } }, vfvalidator };
        //    try
        //    {

        //        Logger.Log("Outlook Helper Started ...");
        //                var today = DateTime.Now.ToString("s").Split("T")[0];
        //                var dparts = today.Split("-");
        //                Currentdataresultpath = Path.Combine(ConfigurationManager.Configuration["FADataResultEmailPath"], dparts[0], dparts[1], dparts[2], $"{ today}.(OpenReadOnly).xlsx");
        //                LastRun = DateTime.Now;
        //              LastForcedRun=   OutlookHelper.ConsolidateEmail(ConfigurationManager.Configuration["FADataStoreName"], ConfigurationManager.Configuration["FADataForceFolder"], ConfigurationManager.Configuration["FADataInFolder"], ConfigurationManager.Configuration["FADataReturnFolder"], Currentdataresultpath,  fvalds , ConfigurationManager.Configuration["FADataSenderAddress"], DATARSTHeader.headers,new FileInfo(ConfigurationManager.Configuration["IA_Template_SuccessPath"]), ConfigurationManager.Configuration["IAOItemsRestricterString"]) ;
        //               LastStandardRun = OutlookHelper.ConsolidateEmail(ConfigurationManager.Configuration["FADataStoreName"], ConfigurationManager.Configuration["FADataInFolder"], ConfigurationManager.Configuration["FADataInFolder"], null, Currentdataresultpath, svalds, ConfigurationManager.Configuration["FADataSenderAddress"], DATARSTHeader.headers, new FileInfo(ConfigurationManager.Configuration["IA_Template_SuccessPath"]), ConfigurationManager.Configuration["IAOItemsRestricterString"]);

                       
        //    }
        //    catch(Exception ex)
        //    {
        //       // OA.Started = false;
        //        Logger.WriteToConsole(ex);
        //        Lasterror = ex;
        //    }

        //}

    }
}
