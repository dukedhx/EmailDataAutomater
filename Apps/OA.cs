using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClassLibrary1.Utils;
using ClassLibrary1.Utils.Persistence;
using ConsoleApp1.Components.Consolidators;
using ConsoleApp1.Components.Processors;
using ConsoleApp1.Entities.Control;
using ConsoleApp1.Utils;
using YamlDotNet.RepresentationModel;

namespace ConsoleApp1.Apps
{
    internal class OA
    {
   
        public static DateTime StartTime,LastRun,LastRunFinished;
        public static Exception Lasterror;
        public static String Currentdataresultpath,LastEmailProcessed,LastEmailLoaded,CurrentEmailSavePath;
        public static int LastStandardRun;
        public static int LastForcedRun;

        public static void newStartOAHelper()
        {
            try
            {
                DirectoryInfo di = new DirectoryInfo("temp");

                foreach (FileInfo file in di.GetFiles())
                    try
                    {
                        file.Delete();
                    }
                    catch { }
                foreach (DirectoryInfo dir in di.GetDirectories())
                    try
                    {
                        dir.Delete(true);
                    }
                    catch { }
            }
            catch (Exception ex)
            {
                Logger.Log("[Init Temp Cleaning Error]" + ex);
            }
          //  DATARSTHeader.Reset();
            Logger.Log("Loading data mapping...");
            var yaml = new YamlStream();
            var headersID = Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["dataResultMappingPath"]);
            if (!DATARSTHeader.headers.ContainsKey(headersID))
            {
                using (var sr = new FileInfo(headersID).OpenText())

                    yaml.Load(sr);
                (yaml.Documents[0].RootNode as YamlMappingNode).Children.Where(e => !ConsolidateHelper.isIgnoredHeader(e.Value?.ToString())).Select(e => new DATARSTHeader(e.Value.ToString(), headersID, DATARSTHeader.headers.GetValueOrDefault(headersID)?.FirstOrDefault(h => h.name == e.Value.ToString())?.value ?? 0)).ToArray();
            }
            var headers = DATARSTHeader.headers[headersID];
            Dictionary<String, String> ResultMap = null;
            using (var sr = new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["emailresponseMappingPath"])).OpenText())

                yaml.Load(sr);

            ResultMap = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.Select(entry => new { Key = entry.Key.ToString(), Value = entry.Value.ToString() }).ToDictionary(prop => prop.Key, prop => prop.Value);
            String vcpath = Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["IAVFValidationConfigPath"]);

            yaml.Load(new FileInfo(vcpath).OpenText());
            var vfvalidator = new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], null, ConfigurationManager.Configuration["IAWorkSheetName"], new XlsxCellValidator()) {  new EmailValidatorConfig(new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["vfreminderMappingPath"]))) { continueOnReject = true, rejectOnInvalid = false, criteria= ((YamlMappingNode)yaml.Documents[0].RootNode).Children.ToDictionary(prop=>prop.Key.ToString(),prop=>prop.Value.ToString()), ID = vcpath, rejTemplate = new FileInfo(ConfigurationManager.Configuration["IA_Template_Success_ReminderPath"]) ,useCustomValidation=true} };

            vcpath = Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["validation0509ConfigPath"]);
            yaml.Load(new FileInfo(vcpath).OpenText());

            var svalidator = new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersion"], ConfigurationManager.Configuration["IAWorkSheetName"]) { new EmailValidatorConfig { ID=vcpath, criteria = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.ToDictionary(prop => prop.Key.ToString(), prop => prop.Value.ToString()), rejFolder = ConfigurationManager.Configuration["FADataRejFolder"], rejTemplate = new FileInfo(ConfigurationManager.Configuration["IA_Template_RejectPath"]), ResultMap = ResultMap } };


            vcpath = Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["validationConfigPath"]);
            yaml.Load(new FileInfo(vcpath).OpenText());

            var old_svalidator = new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersionOLD"], ConfigurationManager.Configuration["IAWorkSheetName"]) { new EmailValidatorConfig {ID=vcpath, criteria = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.ToDictionary(prop => prop.Key.ToString(), prop => prop.Value.ToString()), rejFolder = ConfigurationManager.Configuration["FADataRejFolder"], rejTemplate = new FileInfo(ConfigurationManager.Configuration["IA_Template_RejectPath"]), ResultMap = ResultMap } };
            var fvalds = new EmailValidatorConfigCol[] { svalidator, old_svalidator, vfvalidator };

            vcpath = Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["attention0509ConfigPath"]);
            yaml.Load(new FileInfo(vcpath).OpenText());

            var attn = new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersion"], ConfigurationManager.Configuration["IAWorkSheetName"]){ new EmailValidatorConfig { ID=vcpath, criteria = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.ToDictionary(prop => prop.Key.ToString(), prop => prop.Value.ToString()), rejFolder = ConfigurationManager.Configuration["FADataUrgentFolder"], rejEmails = ConfigurationManager.Configuration["IAurgentNotification"]?.Split(";"), rejTemplate = new FileInfo(ConfigurationManager.Configuration["IAurgentNotificationTemplate"]), ResultMap = ResultMap } };

            vcpath = Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["attentionConfigPath"]);
            yaml.Load(new FileInfo(vcpath).OpenText());

            var old_attn =  new EmailValidatorConfigCol(ConfigurationManager.Configuration["IAMacroSnippet"], ConfigurationManager.Configuration["IATemplateVersionOLD"], ConfigurationManager.Configuration["IAWorkSheetName"]) { new EmailValidatorConfig { ID=vcpath, criteria = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.ToDictionary(prop => prop.Key.ToString(), prop => prop.Value.ToString()), rejFolder = ConfigurationManager.Configuration["FADataUrgentFolder"], rejEmails = ConfigurationManager.Configuration["IAurgentNotification"]?.Split(";"), rejTemplate = new FileInfo(ConfigurationManager.Configuration["IAurgentNotificationTemplate"]), ResultMap = ResultMap } };

            var svalds = new EmailValidatorConfigCol[] {  svalidator, old_svalidator,attn, old_attn, vfvalidator };
            var today = DateTime.Now.ToString("s").Split("T")[0];
            var dparts = today.Split("-");
            Currentdataresultpath = Path.Combine(ConfigurationManager.Configuration["FADataResultPath"], dparts[0], dparts[1], dparts[2], $"{ today}.(OpenReadOnly).xlsx");

            CurrentEmailSavePath= Path.Combine(ConfigurationManager.Configuration["FADataResultEmailPath"], dparts[0], dparts[1], dparts[2]);
            try
            {

                //   LastStandardRun = OutlookHelper.ConsolidateEmail(ConfigurationManager.Configuration["FADataStoreName"], ConfigurationManager.Configuration["FADataInFolder"], ConfigurationManager.Configuration["FADataInFolder"], null, Currentdataresultpath, svalds, ConfigurationManager.Configuration["FADataSenderAddress"], DATARSTHeader.headers, new FileInfo(ConfigurationManager.Configuration["IA_Template_SuccessPath"]), ConfigurationManager.Configuration["IAOItemsRestricterString"]);


                new MailStoreProcessor().process(new MailStoreConfig() { storename = ConfigurationManager.Configuration["FADataStoreName"], validColsCol = fvalds, dpath = Currentdataresultpath, headers = headers, infolder = ConfigurationManager.Configuration["FADataForceFolder"], savemailpath = CurrentEmailSavePath, sucfolder = ConfigurationManager.Configuration["FADataInFolder"], sentonbehalf = ConfigurationManager.Configuration["FADataSenderAddress"], sucTemplate = new FileInfo(ConfigurationManager.Configuration["IA_Template_SuccessPath"]), rejfolder = ConfigurationManager.Configuration["FADataRejFolder"], restricter = ConfigurationManager.Configuration["IAOItemsRestricterString"] }, null, new Dictionary<string, string>() { { EPconfigsEnum.dataResultMappingPath.ToString() , ConfigurationManager.Configuration["dataResultMappingPath"] },

                    { EPconfigsEnum.dataResultValMapping.ToString() , ConfigurationManager.Configuration["dataResultValMapping"] },
                    { EPconfigsEnum.AdminEmail.ToString() , ConfigurationManager.Configuration[EPconfigsEnum.AdminEmail.ToString()] },
                    { PSTconfigsEnum.pstPath.ToString(), ConfigurationManager.Configuration["pstPath"] },
                    { EPconfigsEnum.IAFormatFields.ToString() , ConfigurationManager.Configuration[EPconfigsEnum.IAFormatFields.ToString()] },
                    { PSTconfigsEnum.template.ToString(), ConfigurationManager.Configuration["IAQueryTemplate"] },
                       { EPconfigsEnum.emailLogTemplate.ToString(), ConfigurationManager.Configuration["optIAlogTemplate"] },
                    { XCDconfigsEnum.headersID.ToString(),headersID},
                    { EPconfigsEnum.emailLogConnString.ToString(), ConfigurationManager.Configuration["optlogdbconn"] },
                    { PSTconfigsEnum.pstConnString.ToString(), ConfigurationManager.Configuration["iatdbconn"] }

                });
                
                new MailStoreProcessor() .process(new MailStoreConfig() {  storename= ConfigurationManager.Configuration["FADataStoreName"], validColsCol = svalds,dpath= Currentdataresultpath , headers= headers, infolder= ConfigurationManager.Configuration["FADataInFolder"] , savemailpath= CurrentEmailSavePath, sucfolder = ConfigurationManager.Configuration["FADataInFolder"] ,sentonbehalf = ConfigurationManager.Configuration["FADataSenderAddress"] ,sucTemplate= new FileInfo(ConfigurationManager.Configuration["IA_Template_SuccessPath"]),rejfolder= ConfigurationManager.Configuration["FADataRejFolder"], restricter= ConfigurationManager.Configuration["IAOItemsRestricterString"] },null,new Dictionary<string,string>() { { EPconfigsEnum.dataResultMappingPath.ToString() , ConfigurationManager.Configuration["dataResultMappingPath"] },

                    { EPconfigsEnum.dataResultValMapping.ToString() , ConfigurationManager.Configuration["dataResultValMapping"] },
                    { EPconfigsEnum.AdminEmail.ToString() , ConfigurationManager.Configuration[EPconfigsEnum.AdminEmail.ToString()] },
                    { EPconfigsEnum.IAFormatFields.ToString() , ConfigurationManager.Configuration[EPconfigsEnum.IAFormatFields.ToString()] },
                    { PSTconfigsEnum.pstPath.ToString(), ConfigurationManager.Configuration["pstPath"] },
                    { PSTconfigsEnum.pstConnString.ToString(), ConfigurationManager.Configuration["iatdbconn"] },
                    { EPconfigsEnum.emailLogTemplate.ToString(), ConfigurationManager.Configuration["optIAlogTemplate"] },
                    { EPconfigsEnum.emailLogConnString.ToString(), ConfigurationManager.Configuration["optlogdbconn"] },
                    { XCDconfigsEnum.headersID.ToString(),headersID},
                    { PSTconfigsEnum.template.ToString(), ConfigurationManager.Configuration["IAQueryTemplate"] }

                });
            }
            catch(Exception ex)
            {
                ex.ToString();
            }

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
