using ClassLibrary1.Utils;
using ClassLibrary1.Utils.Persistence;
using ConsoleApp1.Components.Interfaces;
using ConsoleApp1.Entities.Control;
using ConsoleApp1.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ConsoleApp1.Components.Controllers
{
    public class EmailResponseController : IControl<Outlook.MailItem, EmailResponseConfig, Boolean>
    {

        protected String getHtmlBodyFromTemplate(EmailResponseConfig configCol)
        {
            String htmlbody = null, mmsg = configCol.rows == null ? configCol.defaultMessage : $"<table border='1'>{String.Join(Environment.NewLine, configCol.rows.Select(r => $"<tr><td>{r.Item1}</td><td>{r.Item2}</td></tr>"))}</table>";


            if (configCol.template?.Exists == true)
            {

                var tpath = Path.Combine(Tools.GetExecutingPath, $"temp\\{Guid.NewGuid()}{Path.GetExtension(configCol.template.Name)}");
                configCol.template.CopyTo(tpath);
                var template = (Outlook.MailItem)configCol.oApp.Session.OpenSharedItem(tpath);
                htmlbody = template.HTMLBody;
                template.Close(Outlook.OlInspectorClose.olDiscard);
                Marshal.ReleaseComObject(template);
                template = null;
                File.Delete(tpath);
            }

            if (configCol.ResultMap != null)
                foreach (var amp in configCol.ResultMap)
                    mmsg = mmsg.Replace(amp.Key, amp.Value);

            return htmlbody?.Replace("{{message}}", mmsg) ?? mmsg;
        }

        public bool execute(Outlook.MailItem subject, EmailResponseConfig configCol, IDictionary<string, string> configs = null)
        {
            Outlook.MailItem reply = null;
            reply = subject.Reply();
            var sbc = subject.SenderEmailAddress;

            if (configCol.replyRecipients == null)
            {
                reply = subject.Reply();
                try
                {
                  //  var ignccs = ConfigurationManager.Configuration["IgnoredCCAddress"]?.Split(";");
                    for (int i = 1; i <= subject.Recipients.Count; i++)
                    {
                        var recip = subject.Recipients[i];
                        if (recip.Type == (int)Outlook.OlMailRecipientType.olCC )

                            reply.CC += (reply.CC?.Length > 0 ? ";" : "") + recip.Address;

                    }

                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }
            }
            else
            {
                reply = subject.Forward();
                foreach (var e in configCol.replyRecipients.Where(e => OutlookHelper.EmailAddrRegex.IsMatch(e)))
                    reply.Recipients.Add(e);
            }
            if (configCol.ccRecipients != null)
                reply.BCC = OutlookHelper. GetEmailString(configCol.ccRecipients);
            Logger.WriteToConsole($"Prepared validation response for [{reply.Subject}]");


            reply.HTMLBody = getHtmlBodyFromTemplate( configCol);
           


            var subj = reply.Subject;
            reply.SentOnBehalfOfName = configCol.sentonbehalf;


            if (!reply.Recipients.ResolveAll()) reply.CC = "";
            reply.Send();
            Logger.WriteToConsole($"Sent [{subj}]");
            Marshal.ReleaseComObject(reply);
            reply = null;
            return true;
        }
    }
}
