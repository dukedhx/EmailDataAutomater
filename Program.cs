using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Linq;
using ConsoleApp1.Apps;
using ConsoleApp1.Utils;
using ClassLibrary1.Utils.Persistence;
using System.Collections.Generic;
using YamlDotNet.RepresentationModel;

namespace ConsoleApp1
{
     class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            ConfigurationManager.Build(new string[] {  "appsettings.json", "appsettings.development.json" },Directory.GetCurrentDirectory());
            //var sb = ConfigurationManager.Configuration["PersistToMySQL"];
          //  args =new string[]{ "app=sendEmail", @"template=\\fnp1423mp01.lowes.com\Data1\SHARE\EVERYONE\C-QA\Operation Team\Email Template\IA_Template_Reject.msg","to=bryan.huang@lowes.com", "content=json", "sentonbehalf=lgsqadata@lowes.com", @"text={""Error1"":""F3"",""Error2"":""Factory is not inspection allow""}","subject=233", @"resultMap=\\fnp1423mp01.lowes.com\Data1\SHARE\EVERYONE\C-QA\Operation Team\Email Template\invalidStatusMapping.yaml" };
            var argsMap = args.Select(e=>e.Split("=")).Where(e=>!String.IsNullOrWhiteSpace( e[0])).ToDictionary(prop=>prop[0],prop=>prop.Length>1?String.Join('=', prop.Skip(1)):"");
            if (argsMap.Keys.Contains("app"))
                switch (argsMap["app"])
                {
                    case "sendEmail":
                        if (argsMap.ContainsKey("yaml"))
                        {

                            var yaml = new YamlStream();
                            using (var sr = new FileInfo(argsMap["yaml"]).OpenText())
                                yaml.Load(sr);
                            OA.sendEmail(((YamlMappingNode)yaml.Documents[0].RootNode).Children.Select(entry => new KeyValuePair<string, string>(entry.Key.ToString(), entry.Value.ToString())).ToDictionary(prop => prop.Key, prop => prop.Value));
                        }
                        else
                            OA.sendEmail(argsMap); break;
                    case "automateEmail":
                        if (argsMap.ContainsKey("path"))
                         OA.StartOAHelper(argsMap["path"]); else Console.WriteLine("Please provide path to configuration yaml files!"); break;

                }
            else Console.WriteLine("Please provide arguments!");
            
            //OA.StartOAHelper();
        }
    }
}
