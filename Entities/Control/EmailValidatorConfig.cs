using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ConsoleApp1.Entities.Misc;
using YamlDotNet.RepresentationModel;
using ConsoleApp1.Components.Interfaces;

namespace ConsoleApp1.Entities.Control
{
    public class EmailValidatorConfig:IValidationConfig
    {
        public IDictionary<String,String> criteria;
        public String ID;
        public String rejFolder;
        public String[] rejEmails;
        public String[] sucEmails;
        public FileInfo sucTemplate;
        public FileInfo rejTemplate;
        private Boolean _continueOnReject;
        private Boolean _rejectOnInvalid;
        public Boolean useCustomValidation;
        public Dictionary<String, String> ResultMap;

        public EmailValidatorConfig()
        {
            this._continueOnReject = false;
            this._rejectOnInvalid = true;
           
        }

        public EmailValidatorConfig(FileInfo resultmap)
        {
            var yaml = new YamlStream();

            using (var sr = resultmap.OpenText())

                yaml.Load(sr);

            ResultMap = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.Select(entry => new { Key = entry.Key.ToString(), Value = entry.Value.ToString() }).ToDictionary(prop => prop.Key, prop => prop.Value);
        }

       public bool continueOnReject { get => _continueOnReject; set =>   _continueOnReject = value;  }
        public bool rejectOnInvalid { get => _rejectOnInvalid; set => _rejectOnInvalid = value; }
    }
}
