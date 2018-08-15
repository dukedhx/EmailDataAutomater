using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using YamlDotNet.RepresentationModel;

namespace ConsoleApp1.Utils.Helpers
{
    public static class ConfigTools
    {

        
        public static Dictionary<String,String> YamlToDict(String path)
        {
            var yaml = new YamlStream();
            using (var sr = new FileInfo(path).OpenText())

                yaml.Load(sr);

            return ((YamlMappingNode)yaml.Documents[0].RootNode).Children.ToDictionary(entry => entry.Key.ToString(), entry => entry.Value?.ToString());
        }
    }
}
