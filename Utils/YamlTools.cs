using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using YamlDotNet.RepresentationModel;

namespace ConsoleApp1.Util
{
    public static class YamlTools
    {

        public static Dictionary<String, YamlNode> getNodeMapFromYaml(String path)
        {
           
            return getTypedMapFromYaml(path,n=>n);
        }

        public static IEnumerable<T> getFuncMapFromYaml<T>(String path, Func<KeyValuePair<YamlNode,YamlNode>, T> func)
        {
            if (!File.Exists(path)) return new List<T>(0);
            YamlStream yaml = new YamlStream();
            using (var sr = new FileInfo(path).OpenText())

                yaml.Load(sr);
            return ((YamlMappingNode)yaml.Documents[0].RootNode).Children.Select(e=>func(e));
        }

            public static Dictionary<String, T> getTypedMapFromYaml<T>(String path,Func<YamlNode,T> func)
        {
            if (!File.Exists(path)) return new Dictionary<string, T>();
            
            return getFuncMapFromYaml(path,entry => new KeyValuePair<string, T>( entry.Key.ToString(),  func(entry.Value) )).ToDictionary(prop => prop.Key, prop => prop.Value);
        }

        public static Dictionary<String, String> getResultMapFromYaml(String path)
        {

            return getTypedMapFromYaml<String>(path,n=>n.ToString());
        }

        public static IEnumerable<KeyValuePair<String, String>> getResultKVFromYaml(String path)
        {
            
            return getResultMapFromYaml( path);
        }
    }
}
