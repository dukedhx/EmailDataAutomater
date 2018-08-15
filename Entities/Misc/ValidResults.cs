using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConsoleApp1.Entities.Misc
{
    public class ValidResults:Dictionary<String,Boolean?>
    {

        public Dictionary<String, MatcherCol> MatcherDict=new Dictionary<String, MatcherCol> ();
        public String sheetName;
        public ValidResultConfig config;

        private IEnumerable<KeyValuePair<String, Tuple<int, String>>> _Matchers;
        public IEnumerable<KeyValuePair<String, Tuple<int, String>>> Matchers { get { return _Matchers == null ? _Matchers = MatcherDict.Values.SelectMany(x => x) : _Matchers; } }
        public Dictionary<String, String> vals = new Dictionary<String, String>();
        public Boolean AddAll (IEnumerable<KeyValuePair<String,Boolean?>> vals)
        {
            
            if (vals is ValidResults)
            {
                
                var avals=vals as ValidResults;
                sheetName = avals.sheetName;
              //  config = avals.config;
                MatcherDict = MatcherDict.Concat(avals.MatcherDict).GroupBy(prop => prop.Key).ToDictionary(prop => prop.Key, prop => prop.First().Value);
                this.vals=this.vals.Concat( avals.vals).GroupBy(prop => prop.Key).ToDictionary(prop => prop.Key, prop => prop.First().Value);
            }
            
            if (vals == null) return false;
            foreach (var val in vals) this[val.Key] = val.Value;
            return true;
        }
    }
}
