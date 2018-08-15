using ClassLibrary1.Utils;
using ClassLibrary1.Utils.Persistence;
using ConsoleApp1.Components.Interfaces;
using ConsoleApp1.Components.Contollers;
using ConsoleApp1.Entities.Misc;
using ConsoleApp1.Utils;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
//using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ConsoleApp1.Components.Consolidators
{
    public enum XCDconfigsEnum { DRoptions=0, dpath = 1, macrosnippet = 2, version = 3, resultMap=4, IAFormatFields=5, dataResultValMapping=6,wwsn=7,cwwsn=8,headersID=9,xlsxtemplate=10 }

    public class XlsxConsolidator : IConsolidator<IEnumerable< IProcessor<IDictionary<String, String>, Object>>>
    {
        public IEnumerable<KeyValuePair<string, string>> evals;
        private IEnumerable<KeyValuePair<string, string>> cevals;
        public IDictionary<string, IEnumerable<KeyValuePair<string, string>>> bindmaps;
        public IEnumerable<SealedNameList> headers;

        protected IEnumerable<KeyValuePair<string,string>> generateAdr(Dictionary<String,String> ResultMap, Dictionary<String, Dictionary<String, String>> drvm)
        {
            return cevals==null?null:cevals.Where(e => !ResultMap.ContainsKey(e.Key) || drvm.ContainsKey(ResultMap.GetValueOrDefault(e.Key))).Select(e => new KeyValuePair<string,string> ( ResultMap.GetValueOrDefault(e.Key) ?? e.Key, e.Value ));
        }


       

        public bool consolidate(IEnumerable<IProcessor<IDictionary<String, String>, Object>> subject, Dictionary<string, ValidResults> resultsCol, IDictionary<String, String> configs)
        {
          
                
             
                    bool success = true;

            if(subject?.Any()==true)
            foreach (var results in resultsCol)
            {
                configs[XCDconfigsEnum.wwsn.ToString()] = results.Value.sheetName;
                configs[XCDconfigsEnum.dpath.ToString()] = results.Key;
                cevals = evals.Concat(results.Value.Select(e => new KeyValuePair<String, String>(e.Key, "")).Append(new KeyValuePair<string, string>(ConfigHelper.Filename, Path.GetFileName(results.Key)))).Concat(results.Value.vals);
                success = writeDataResult(subject, results.Value.MatcherDict,  configs, bindmaps);
            }
                

                return success;              

          
        }

        public bool postProduction(IEnumerable<IProcessor<IDictionary<String, String>, Object>> writeObject,  IDictionary<string, string> configs)
        {
            return true;
        }

        protected Dictionary<String, Dictionary<string, string>> getValuesMapDict(IEnumerable<KeyValuePair<string, string>> ppmap)
        {
           return ppmap?.ToDictionary(prop => prop.Key, prop => prop.Value?.Split(",")?.Select(e => e.Split(":"))?.Where(p => p.Length > 1)?.Select(p => new { Key = p[0], Value = p[1] }).ToDictionary(p => p.Key, p => p.Value));
        }

        protected Object PreProcessValue(Object value,String key,IDictionary<String,String>values, Dictionary<String,Dictionary<string, string>>resultMap)
        {
            if (values != null)

                values[key] =resultMap?.GetValueOrDefault(key)?.GetValueOrDefault(value.ToString())??value?.ToString();
            
            return value;
        }
        

        public bool writeDataResult(IEnumerable<IProcessor<IDictionary<String, String>, Object>>  writeObject,Dictionary<string, MatcherCol> matcherdict, IDictionary<string, string> configs, IDictionary<string, IEnumerable<KeyValuePair<string, string>>> bindmaps)
        {

            using (var dp = new ExcelPackage(new FileInfo(configs[XCDconfigsEnum.dpath.ToString()])))
            {
                var ws = dp.Workbook.Worksheets[configs[XCDconfigsEnum.wwsn.ToString()]];
                var wb = dp.Workbook;
                

                var ResultMap = bindmaps[XCDconfigsEnum.resultMap.ToString()].ToDictionary(prop => prop.Key, prop => prop.Value);




             //   var fflds = bindmaps[XCDconfigsEnum.IAFormatFields.ToString()].ToDictionary(prop => prop.Key, prop => prop.Value?.Split(","));



                var drvm = getValuesMapDict(bindmaps[XCDconfigsEnum.dataResultValMapping.ToString()]);


                var ppresultMap = getValuesMapDict(bindmaps.ContainsKey(PSTconfigsEnum.ppResultMap.ToString()) ? bindmaps[PSTconfigsEnum.ppResultMap.ToString()]:null);

               
                var pdr = ResultMap.Where(e => wb.Names.ContainsKey(e.Key)).Select(e => new { Key = e.Value, Value = wb.Names[e.Key].Value.ToString(),  wb.Names[e.Key].Start.Column });
                //  var adr = generateAdr(ResultMap,drvm);


                foreach (var matcher in matcherdict.Where(md => md.Value.isRangeValid).Select(md => md.Value.range).Collapse(Comparer<int>.Create((x, y) => { return x - y; })).Select(rr => Tuple.Create(rr, matcherdict.Values.Where(v => v.Values.Any()).SelectMany(e => e.Keys).Where(e => ResultMap.ContainsKey(e)).Select(e => ResultMap[e]))))
                {
                    
                    //  var rdr = ;
                    int StartRow = matcher.Item1.Start - 1, EndRow = matcher.Item1.End;
                    while (StartRow++ < EndRow)
                    {
                        Dictionary<String, String> values = new Dictionary<string, string>();

                        if (ws.Cells[string.Format("{0}:{0}", StartRow)].All(c => string.IsNullOrWhiteSpace(c.Value?.ToString())))
                            continue;
                        
                        foreach (var k in pdr)
                        

                            PreProcessValue( k.Value,k.Key,values, ppresultMap);
                           

                        
                        foreach (var k in generateAdr(ResultMap, drvm))

                          

                                PreProcessValue( drvm.GetValueOrDefault(k.Key)?.GetValueOrDefault(k.Value) ?? k.Value, k.Key, values, ppresultMap);
                         

                        foreach (var k in matcher.Item2)
                          PreProcessValue(ws.Cells[StartRow, pdr.First(p => p.Key == k).Column].Value, k, values, ppresultMap);

                        foreach(var k in cevals.Where(c=>!values.ContainsKey(c.Key)))
                            PreProcessValue(k.Value, k.Key, values, ppresultMap);



                        Task.WaitAll(writeObject.Select(p=> Task.Factory.StartNew(() => p?.process(values, null, configs))).ToArray());
                       
                    }

                }
                return true;
            }
        }

        
       
    }
}
