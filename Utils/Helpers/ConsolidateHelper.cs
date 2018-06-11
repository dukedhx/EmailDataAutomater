using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ConsoleApp1.Entities.Misc;
using YamlDotNet.RepresentationModel;
using ClassLibrary1.Utils.Persistence;
using ClassLibrary1.Utils;
using ConsoleApp1.Entities.Control;

namespace ConsoleApp1.Utils
{
    public class ConsolidateHelper
    {

        public static string NewID
        {
            get
            {
                return Guid.NewGuid().ToString("N");
            }
        }

        public static string TimeLapsedID
        {
            get
            {
                return DateTime.Now.Ticks.ToString().Substring(4, 6);
            }
        }

        public static string NewInt
        {
            get
            {
                return new Random().Next(0, 999999).ToString("D6");
            }
        }

        public static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";
            
            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }

        //public static IDictionary<FileInfo, ValidResults> validateWS(String wwsn,String version,String macrosnippet,FileInfo wsfile,IEnumerable<FileInfo> vfiles,CellValidator cv)
        //{
        //    IDictionary<FileInfo, ValidResults> errorscol = new Dictionary<FileInfo, ValidResults>();
        //    using (var p = new ExcelPackage(wsfile))
        //    {
        //        var wb = p.Workbook;
        //        if ((!String.IsNullOrWhiteSpace(macrosnippet)&&(wb.VbaProject?.Modules["ThisWorkbook"]==null|| wb.VbaProject?.Modules["ThisWorkbook"].Code?.Contains(macrosnippet) != true)) || wb.Worksheets.SingleOrDefault(w => w.Name == wwsn) == null||wb.Names["version"]?.Value?.ToString()!= version)  return null; 
        //        var yaml = new YamlStream();
        //        foreach (var vfile in vfiles)
        //        {
        //            ValidResults errors = new ValidResults();
        //            errors.sheetName = wwsn;
        //            errorscol.Add(vfile, errors);
        //            using (var sr = vfile.OpenText())

        //                yaml.Load(sr);

        //            var MatcherDict = errors.MatcherDict;


        //            foreach (var entry in ((YamlMappingNode)yaml.Documents[0].RootNode).Children)
        //            {




        //                var keys = entry.Key.ToString().Split("@");
        //                var keyname = keys.Last();
        //                var matcher = entry.Value.ToString();
        //                if (keys.Length > 1)
        //                {
        //                    errors["hasRow"] = false;

        //                    MatcherCol amc;
        //                    var dkey = keys[0].ToLower();
        //                    if (!MatcherDict.TryGetValue(dkey, out amc)) MatcherDict.Add(dkey, amc = new MatcherCol());

        //                    if (keyname.ToLower().Equals("start") && wb.Names.ContainsKey(matcher)) amc.StartRow = wb.Names[matcher].End.Row + 1;
        //                    else if (keyname.ToLower().Equals("end") && wb.Names.ContainsKey(matcher)) amc.EndRow = wb.Names[matcher].Start.Row - 1;
        //                    else if (wb.Names.ContainsKey(keyname)) amc.Add(keyname, Tuple.Create(wb.Names[keyname].Start.Column, matcher));
        //                    else errors.Add(keyname, null);
        //                }

        //                else if ((keys = entry.Key.ToString().Split("!")).Length > 1)
        //                {
        //                    switch (keys[0]) {
        //                        case "control":
        //                            var proc = Tools.RunProcess(@"cscript", new Dictionary<string, string>() { { Constants.argsEnum.redirectstdoutput.ToString(), "true" }, { Constants.argsEnum.pargs.ToString(), $"\"{Path.Combine(Tools.GetExecutingPath, @"scripts\getControlValue.vbs")}\"  \"{wsfile.FullName}\" \"{keys[1]}\"" } });
        //                          var val=  proc.StandardOutput.ReadToEnd().Split(Environment.NewLine).Where(l=>!String.IsNullOrWhiteSpace(l)).Last();
        //                            var rst = cv.validate(val, matcher);
        //                            if (rst == true)
        //                                errors.vals.Add(keys[1], val);
        //                            else errors.Add(keys[1], rst);
        //                            proc.WaitForExit();                                   

        //                            break;

        //                }
        //                }
        //                else if (wb.Names.ContainsKey(keyname))
        //                {
        //                 var val=   wb.Names[keyname].Value?.ToString();
        //                    var result = cv.validate(val, matcher);
        //                    if (result != true)
        //                        errors.Add(keyname, result);
        //                    else errors.vals.Add(keyname, val);

        //                }

        //                else errors.Add(keyname, null);
        //            }
        //        }
        //        var matchercols = errorscol.Select(ec => ec.Value.MatcherDict.Select(md => md.Value)).SelectMany(x=>x);
        //        ExcelWorksheet worksheet = wb.Worksheets[wwsn];
        //        foreach (var amc in matchercols.Where(md=>md.isRangeValid).Select(md => md.range).Collapse(Comparer<int>.Create((x,y)=> { return x - y; })).Select(r=>Tuple.Create ( r, errorscol.Where(ec => ec.Value.MatcherDict.Values.Any()).Select(ec=>ec.Value))))
        //            {

        //                for (int i = amc.Item1.Start; i < amc.Item1.End; i++)
        //                {
        //                    if (worksheet.Cells[string.Format("{0}:{0}", i)].All(c => string.IsNullOrWhiteSpace(c.Text)))
        //                        continue;
        //                foreach (var aerrors in amc.Item2){
        //                    if(aerrors.ContainsKey("hasRow"))
        //                    aerrors.Remove("hasRow");
                            
        //                    aerrors.AddAll(aerrors.Matchers.Select(item => (item.Key, cv.validate(worksheet.Cells[i, item.Value.Item1].Value?.ToString(), item.Value.Item2))).Where(item => item.Item2 != true).ToDictionary(prop => $"{prop.Item1}[Row:{i}]", prop => prop.Item2));
        //                }
        //                }
        //                // if (!errors.Any(e => e.Value != true)) errors.AddAll(amc.Select(a=>new KeyValuePair<string, bool?> (a.Key,true)).Where(e=>wb.Names.ContainsKey(e.Key)));
        //            }
                    
                
        //    }


        //    return errorscol;
        //}

        
        public static Boolean isIgnoredHeader(String header)
        {
            return String.IsNullOrWhiteSpace(header) || String.Compare(header, "null", StringComparison.OrdinalIgnoreCase) == 0;
        }

        public static ExcelWorksheet GetWorksheetOrAdd(ExcelWorkbook wb, String wsn)
        {
            return wb.Worksheets.Any(ws => ws.Name == wsn) ? wb.Worksheets[wsn] : wb.Worksheets.Add(wsn);
        }

        public static int? GetColIDbyName(ExcelWorksheet ws, String ColName, int ColRow)
        {

            return ws.Cells[string.Format("{0}:{0}", ColRow)].FirstOrDefault(c => string.Equals(c.Value?.ToString(), ColName, StringComparison.OrdinalIgnoreCase))?.Columns;
        }

        public static IEnumerable<(String, String)> GetRowByID(ExcelWorksheet ws, String target, String colname, int ColRow=1)
        {
            var colindex = GetColIDbyName(ws,colname, ColRow);
            return colindex == null ? null : GetRowByID(ws,target, (int)colindex, ColRow);
        }

            public static IEnumerable<(String,String)> GetRowByID(ExcelWorksheet ws, String target, int ColIndex, int ColRow=1)
        {
            
            var tcellrow = ws.Cells[ColRow + 1, ColIndex, ws.Dimension.Rows, ColIndex].FirstOrDefault(c => string.Equals(c.Value?.ToString(), target, StringComparison.OrdinalIgnoreCase))?.Start.Row;
            return tcellrow == null ? null : ws.Cells[string.Format("{0}:{0}", tcellrow)].Select(c=> ( ws.Cells[ColRow,c.Start.Column].Value?.ToString(),c.Value?.ToString())).ToArray();
        }


        //public static Boolean? writeDataResult(String dpath, String apath, String wsn, IEnumerable<KeyValuePair<String,String>>evals, IEnumerable<SealedNameList> headers, Dictionary<String, MatcherCol> matcherdict, Boolean lockobj)
        //{
        //    long ticks = TimeSpan.FromMinutes(5).Ticks;
        //    while (ticks-- > 0)
        //    {
        //        if (!lockobj|| Monitor.TryEnter(lockobj))
        //        {
        //            try {
        //                using (var p = new ExcelPackage(new FileInfo(apath))) 
        //                    return writeDataResult(dpath,wsn, GetWorksheetOrAdd(p.Workbook,wsn)  , evals,headers, matcherdict);
        //             }
        //            catch (Exception ex)
        //        {
        //            Logger.Log(ex);
        //            return false;
        //        }
        //        finally { if (lockobj) Monitor.Exit(lockobj); }
        //    }
        //        else Thread.Sleep(1);
        //}
        //    return null;
        //}

        //public static Boolean writeDataResult(String dpath,String wwsn, ExcelWorksheet worksheet, IEnumerable<KeyValuePair<String,String>>evals, IEnumerable<SealedNameList> headers, Dictionary<String, MatcherCol> matcherdict)
        //{
        //    Logger.Log($"Writing data result for {dpath}");

        //    using (var dp = new ExcelPackage(new FileInfo(dpath)))

        //        try
        //        {
        //            var ws = dp.Workbook.Worksheets[wwsn];
        //            var wb = dp.Workbook;
        //            foreach (var header in headers)
        //                worksheet.Cells[1, header.value].Value = header.name;
        //            var yaml = new YamlStream();
        //            using (var sr = new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["dataResultMappingPath"])).OpenText())

        //                yaml.Load(sr);

        //            var ResultMap = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.Select(entry => new { Key = entry.Key.ToString(), Value = entry.Value.ToString() }).ToDictionary(prop => prop.Key, prop => prop.Value);


        //            using (var sr = new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["IAFormatFields"])).OpenText())

        //                yaml.Load(sr);

        //            var fflds = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.Select(entry => new { Key = entry.Key.ToString(), Value = entry.Value.ToString() }).ToDictionary(prop => prop.Key, prop => prop.Value?.Split(","));

        //            using (var sr = new FileInfo(Path.Combine(ConfigurationManager.Configuration["excPath"] ?? "", ConfigurationManager.Configuration["dataResultValMapping"])).OpenText())

        //                yaml.Load(sr);

        //            var drvm = ((YamlMappingNode)yaml.Documents[0].RootNode).Children.Select(entry => new { Key = entry.Key.ToString(), Value = entry.Value.ToString() }).ToDictionary(prop => prop.Key, prop => prop.Value?.Split(",").Select(e => e.Split(":")).Where(p => p.Length>1).Select(p=>new { Key=p[0],Value=p[1]}).ToDictionary(p => p.Key, p => p.Value));




        //            int r = worksheet.Dimension.End.Row + 1;
        //            var pdr = ResultMap.Where(e => wb.Names.ContainsKey(e.Key)).Select(e => new { Key = e.Value, Value = wb.Names[e.Key].Value.ToString(), Column = wb.Names[e.Key].Start.Column });
        //            var adr = evals.Where(e => !ResultMap.ContainsKey(e.Key)||drvm.ContainsKey(ResultMap.GetValueOrDefault(e.Key))).Select(e=>new { Key=ResultMap.GetValueOrDefault(e.Key)??e.Key,e.Value});
                    
        //            foreach (var matcher in matcherdict.Where(md=>md.Value.isRangeValid).Select(md=>md.Value.range).Collapse(Comparer<int>.Create((x, y) => { return x - y; })).Select(rr => Tuple.Create(rr, matcherdict.Values.Where(v=>v.Values.Any()).SelectMany(e=>e.Keys).Where(e => ResultMap.ContainsKey(e)).Select(e => ResultMap[e])))) {
                       
        //              //  var rdr = ;
        //                int StartRow = matcher.Item1.Start-1, EndRow = matcher.Item1.End;
        //                while (StartRow++ < EndRow)
        //                {
        //                    if (ws.Cells[string.Format("{0}:{0}", StartRow)].All(c => string.IsNullOrWhiteSpace(c.Text)))
        //                        continue;
                            
        //                    foreach (var k in pdr)
        //                    {

        //                        var cell = worksheet.Cells[r, headers.First(h => string.Equals(h.name, k.Key, StringComparison.OrdinalIgnoreCase)).value];
        //                        if(fflds["number"]?.Any(f=>f.Equals(k.Key, StringComparison.OrdinalIgnoreCase))==true)
        //                        cell.Style.Numberformat.Format = "0";                             
                                
        //                            cell.Value = fflds["date"]?.Any(f => f.Equals(k.Key, StringComparison.OrdinalIgnoreCase)) ==true? DateTime.Parse(k.Value).ToString("MM/dd/yyyy") : k.Value;

        //                    }
        //                    foreach (var k in adr)
                            

                                
        //                        worksheet.Cells[r, headers.First(h => string.Equals(h.name, k.Key, StringComparison.OrdinalIgnoreCase)).value].Value = drvm.GetValueOrDefault(k.Key)?[k.Value]??k.Value;
                            
                            
        //                    foreach (var k in matcher.Item2)
        //                        worksheet.Cells[r, headers.First(h => string.Equals(h.name, k, StringComparison.OrdinalIgnoreCase)).value].Value = ws.Cells[StartRow, pdr.First(p => p.Key == k).Column].Value;
        //                    r++;
        //                }

        //            }
        //            Logger.Log($"{dpath}:success");
        //            return true;
        //        }
        //        catch (Exception ex)
        //        {
        //            Logger.Log(ex);
        //        }
        //        return false;


        //}

      
    }
}
