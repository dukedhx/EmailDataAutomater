using ClassLibrary1.Utils;
using ClassLibrary1.Utils.Persistence;
using ConsoleApp1.Components.Interfaces;
using ConsoleApp1.Entities.Control;
using ConsoleApp1.Entities.Misc;
using ConsoleApp1.Utils;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using YamlDotNet.RepresentationModel;


namespace ConsoleApp1.Components.Controllers
{
    public enum WSVconfigsEnum { wwsn = 1, macrosnippet = 2, version = 3 }
    public class WorkSheetValidationController : IControl<FileInfo, IValidationConfigCol<IValidator<String, String>, EmailValidatorConfig>, IDictionary<String, ValidResults>>
    {
       

        public IDictionary<String, ValidResults> execute(FileInfo subject, IValidationConfigCol<IValidator<string, string>, EmailValidatorConfig> configCol, IDictionary<String, String> configs = null)
        {
            IDictionary<String, ValidResults> errorscol = new Dictionary<String, ValidResults>();
            using (var p = new ExcelPackage(subject))
            {
                var wb = p.Workbook;
                if ((!String.IsNullOrWhiteSpace(configs?[WSVconfigsEnum.macrosnippet.ToString()]) && (wb.VbaProject?.Modules["ThisWorkbook"] == null || wb.VbaProject?.Modules["ThisWorkbook"].Code?.Contains(configs?[WSVconfigsEnum.macrosnippet.ToString()]) != true)) || wb.Worksheets.SingleOrDefault(w => w.Name == configs?[WSVconfigsEnum.wwsn.ToString()]) == null || (configs?.ContainsKey(WSVconfigsEnum.version.ToString())==true&&!String.IsNullOrWhiteSpace(configs[WSVconfigsEnum.version.ToString()]) &&wb.Names["version"]?.Value?.ToString() != configs?[WSVconfigsEnum.version.ToString()])) return null;
                foreach (var vc in configCol.Where(vc => vc.ID != null && vc.criteria!=null))
                {
                    ValidResults errors = new ValidResults();
                    errors.sheetName = configs?[WSVconfigsEnum.wwsn.ToString()];
                    errorscol.Add(vc.ID, errors);

                    
                        var MatcherDict = errors.MatcherDict;


                        foreach (var entry in vc.criteria)
                        {




                            var keys = entry.Key.ToString().Split("@");
                            var keyname = keys.Last();
                            var matcher = entry.Value.ToString();
                            if (keys.Length > 1)
                            {
                                errors["hasRow"] = false;

                                MatcherCol amc;
                                var dkey = keys[0].ToLower();
                                if (!MatcherDict.TryGetValue(dkey, out amc)) MatcherDict.Add(dkey, amc = new MatcherCol());

                                if (keyname.ToLower().Equals("start") && wb.Names.ContainsKey(matcher)) amc.StartRow = wb.Names[matcher].End.Row + 1;
                                else if (keyname.ToLower().Equals("end") && wb.Names.ContainsKey(matcher)) amc.EndRow = wb.Names[matcher].Start.Row - 1;
                                else if (wb.Names.ContainsKey(keyname)) amc.Add(keyname, Tuple.Create(wb.Names[keyname].Start.Column, matcher));
                                else errors.Add(keyname, null);
                            }

                            else if ((keys = entry.Key.ToString().Split("!")).Length > 1)
                            {
                                switch (keys[0])
                                {
                                    case "control":
                                        var proc = Tools.RunProcess(@"cscript", new Dictionary<string, string>() { { Constants.argsEnum.redirectstdoutput.ToString(), "true" }, { Constants.argsEnum.pargs.ToString(), $"\"{Path.Combine(Tools.GetExecutingPath, @"scripts\getControlValue.vbs")}\"  \"{subject.FullName}\" \"{keys[1]}\"" } });
                                        var val = proc.StandardOutput.ReadToEnd().Split(Environment.NewLine).Where(l => !String.IsNullOrWhiteSpace(l)).Last();
                                        var rst = configCol.UnitValidator.validate(val, matcher);
                                        if (rst == true)
                                            errors.vals.Add(keys[1], val);
                                        else errors.Add(keys[1], rst);
                                        proc.WaitForExit();

                                        break;

                                }
                            }
                            else if (wb.Names.ContainsKey(keyname))
                            {
                                var val = wb.Names[keyname].Value?.ToString();
                                var result = configCol.UnitValidator.validate(val, matcher);
                                if (result != true)
                                    errors.Add(keyname, result);
                                else errors.vals.Add(keyname, val);

                            }

                            else errors.Add(keyname, null);
                        }
                    
                }
                var matchercols = errorscol.Select(ec => ec.Value.MatcherDict.Select(md => md.Value)).SelectMany(x => x);
                ExcelWorksheet worksheet = wb.Worksheets[configs?[WSVconfigsEnum.wwsn.ToString()]];
                foreach (var amc in matchercols.Where(md => md.isRangeValid).Select(md => md.range).Collapse(Comparer<int>.Create((x, y) => { return x - y; })).Select(r => Tuple.Create(r, errorscol.Where(ec => ec.Value.MatcherDict.Values.Any()).Select(ec => ec.Value))))
                {

                    for (int i = amc.Item1.Start; i < amc.Item1.End; i++)
                    {
                        if (worksheet.Cells[string.Format("{0}:{0}", i)].All(c => string.IsNullOrWhiteSpace(c.Value?.ToString())))
                            continue;
                        foreach (var aerrors in amc.Item2)
                        {
                            if (aerrors.ContainsKey("hasRow"))
                                aerrors.Remove("hasRow");

                            aerrors.AddAll(aerrors.Matchers.Select(item => (item.Key, configCol.UnitValidator.validate(worksheet.Cells[i, item.Value.Item1].Value?.ToString(), item.Value.Item2))).Where(item => item.Item2 != true).ToDictionary(prop => $"{prop.Item1}[Row:{i}]", prop => prop.Item2));
                        }
                    }
                    // if (!errors.Any(e => e.Value != true)) errors.AddAll(amc.Select(a=>new KeyValuePair<string, bool?> (a.Key,true)).Where(e=>wb.Names.ContainsKey(e.Key)));
                }


            }


            return errorscol;
        }
    }
}
