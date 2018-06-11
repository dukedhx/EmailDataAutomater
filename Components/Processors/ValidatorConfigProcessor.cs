using ConsoleApp1.Components.Interfaces;
using ConsoleApp1.Entities.Control;
using ConsoleApp1.Entities.Misc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ConsoleApp1.Components.Controllers
{
  public  abstract class ValidatorConfigProcessor<T> : IProcessor<T, IControl<FileInfo, IValidationConfigCol<IValidator<String, String>, EmailValidatorConfig>, Dictionary<string,ValidResults>>>
    {
        public bool success=true;
        public IEnumerable<EmailValidatorConfigCol> vconfigCol;
        protected abstract String GetRelativeId(String id);

        public abstract bool? process(T subject, IControl<FileInfo, IValidationConfigCol<IValidator<string, string>, EmailValidatorConfig>, Dictionary<string, ValidResults>> controller, IDictionary<string, string> configs = null);

        public Dictionary<String, ValidResults> ExecuteValidationConfig(IEnumerable<EmailValidatorConfigCol> validColsCol, Dictionary<EmailValidatorConfig, IEnumerable<Tuple<String, String, Boolean?>>> avalCols,String path, Dictionary<String, ValidResults> resultsCol)
        {

            ValidResults vr = new ValidResults();

            foreach (var validCols in validColsCol)
            {

              var aresultscol=  new WorkSheetValidationController().execute(new FileInfo(path), validCols, new Dictionary<String, String>() { { WSVconfigsEnum.macrosnippet.ToString(), validCols.macroSnippet }, { WSVconfigsEnum.version.ToString(), validCols.version }, { WSVconfigsEnum.wwsn.ToString(), validCols.sheetName } });
                
                if (aresultscol == null)
                    continue;

                var stopOnReject = false;

                //Logger.WriteToConsole($"Processing results [{attachment.FileName}]");
                foreach (var aresults in aresultscol)
                {
                    var t = validCols.FirstOrDefault(vc => vc.ID.Equals(aresults.Key));
                    var results = aresults.Value;
                    if (results.Any(r => r.Value != true))
                    {

                        if (!t.continueOnReject)




                            avalCols.Clear();


                        if (t.rejectOnInvalid) success = false;
                        var rsts = results.Select(r => Tuple.Create(GetRelativeId(path), r.Key, r.Value));
                        if (avalCols.ContainsKey(t))
                            avalCols[t].Concat(rsts);
                        else avalCols.Add(t, new LinkedList<Tuple<String, String, Boolean?>>(rsts));



                       // 



                        if (!t.continueOnReject)
                        {
                            stopOnReject = true;
                            break;
                        }
                    }
                    if (success)



                        vr.AddAll(results);


                    if (t.sucTemplate != null && !avalCols.ContainsKey(t))
                        avalCols.Add(t, null);

                }

                if (stopOnReject) break;

                
            }

            if (success)
            {

                if (!String.IsNullOrWhiteSpace(vr.sheetName))
                    resultsCol.Add(path, vr);
                //  rhtml.Add((vr.ID = attachment.FileName), "Success");

            }

            return resultsCol;

        }
    }
}
