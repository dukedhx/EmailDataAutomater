using ClassLibrary1.Utils;
using ClassLibrary1.Utils.Persistence;
using ConsoleApp1.Components.Consolidators;
using ConsoleApp1.Components.Interfaces;
using ConsoleApp1.Entities.Misc;
using ConsoleApp1.Utils;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConsoleApp1.Components.Contollers
{
    public class XlsxResultProcessor : IProcessor<IDictionary<String, String>, Object>, IDisposable
    {
        private Dictionary<string, string[]> fflds;
        private ExcelPackage p;
        private ExcelWorksheet worksheet;
        private int r,c=0;
        public IEnumerable<SealedNameList> headers;
        //public IEnumerable<SealedNameList> headers { set { if (headers == null)_headers = value  } get => _headers; }
        public XlsxResultProcessor(String path,String cwwsn, IDictionary<string, IEnumerable<KeyValuePair<string, string>>> bindmaps, IEnumerable<SealedNameList>  headers)
        {
            this.fflds = bindmaps[XCDconfigsEnum.IAFormatFields.ToString()].ToDictionary(prop => prop.Key, prop => prop.Value?.Split(","));

            p = new ExcelPackage(new System.IO.FileInfo(path));
            worksheet = ConsolidateHelper.GetWorksheetOrAdd(p.Workbook, cwwsn);

            
            this.headers = headers;
            if(!p.Workbook.Names.ContainsKey("_end_")&&(worksheet.Dimension?.End?.Row ??0) <1)
            foreach (var header in headers)
                worksheet.Cells[1, header.value].Value = header.name; 
            if (p.Workbook.Names.ContainsKey("_end_")) {
                r = p.Workbook.Names["_end_"].Start.Row -1;
                c = p.Workbook.Names["_end_"].Start.Column - 1;}
            else
            r = (worksheet.Dimension?.End?.Row ?? 0) + 1;


        }

        public  bool? process(IDictionary<string, string> subject, object controller, IDictionary<string, string> configs = null)
        {
            

            foreach (var k in subject)
            {
                var header = headers.FirstOrDefault(h => string.Equals(h.name, k.Key, StringComparison.OrdinalIgnoreCase));
                if (header != null) {
                    
                    var cell = worksheet.Cells[r, c+header.value];
                    cell.Value = fflds.GetValueOrDefault("date")?.Any(f => f.Equals(k.Key, StringComparison.OrdinalIgnoreCase)) == true ? DateTime.Parse(k.Value).ToString("MM/dd/yyyy") : k.Value;

                    if (fflds.GetValueOrDefault("number")?.Any(f => f.Equals(k.Key, StringComparison.OrdinalIgnoreCase)) == true)
                        cell.Style.Numberformat.Format = "0";

                
                }
            }
            //if(subject.Any())
            r++;
            worksheet.InsertRow(r, 1);

            return true;
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls


        protected void postProduction()
        {
            //var ccoli = headers.FirstOrDefault(h => string.Equals(h.name, "count po", StringComparison.OrdinalIgnoreCase))?.value ?? -1;
            //var pocol = ConsolidateHelper.GetColumnName(headers.First(h => string.Equals(h.name, "IA PO #", StringComparison.OrdinalIgnoreCase)).value - 1);
            //if (ccoli > -1 && !String.IsNullOrWhiteSpace(pocol))
            //{
            //    int r = worksheet.Dimension.End.Row + 1;

            //    worksheet.Cells[2, ccoli].CreateArrayFormula($"SUM(1/COUNTIF({pocol}2:{pocol}{r - 1},{pocol}2:{pocol}{r - 1}))");
            //    worksheet.Cells[2, ccoli].Calculate();
            //    int iac = 0;
            //    var tcell = worksheet.Cells[2, headers.First(h => string.Equals(h.name, "Total IA Count", StringComparison.OrdinalIgnoreCase)).value];
            //    Int32.TryParse(tcell.Value?.ToString(), out iac);
            //    tcell.Value = iac + 1;

            //}
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // var worksheet = ConsolidateHelper.GetWorksheetOrAdd(writeObject.Workbook, configs[XCDconfigsEnum.cwwsn.ToString()]);
                    postProduction();
                    
                    p.Save();
                    var pfname = p.File.FullName;
                    p.Dispose();
                    Logger.WriteToConsole($"Successfully validated saved DataResults to [{pfname}]");
                    // TODO: dispose managed state (managed objects).
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~XlsxResultProcessor() {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }
}
