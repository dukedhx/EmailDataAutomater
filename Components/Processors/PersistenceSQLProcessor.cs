using ClassLibrary1.Utils;
using ClassLibrary1.Utils.Persistence;
using ConsoleApp1.Components.Consolidators;
using ConsoleApp1.Components.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConsoleApp1.Components.Contollers
{
    public enum PSTconfigsEnum { template=1,ppResultMap=2, pstPath=3,pstConnString=4}
    public class PersistenceSQLProcessor : IProcessor<IDictionary<String, String>, Object>, IDisposable
    {
        private readonly String path;
        private readonly StreamWriter handle;
        private Dictionary<string, string[]> fflds;

        protected bool finalize()
        {
            return String.IsNullOrWhiteSpace(path )?false: Tools.FileOverWriteMove(path, path + ".sql");

        }

        public PersistenceSQLProcessor(String path,String connstring, IDictionary<string, IEnumerable<KeyValuePair<string, string>>> bindmaps)
        {
            if(bindmaps!=null)
            this.fflds = bindmaps[XCDconfigsEnum.IAFormatFields.ToString()].ToDictionary(prop => prop.Key, prop => prop.Value?.Split(","));
            if (!String.IsNullOrWhiteSpace(path)) {
                handle = File.AppendText(path);
                if(!String.IsNullOrWhiteSpace(connstring))
                handle.WriteLine($"/* {connstring} */");
                this.path = path;
            }
        }

        public  bool? process(IDictionary<string, string> subject, Object controller, IDictionary<string, string> configs = null)
        {
            try
            {
                var queryTemplate = configs[PSTconfigsEnum.template.ToString()];
                if (String.IsNullOrWhiteSpace(queryTemplate) || subject == null || !subject.Keys.Any()) return false;
                handle.WriteLine(subject.Aggregate(queryTemplate, (c, v) => c.Replace($"@'{v.Key}'@", MySQLHelper.sanitizeQuery(fflds?.GetValueOrDefault("date")?.Any(f => f.Equals(v.Key, StringComparison.OrdinalIgnoreCase)) == true ? DateTime.Parse(v.Value).ToString("MM/dd/yyyy") : v.Value))) + ";");
                return true;

            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                return false;
            }
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    
                        handle?.Flush();
                        handle?.Close();
                    
                    //    finalize();
                    
                    // TODO: dispose managed state (managed objects).
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~PersistenceProcessor() {
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
