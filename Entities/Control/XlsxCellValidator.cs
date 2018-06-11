using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ClassLibrary1.Utils;
using ConsoleApp1.Entities.Misc;
using ConsoleApp1.Utils;
using OfficeOpenXml;

namespace ConsoleApp1.Entities.Control
{
   public class XlsxCellValidator : CellValidator
    {
        
        public override bool? validate(string value, string matcher = null)
        {
            var parts = matcher.Split(",");
           var path= parts[1];
            if (File.Exists(path)) {
                using (FileStream fs = new FileStream(path, FileMode.Open))
                {
                    if(fs.CanRead)
                    using (ExcelPackage p = new ExcelPackage(fs))
                        return ConsolidateHelper.GetRowByID(parts.Length < 3 ? p.Workbook.Worksheets[0] : p.Workbook.Worksheets[parts[2]], value, parts[0], parts.Length < 4 ? 1 : Int32.Parse(parts[4]))!=null;

                }
            }
            Logger.Log($"[Xlsx source does not exist or not readable]{path}");
            return true;
        }
    }
}
