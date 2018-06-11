using ConsoleApp1.Components.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConsoleApp1.Entities.Control
{
   public  class EmailValidatorConfigCol: IValidationConfigCol<IValidator<String,String>, EmailValidatorConfig>
    {
        public String macroSnippet;
        public String version;
        public String sheetName;

        public EmailValidatorConfigCol(string macroSnippet, string version, string sheetName, IValidator<String, String> cellValidator =null)
        {
            this.macroSnippet = macroSnippet;
            this.version = version;
            this.sheetName = sheetName;
            this.UnitValidator = cellValidator == null? new CellValidator(): cellValidator;
        }
    }
}
