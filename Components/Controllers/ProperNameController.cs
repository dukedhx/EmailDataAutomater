using ConsoleApp1.Components.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Components.Contollers
{
    public class ProperNameController : IControl<String, Object, String>
    {
        public string execute(string subject, object configCol, IDictionary<string, string> configs = null)
        {
            var today = DateTime.Now.ToString("s").Split("T")[0];
            var dparts = today.Split("-");
            return (subject??"").Replace("{today}", today).Replace("{day}", dparts[2]).Replace("{month}", dparts[1]).Replace("{year}", dparts[0]).Replace("{guid}",Guid.NewGuid().ToString());
        }
    }
}
