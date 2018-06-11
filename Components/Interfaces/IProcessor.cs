using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Components.Interfaces
{
  public  interface IProcessor<T,V>
    {
        bool? process(T subject, V controller, IDictionary<String, String> configs =null); 
    }
}
