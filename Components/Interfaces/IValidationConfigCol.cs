using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Components.Interfaces
{
     public abstract class IValidationConfigCol<T,V>:List<V>
    { 
       public T UnitValidator { get; set; }
    }
}
