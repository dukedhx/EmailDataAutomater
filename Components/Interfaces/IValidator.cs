using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Components.Interfaces
{
  public  interface IValidator<T, V> 
    {
        Boolean? validate(T value, V matcher);
    }
}
