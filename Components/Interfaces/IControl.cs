using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Components.Interfaces
{
    public interface IControl<T,C,R>
    {
        R execute(T subject, C configCol,  IDictionary<String,String> configs = null);
    }
}
