using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Components.Interfaces
{
    public interface IValidationConfig
    {
         Boolean continueOnReject { get; set; }
         Boolean rejectOnInvalid { get; set; }
    }
}
