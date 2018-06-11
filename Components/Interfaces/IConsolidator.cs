using ClassLibrary1.Utils.Persistence;
using ConsoleApp1.Entities.Misc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1.Components.Interfaces
{
   public interface IConsolidator<T>
    {
         Boolean writeDataResult( T writeObject,  Dictionary<String, MatcherCol> matcherdict, IDictionary<String, String> configs, IDictionary<String, IEnumerable< KeyValuePair<String, String>>> bindmaps);

        Boolean consolidate(T writeObject, Dictionary<String, ValidResults> resultsCol, IDictionary<String, String> configs);

        Boolean postProduction(T writeObject,   IDictionary<String, String> configs);
    }
}
