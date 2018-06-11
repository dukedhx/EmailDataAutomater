using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConsoleApp1.Entities.Misc
{
    public class MatcherCol: Dictionary<String,Tuple<int,String>>
    {
        public Boolean isRangeValid { get {return StartRow > 0 && EndRow > StartRow; } }
        public Range<int> range;
        public String ID;
        public int StartRow { get { return range.Start; } set { if (range == null) range = new Range<int>(); range.Start = value; } }
        public int EndRow { get { return range.End; } set { if (range == null) range = new Range<int>(); range.End = value; } }
    }
}
