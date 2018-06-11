using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConsoleApp1.Entities.Misc
{
    public class Range<T>
    {

        public T Start;
        public T End;

        public Range()
        {
        }

        public Range(T start, T end)
        {
            Start = start;
            End = end;
        }

    }
}
