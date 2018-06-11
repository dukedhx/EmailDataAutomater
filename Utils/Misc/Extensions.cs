using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ConsoleApp1.Entities.Misc;

namespace ConsoleApp1.Utils
{
    internal static class Extensions
    {
        public static DateTime AddBusinessDays(this DateTime source,
                                         int businessDays)
        {
            var dayOfWeek = businessDays < 0
                                  ? ((int)source.DayOfWeek - 12) % 7
                                  : ((int)source.DayOfWeek + 6) % 7;

            switch (dayOfWeek)
            {
                case 6:
                    businessDays--;
                    break;
                case -6:
                    businessDays++;
                    break;
            }

            return source.AddDays(businessDays + ((businessDays + dayOfWeek) / 5) * 2);
        }


        public static IEnumerable<Range<T>> Collapse<T>(this IEnumerable<Range<T>> me, IComparer<T> comparer)
        {
            List<Range<T>> orderdList = me.OrderBy(r => r.Start).ToList();
            List<Range<T>> newList = new List<Range<T>>();
            if (orderdList.Any())
            {
                T max = orderdList[0].End;
                T min = orderdList[0].Start;

                foreach (var item in orderdList.Skip(1))
                {
                    if (comparer.Compare(item.End, max) > 0 && comparer.Compare(item.Start, max) > 0)
                    {
                        newList.Add(new Range<T> { Start = min, End = max });
                        min = item.Start;
                    }
                    max = comparer.Compare(max, item.End) > 0 ? max : item.End;
                }
                newList.Add(new Range<T> { Start = min, End = max });
            }
            return newList;
        }
    }
}
