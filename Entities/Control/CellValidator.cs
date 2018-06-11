using ConsoleApp1.Components.Interfaces;
using ConsoleApp1.Utils;
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ConsoleApp1.Entities.Control
{
    public class CellValidator:IValidator<String,String>
    {
        public Regex regex;
        public CellValidator()
        {


        }

        public CellValidator(String matcher):this()
        {

            regex = new Regex(matcher);
        }

        public virtual Boolean? validate(String value, String matcher=null)
        {
            Boolean? result = null;


            if (matcher.StartsWith("^"))
                result = (String.IsNullOrWhiteSpace(matcher)?regex: new Regex(matcher)).IsMatch(value ?? "");
            else
            {
                var marr = matcher.Split(",");
                switch (marr[0])
                {
                    case "date":
                        DateTime date, mdate = DateTime.Now.Date;

                        if (DateTime.TryParse(value, out date))
                            {
                            if (marr.Length == 1) result = true;
                            else
                            {
                                if ("now".Equals(marr[2]?.ToLower()) || DateTime.TryParse(marr[2], out mdate))
                                {
                                    if (marr.Length > 3 && new Regex(@"^[\+|\-]\d+$").IsMatch(marr[3])) mdate = mdate.AddBusinessDays(Int32.Parse(marr[3]));
                                    if (result != true && marr[1].Contains(">"))
                                        result = date.CompareTo(mdate) > 0;
                                    if (result != true && marr[1].Contains("="))
                                        result = date.CompareTo(mdate) == 0;
                                    if (result != true && marr[1].Contains("<"))
                                        result = date.CompareTo(mdate) < 0;
                                }
                            }
                        }

                        else result = false;
                        break;

                }
            }

            return result;
        }

    }
}
