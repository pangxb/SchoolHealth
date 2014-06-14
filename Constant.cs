using System;
using System.Collections.Generic;
using System.Text;

namespace SchoolHealth
{
    public static class Constant
    {
        public static readonly string Grade = "年级";
        public static Dictionary<string,List<string>> Grades {private set;get;}


        static Constant()
        {
            Grades = new Dictionary<string, List<string>>();
            Grades[Grade] = new List<String>()
            {
                   "一年级","二年级","三年级","四年级","五年级","六年级","七年级","八年级","九年级","高一","高二","高三",
            };
         


        }
    }
}
