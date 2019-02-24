using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AsposeWordsHelper
{
    public class NumberHelper
    {       
        public static readonly string CN_CHAR = "一二三四五六七八九";
        public static readonly string NUMBER_CHAR = "123456789";
        public static readonly string CN_TEN = "十";
        public static readonly string NUMBER_TEN = "10";

        public static string CN2Number(string value)
        {
            if(value.Length==1)
            {
                return GetNumberOfCNChar(value);
            }
            else if(value.Length>=2)
            {
                if(value.StartsWith(CN_TEN))
                {
                    return "1" + GetNumberOfCNChar(value.Substring(1,1));
                }               
                else
                {
                    int index = value.IndexOf(CN_TEN);
                    if(index>0)
                    {
                        string[] items = value.Split(new string[] { CN_TEN }, StringSplitOptions.RemoveEmptyEntries);

                        if (items.Length == 1)
                        {
                            return GetNumberOfCNChar(items[0]) + "0";
                        }
                        else if (items.Length == 2)
                        {
                            return GetNumberOfCNChar(items[0]) + GetNumberOfCNChar(items[1]);
                        }
                    }                                 
                }
            }
            return value;
        }        

        public static string Number2CN(string value)
        {
            if(value.Length==1)
            {
                return GetCNCharOfNumber(value);
            }
            else if(value.Length>=2)
            {
                if(value.Length==2)
                {
                    string firstChar = value[0].ToString();
                    return (firstChar=="1"? "" : GetCNCharOfNumber(firstChar)) + CN_TEN + GetCNCharOfNumber(value[1].ToString());
                }               
            }
            return value;
        }

        public static string GetNumberOfCNChar(string value)
        {
            int index = CN_CHAR.IndexOf(value);
            if (index >= 0)
            {
                return NUMBER_CHAR[index].ToString();
            }
            else if (value == CN_TEN)
            {
                return NUMBER_TEN;
            }
            return string.Empty;
        }

        public static string GetCNCharOfNumber(string value)
        {
            int index = NUMBER_CHAR.IndexOf(value);
            if (index >= 0)
            {
                return CN_CHAR[index].ToString();
            }
            else if (value == NUMBER_TEN)
            {
                return CN_TEN;
            }
            return string.Empty;
        }

        public static string Number2Roman(int number)
        {
            if ((number < 0) || (number > 3999)) throw new ArgumentOutOfRangeException("insert value betwheen 1 and 3999");
            if (number < 1) return string.Empty;
            if (number >= 1000) return "M" + Number2Roman(number - 1000);
            if (number >= 900) return "CM" + Number2Roman(number - 900);
            if (number >= 500) return "D" + Number2Roman(number - 500);
            if (number >= 400) return "CD" + Number2Roman(number - 400);
            if (number >= 100) return "C" + Number2Roman(number - 100);
            if (number >= 90) return "XC" + Number2Roman(number - 90);
            if (number >= 50) return "L" + Number2Roman(number - 50);
            if (number >= 40) return "XL" + Number2Roman(number - 40);
            if (number >= 10) return "X" + Number2Roman(number - 10);
            if (number >= 9) return "IX" + Number2Roman(number - 9);
            if (number >= 5) return "V" + Number2Roman(number - 5);
            if (number >= 4) return "IV" + Number2Roman(number - 4);
            if (number >= 1) return "I" + Number2Roman(number - 1);
            throw new ArgumentOutOfRangeException("something bad happened");
        }

        public static string GetNumberOfRomanChar(string value)
        {
            string[] replaceRom = { "CM", "CD", "XC", "XL", "IX", "IV" };
            string[] replaceNum = { "DCCCC", "CCCC", "LXXXX", "XXXX", "VIIII", "IIII" };
            string[] roman = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            int[] arabic = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            return Enumerable.Range(0, replaceRom.Length)
                .Aggregate
                (
                    value,
                    (agg, cur) => agg.Replace(replaceRom[cur], replaceNum[cur]),
                    agg => agg.ToArray()
                )
                .Aggregate
                (
                    0,
                    (agg, cur) =>
                    {
                        int idx = Array.IndexOf(roman, cur.ToString());
                        return idx < 0 ? 0 : agg + arabic[idx];
                    },
                    agg => agg
                ).ToString();
        }
    }      
}
