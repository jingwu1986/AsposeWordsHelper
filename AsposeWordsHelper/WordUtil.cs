using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Reflection;

namespace AsposeWordsHelper
{
    public class WordUtil
    {
        public static bool IsEmptyColor(Color color)
        {
            return color.Name == "0";
        }

        public static bool IsEmptyFontSize(double size)
        {
            return size == 0f;
        }

        public static Color PickColor(params Color[] colors)
        {
            if(colors!=null)
            {
                foreach(Color item in colors)
                {
                    if (!IsEmptyColor(item))
                    {
                        return item;
                    }
                }
            }
            
            return Color.Black;
        }

        public static double PickValue(params double[] values)
        {
            if (values != null)
            {
                foreach (double item in values)
                {
                    if (!IsEmptyFontSize(item))
                    {
                        return item;
                    }
                }
            }
            return 0;
        }

        public static Dictionary<string,object> ConvertToDictionary(dynamic obj)
        {          
            if(obj!=null)
            {
                if(obj is Dictionary<string,object>)
                {
                    return obj;
                }

                Dictionary<string, object> dict = new Dictionary<string, object>();

                IEnumerable<PropertyInfo> properties = obj.GetType().GetProperties();
                foreach (PropertyInfo property in properties)
                {
                    object value= property.GetValue(obj, null);
                    dict.Add(property.Name, value);
                }
                return dict;
            }

            return null;
        }       

        public static double CalculateTableColumnWidth(WordTable table, WordTableColumn column)
        {
            if(column.FixedWidth)
            {
                return column.Width;
            }

            double tableWidth = table.ActualWidth;

            if(table.FitColumns)
            {
                double totalWidth = table.Columns.Sum(item=>item.Width);
                if(totalWidth==0)
                {
                    totalWidth = tableWidth;
                }
                return Math.Round(column.Width / totalWidth * tableWidth,2);
            }
            else
            {
                if(table.ColumnWidthIsPercent)
                {
                    return Math.Round(tableWidth * column.Width / 100, 2);
                }
                else
                {
                    return column.Width;
                }
            }
        }

        public static int GetWordNodeLevel(WordNode node)
        {
            if(node.Parent==null)
            {
                return 1;
            }
            else
            {
                int level = GetWordNodeLevel(node, 1);                
                return level;
            }
        }

        private static int GetWordNodeLevel(WordNode node, int level)
        {
            if(node?.Parent!=null)
            {
               return GetWordNodeLevel(node.Parent, ++level);
            }
            return level;
        }

        public static string GetWordLevelTitleNumber(int level, int order, WordTitleNode parentNode, AutoNumberType autoNumberType)
        {
            if(level==1)
            {
                return NumberToOtherChar(order, autoNumberType);
            }
            else if(level>1 && parentNode!=null)
            {               
                string parentNumber = parentNode.Number;
                if(parentNode.Level==1)
                {
                    parentNumber = OtherCharToNumber(parentNumber, autoNumberType);
                }
                return $"{parentNumber}.{order}";
            }
            return string.Empty;
        } 
        
        private static string NumberToOtherChar(int number, AutoNumberType autoNumberType)
        {
            switch(autoNumberType)
            {               
                case AutoNumberType.CN:
                    return NumberHelper.Number2CN(number.ToString());
                case AutoNumberType.Roman:
                    return NumberHelper.Number2Roman(number);
                default:
                    return number.ToString();
            }
        }

        private static string OtherCharToNumber(string number, AutoNumberType autoNumberType)
        {
            switch(autoNumberType)
            {
                case AutoNumberType.CN:
                    return NumberHelper.GetNumberOfCNChar(number);
                case AutoNumberType.Roman:
                    return NumberHelper.GetNumberOfRomanChar(number);
                default:
                    return number;
            }           
        }
        
        public static WordTableMergeOption GetTableMergeOption(WordTable table, int rowIndex, int columnIndex)
        {
            return table.MergeOptions.FirstOrDefault(item=>rowIndex>= item.StartRowIndex && rowIndex<=item.StartRowIndex+item.Rowspan-1 && columnIndex>=item.StartColumnIndex && columnIndex<=item.StartColumnIndex+item.Colspan-1);
        }
    }
}
