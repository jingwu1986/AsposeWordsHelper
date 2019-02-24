using System.Collections.Generic;

namespace AsposeWordsHelper
{
    public class WordTable
    {
        public WordTable() { }

        public WordTable(List<WordTableColumn> columns, dynamic data)
        {
            this.Columns = columns;
            this.Data = data;
        }

        public List<WordTableColumn> Columns { get; set; } = new List<WordTableColumn>();
        public dynamic Data = new List<dynamic>();
        public List<WordTableMergeOption> MergeOptions = new List<WordTableMergeOption>();
        public WordTableFooter Footer { get; set; }

        #region Configs       
        public bool FitColumns { get; set; } = true;      

        public bool AllowAutoFit { get; set; } = false;    

        public double Width { get; set; }
        public double ActualWidth { get; internal set; }       

        public bool TableWidthIsPercent { get; set; } = true;
     
        public bool ColumnWidthIsPercent { get; set; } = true;
     
        public bool ShowHeader { get; set; } = true;
     
        public bool AutoRowHeight { get; set; } = true;
       
        public double RowHeight { get; set; } = 20;

        public double FontSize { get; set; } = 10;   
        
        public bool StrictToUseColumnWidth { get; set; }
        #endregion
    }

    public class WordTableMergeOption
    {      
        public int StartRowIndex { get; set; } = -1;           
      
        public int StartColumnIndex { get; set; } = -1;

        public int Colspan { get; set; } = 1;
       
        public int Rowspan { get; set; } = 1;
    }
}
