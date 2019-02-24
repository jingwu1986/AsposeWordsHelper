using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

namespace AsposeWordsHelper
{
    public class WordTableColumn
    {
        public WordTableColumn() { }
        public WordTableColumn(string title)
        {
            this.Title = title;
            this.Name = title;
        }
        public WordTableColumn(string title, string name)
        {          
            this.Title = title;
            this.Name = name;
        }      
     
        public string Name { get; set; }      
        public string Title { get; set; }
        public double FontSize { get; set; }
        public string FontName { get; set; }
        public CellVerticalAlignment VerticalAlignment { get; set; } = CellVerticalAlignment.Center;
        public ParagraphAlignment Alignment { get; set; } = ParagraphAlignment.Center;
        public double Width { get; set; }   
        public Color FontColor { get; set; } = Color.Black;
        public Color ValueFontColor { get; set; } = Color.Black;
        public Color BackgroundColor { get; set; }       
        public bool FixedWidth { get; set; }      
        public bool IsRowHeader { get; set; }      
        public double ActualWidth { get; internal set; }      
        public bool IsRowNumber { get; set; }       
        public bool AutoMerge { get; set; }
    }
}
