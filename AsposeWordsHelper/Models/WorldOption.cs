using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Aspose.Words;

namespace AsposeWordsHelper
{
    public class WordDocumentOption
    {
        public bool NeedTableOfContents { get; set; } = false;
        public SaveFormat SaveFormat { get; set; } = SaveFormat.Docx;
        public PageSetup PageSetup { get; set; }
        public string FontName { get; set; } = "宋体";
        public double FontSize { get; set; } = 12;
        public double TableRowHeight { get; set; } = 20;
        public double LineSpacing { get; set; } = 13;
        public double PageWidth { get; set; } = 786;
        public bool ParagraphLeftIndentAsBefore { get; set; } = true;

        /// <summary>
        /// Default Width(%)。
        /// </summary>
        public double TableDefaultWidth { get; set; } = 90;

        /// <summary>
        /// Whether insert a ParagraphBreak node after all child nodes inserted.
        /// </summary>
        public bool InsertParagraphBreakWhenChildrenEnd { get; set; }
        public double ParagraphFirstLineIndent { get; set; }

        public Color TableHeaderCellBackgroundColor = ColorTranslator.FromHtml("#EEEEEE");
        public bool AutoNumberTitle { get; set; }
        public AutoNumberType AutoNumberType { get; set; } = AutoNumberType.EN;

        public double DefaultLeftIndent = 10;          
        public double TableFontSize { get; set; }
    }

    public enum AutoNumberType
    {
        EN,
        CN,
        Roman
    }
  
    public class WordTitleOption
    {       
        public double FontSize { get; set; }       
        public ParagraphAlignment Alignment { get; set; } = ParagraphAlignment.Left;       
        public StyleIdentifier StyleIdentifier { get; set; } = StyleIdentifier.Heading1;       
        public double LeftIndent { get; set; }       
        public bool Bold { get; } = true;      
        public bool Italic { get; set; }
    }
 
    public class WorldParagraphOption
    {       
        public double FontSize { get; set; } = 12;       
        public ParagraphAlignment Alignment { get; set; } = ParagraphAlignment.Left;       
        public double FirstLineIndent { get; set; }      
        public double LeftIndent { get; set; }       
        public bool IsHtml { get; set; }       
        public bool Bold { get; set; }       
        public bool Italic { get; set; }      
        public bool HtmlUseBuilderFormatting { get; set; }
    }    
}
