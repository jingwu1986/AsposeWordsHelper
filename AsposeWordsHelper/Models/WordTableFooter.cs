using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsposeWordsHelper
{   
    public class WordTableFooter
    {
        public WordTableFooter(string content)
        {
            this.Content = content;
        }
     
        public string Content { get; set; }
       
        public bool IsHtml { get; set; }
       
        public bool ShowBorder { get; set; }
    }
}
