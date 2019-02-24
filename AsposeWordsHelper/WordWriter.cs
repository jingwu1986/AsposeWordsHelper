using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Aspose.Words;

namespace AsposeWordsHelper
{
    public class WordWriter
    {
        private List<WordNode> nodes = new List<WordNode>();
        private WordDocumentOption option = new WordDocumentOption();
        public string TemplateFile { get; set; }

        public WordWriter(WordDocumentOption option)
        {
            this.option = option==null? new WordDocumentOption():option;
        }
        public WordWriter(List<WordNode> nodes, WordDocumentOption option)
        {
            this.nodes = nodes;
            this.option = option == null ? new WordDocumentOption() : option;
        }       

        public void Write(Stream stream)
        {
            WordGenerator word = this.BuildWordNodes();
            word.SaveTo(stream);
        }

        public void Write(string filePath)
        {
            WordGenerator word = this.BuildWordNodes();
            word.SaveTo(filePath);
        }

        private WordGenerator BuildWordNodes()
        {
            WordGenerator word = new WordGenerator(this.TemplateFile, this.option);

            int order = 1;
            foreach (WordNode node in this.nodes)
            {
                if(node is WordTitleNode)
                {
                    WordTitleNode titleNode = node as WordTitleNode;
                    if(titleNode.Order==0)
                    {
                        titleNode.Order = order;
                        order++;
                    }                    
                }
                this.BuildWordNode(word, node);               
            }

            return word;
        }        

        public void BuildWordNode(WordGenerator word, WordNode node)
        {
            if(node is WordTitleNode)
            {
                WordTitleNode titleNode = node as WordTitleNode;

                int level = WordUtil.GetWordNodeLevel(node);
               
                titleNode.Level = level;

                string no = "";
                string seperator = "";
                if (word.Option.AutoNumberTitle)
                {
                    int order = titleNode.Order>0 ? titleNode.Order: CalculateTitleNodeOrder(titleNode);
                    titleNode.Order = order;

                    no = WordUtil.GetWordLevelTitleNumber(level, order, node.Parent as WordTitleNode, word.Option.AutoNumberType);
                    titleNode.Number = no;

                    titleNode.Option.LeftIndent = word.Option.DefaultLeftIndent * level;

                    seperator = level == 1 ? (word.Option.AutoNumberType==AutoNumberType.CN? "、":".") : " ";
                }               
              
                string title = $"{no}{seperator}{titleNode.Name}";
                word.AddTitle(title, titleNode.Option);
            }
            else if(node is WordParagraphBreakNode)
            {
                word.AddParagraphBreak();
            }
            else if(node is WordLineBreakNode)
            {
                word.AddLineBreak();
            }
            else if (node is WordPageBreakNode)
            {
                word.AddPageBreak();
            }
            else if (node is WordParagraphNode)
            {
                WordParagraphNode paragraphNode = node as WordParagraphNode;

                word.AddParagraph(paragraphNode.Content, paragraphNode.Optiton);
            }
            else if(node is WordTableNode)
            {
                WordTableNode tableNode = node as WordTableNode;
                word.AddTable(tableNode.Content);
            }
            else if(node is WordImageNode)
            {
                WordImageNode imageNode = node as WordImageNode;
                word.AddImage(imageNode.Content);
            }
            else if(node is WordOleObjectNode)
            {
                WordOleObjectNode oleObjectNode = node as WordOleObjectNode;
                word.AddOleObject(oleObjectNode.Content);
            }

            if(node.Children!=null && node.Children.Count>0)
            {
                foreach(WordNode childNode in node.Children)
                {
                    this.BuildWordNode(word, childNode);
                }

                if (word.Option.InsertParagraphBreakWhenChildrenEnd)
                {
                    word.AddParagraphBreak();                 
                }
            }                  
        }

       
        private int CalculateTitleNodeOrder(WordTitleNode titleNode)
        {
            if(titleNode?.Parent?.Children!=null)
            {
                int maxOrder = titleNode.Parent.Children.OfType<WordTitleNode>().Max(item=>item.Order);
                return maxOrder + 1;
            }
            return 0;
        }
    }
}
