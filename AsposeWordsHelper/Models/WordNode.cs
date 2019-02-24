using System.Collections.Generic;

namespace AsposeWordsHelper
{
    public class WordNode
    {
        public string Name { get; set; }
        public List<WordNode> Children { get; } = new List<WordNode>();
        public WordNode Parent { get; set; }

        public WordNode() { }

        public WordNode(string name, params WordNode[] content)
        {
            this.Name = name;
            if (content != null)
            {
                this.Children.AddRange(content);
                this.Children.ForEach(item =>
                {
                    item.Parent = this;
                });
            }
        }
    }

    #region Title
    public class WordTitleNode : WordNode
    {
        public WordTitleNode() { }

        public WordTitleNode(string name, WordTitleOption option = null, params WordNode[] content) : base(name, content)
        {
            this.Option = option == null ? new WordTitleOption() : option;
        }

        public WordTitleOption Option { get; set; } = new WordTitleOption();
       
        public int Level { get; internal set; }
       
        public int Order { get; set; }
      
        public string Number { get; set; }
    }
    #endregion

    #region Break
    public class WordParagraphBreakNode : WordNode
    {
    }

    public class WordLineBreakNode : WordNode
    {
    }

    public class WordPageBreakNode : WordNode
    {
    }
    #endregion

    #region Paragraph
    public class WordParagraphNode : WordNode
    {
        public WordParagraphNode(string content, WorldParagraphOption option = null)
        {
            this.Content = content;
            this.Optiton = option == null ? new WorldParagraphOption() : option;
        }
        public string Content { get; set; }
        public WorldParagraphOption Optiton { get; set; } = new WorldParagraphOption();
    }
    #endregion

    #region Table
    public class WordTableNode : WordNode
    {
        public WordTableNode() { }
        public WordTableNode(WordTable content)
        {
            this.Content = content;
        }
        public WordTable Content { get; set; }
    }
    #endregion

    #region Ole Object
    public class WordOleObjectNode : WordNode
    {
        public WordOleObject Content { get; set; }

        public WordOleObjectNode(WordOleObject oleObject)
        {
            this.Content = oleObject;
        }
    }
    #endregion

    #region Image
    public class WordImageNode : WordNode
    {
        public WordImage Content { get; set; }

        public WordImageNode() { }
        public WordImageNode(WordImage image)
        {
            this.Content = image;
        }
    } 
    #endregion
}
