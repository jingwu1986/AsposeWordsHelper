using System.Drawing;
using System.IO;

namespace AsposeWordsHelper
{
    public class WordOleObject
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string FilePath { get; set; }
        public Stream Stream { get; set; }
       
        public bool ShowAsIcon { get; set; }
       
        public string IconFilePath { get; set; }
       
        public Image IconImage { get; set; }

        /// <summary>
        /// Whether delete file after used(usually, set it to true if the file is temporary).
        /// </summary>
        public bool DeleteFileAfterUsed { get; set; }

        public WordOleObjectType Type { get; set; }
      
        public WordOleObjectIconPosition IconPosition { get; set; } = WordOleObjectIconPosition.Top;
      
        public bool IconAsThumbnail { get; set; }
     
        public int ThumbnailIconWidth { get; set; } = 24;
      
        public int ThumbnailIconHeight { get; set; } = 24;
      
        public bool ShowAsDetailList { get; set; }
      
        public int LeftIndent { get; set; } = 20;

        public WordOleObject(string name, string filePath, bool showAsIcon = false, Image iconImage = null)
        {
            this.Name = name;
            this.FilePath = filePath;
            this.Type = WordOleObjectType.FilePath;
            this.Init(showAsIcon, iconImage);
        }

        public WordOleObject(string name, Stream stream, bool showAsIcon = false, Image iconImage = null)
        {
            this.Name = name;
            this.Stream = stream;
            this.Type = WordOleObjectType.Stream;
            this.Init(showAsIcon, iconImage);
        }

        private void Init(bool showAsIcon, Image iconImage)
        {
            this.ShowAsIcon = showAsIcon;
            this.IconImage = iconImage;
        }
    }

    public enum WordOleObjectType
    {
        FilePath=1,
        Stream=2
    }

    public enum WordOleObjectIconPosition
    {
        Top=0,
        Left=1
    }
}
