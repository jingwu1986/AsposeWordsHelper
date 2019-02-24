using System.Drawing;
using System.IO;

namespace AsposeWordsHelper
{
    public class WordImage
    {
        public string FilePath { get; set; }
        public Stream Stream { get; set; }
        public byte[] Bytes { get; set; }
        public Image Image { get; set; }
        public WordImageType Type { get; set; }

        /// <summary>
        /// Whether auto fit width of document(if exceed).
        /// </summary>
        public bool FitPageWidth { get; set; } = true;

        /// <summary>
        /// Whether delete file after used(usually, set it to true if the file is temporary).
        /// </summary>
        public bool DeleteFileAfterUsed { get; set; }

        public WordImage(string filePath)
        {
            this.FilePath = filePath;
            this.Type = WordImageType.FilePath;            
        }

        public WordImage(Stream stream)
        {
            this.Stream = stream;
            this.Type = WordImageType.Stream;          
        }

        public WordImage(byte[] bytes)
        {
            this.Bytes = bytes;
            this.Type = WordImageType.Bytes;
        }

        public WordImage(Image image)
        {
            this.Image = image;
            this.Type = WordImageType.Image;
        }

        public Image GetImage()
        {
            switch(this.Type)
            {
                case WordImageType.Image:
                    return this.Image;
                case WordImageType.FilePath:
                    return Image.FromFile(this.FilePath);
                case WordImageType.Stream:
                    return Image.FromStream(this.Stream);
                case WordImageType.Bytes:
                    return Image.FromStream(new MemoryStream(this.Bytes));
            }
            return null;
        }
    }

    public enum WordImageType
    {
        FilePath=1,
        Stream=2,
        Bytes=3,
        Image=4
    }
}
