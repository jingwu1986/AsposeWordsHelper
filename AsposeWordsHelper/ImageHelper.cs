using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsposeWordsHelper
{
    public class ImageHelper
    {
        public static Image ResizeImage(Image img, int width, int height, bool keepHeight=false)
        {
            int sourceWidth = img.Width;
            int sourceHeight = img.Height;

            float nPercent = 0;
            float nPercentW = 0;
            float nPercentH = 0;

            nPercentW = (width / (float)sourceWidth);
            nPercentH = (height / (float)sourceHeight);

            if (nPercentH < nPercentW)
            {
                nPercent = nPercentH;
            }
            else
            {
                nPercent = nPercentW;
            }

            int destWidth = (int)(sourceWidth * nPercent);
            int destHeight = keepHeight? height : (int)(sourceHeight * nPercent);

            Bitmap b = new Bitmap(destWidth, destHeight);
            Graphics g = Graphics.FromImage((Image)b);
            //g.SmoothingMode = SmoothingMode.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            //g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            g.DrawImage(img, 0, 0, destWidth, destHeight);
            g.Dispose();

            return (Image)b;
        }
    }
}
