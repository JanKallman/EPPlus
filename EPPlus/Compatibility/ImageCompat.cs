using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Compatibility
{
    internal class ImageCompat
    {
        internal static byte[] GetImageAsByteArray(Image image)
        {
            var ms = new MemoryStream();
            if (image.RawFormat.Guid == ImageFormat.Gif.Guid)
            {
                image.Save(ms, ImageFormat.Gif);
            }
            else if (image.RawFormat.Guid == ImageFormat.Bmp.Guid)
            {
                image.Save(ms, ImageFormat.Bmp);
            }
            else if (image.RawFormat.Guid == ImageFormat.Png.Guid)
            {
                image.Save(ms, ImageFormat.Png);
            }
            else if (image.RawFormat.Guid == ImageFormat.Tiff.Guid)
            {
                image.Save(ms, ImageFormat.Tiff);
            }
            else
            {
                image.Save(ms, ImageFormat.Jpeg);
            }

            return ms.ToArray();
        }
    }
}
