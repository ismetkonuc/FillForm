using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Policy;
using System.Text;

namespace BitMiracle.Docotic.Pdf.Samples
{
    // bu class google drive dışındaki bir siteden fotoğraf indiriyor şimdilik kullanılmıyor.
    public static class DownloadImage
    {
        public static Bitmap SaveImage(string filename, string imageUrl)
        {
            ImageFormat format = FindImageFormat(imageUrl);
            WebClient client = new WebClient();

            Stream stream = Stream.Null;
            try
            {
                stream = client.OpenRead(imageUrl);
                
            }
            catch (WebException e)
            {
                Console.WriteLine("404 Hatası: Fotoğraf Bulunamadı:" + imageUrl);
                Console.WriteLine(e.Response);
                return null;
            }

            Bitmap img; 
            img = new Bitmap(stream);
            if (img != null)
            {
                img.Save(filename, format);
            }

            stream.Flush();
            stream.Close();
            client.Dispose();
            return img;
        }

        private static ImageFormat FindImageFormat(string imageUrl)
        {
            ImageFormat image;

            string format = imageUrl.Substring(imageUrl.Length-4);

            if (format.Contains("j"))
            {
                image = ImageFormat.Jpeg; 
            }

            else if (format.Contains("png"))
            {
                image = ImageFormat.Png;
            }

            else
            {
                return null;
            }
            return image;
        }
    }
}
