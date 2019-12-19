using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Windows.Forms;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Net.Http;
using System.Net.Http.Headers;
using QRCodeEncoderLibrary;

namespace QRCodeAddInForMSCE
{
    public class QRCodeImageFunctions
    {
        [DataContract]
        private class ShortenThisURLWithAzure
        {
            [DataMember]
            public string Input { get; set; }
        }

        [DataContract]
        private class ShortenedURLResponseFromAzure
        {
            [DataMember]
            public string ShortUrl { get; set; }
            [DataMember]
            public string LongUrl { get; set; }
        }

        [DataContract]
        private class ShortenedURLResponseFromAzureContainer
        {
            [DataMember]
            public ShortenedURLResponseFromAzure data { get; set; }
        }

        /// <summary>
        /// Calls the my Azure API to shorten the long URL into an adddress like shl.pw/2v and returns QR Code image of the shortened URL
        /// </summary>
        /// <param name="sLongURL">URL to shorten and generate image of</param>
        /// <param name="iSize">Size of image to generate with ZXing</param>
        /// <param name="iCropBorder">Sometimes too much white space around QR Code and needs cropping</param>
        /// <returns>The (optionally cropped) QRCode Image of the shortened URL from ZXing</returns>
        public static Image GetQRCodeImage(string sLongURL, int iSize, int iCropBorder)
        {
            string url = "https://pwqrazurefunctionapp.azurewebsites.net/api/ShortenURLPost?code=bxwoszCCCjEGB2KTqiw/tcNCoNCKMdu9jdcbB2xiXcUoxPhMHWRTDA==";

            Image img = null;

            try
            {
                // build WSG file link
                var linkToShorten = new ShortenThisURLWithAzure
                {
                    Input = sLongURL
                };

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);
                HttpClient connectClient = new HttpClient();
                connectClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                var body = JSONHelper.SerializeJSon<ShortenThisURLWithAzure>(linkToShorten);
                request.Content = new StringContent(body, Encoding.UTF8, "application/json");

                var response = connectClient.SendAsync(request).Result;
                string responseContent = response.Content.ReadAsStringAsync().Result;

                BPSUtilities.WriteLog(responseContent);

                // this is dumb
                ShortenedURLResponseFromAzure shortUrl = JSONHelper.DeserializeJSon<ShortenedURLResponseFromAzure>(responseContent.Replace("[", "").Replace("]", ""));

                if (shortUrl != null)
                {
                    if (!string.IsNullOrEmpty(shortUrl.ShortUrl))
                    {
                        BPSUtilities.WriteLog($"'{shortUrl.LongUrl}' shortened to '{shortUrl.ShortUrl}'");

                        img = (Image)CreateQRCode(shortUrl.ShortUrl, iSize);

                        BPSUtilities.WriteLog($"Original QR Code size is {img.Width} x {img.Height}.");

                        if (img.Width > iSize)
                        {
                            Image img3 = ResizeImage(img, iSize, iSize, ImageFormat.Png);
                            img = img3;
                        }

                        BPSUtilities.WriteLog(string.Format("QR Code is {0} X {1}.", img.Width, img.Height, iCropBorder));
                    }
                }
                else
                {
                    BPSUtilities.WriteLog("Short URL not generated.");
                }
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog("Error: " + ex.Message);
                BPSUtilities.WriteLog(ex.StackTrace);
            }

            return img;
        }

        /// <summary>
        /// Calls the my Azure API to shorten the long URL into an adddress like shl.pw/2v and returns QR Code image of the shortened URL
        /// </summary>
        /// <param name="sLongURL">URL to shorten and generate image of</param>
        /// <param name="iSize">Size of image to generate with ZXing</param>
        /// <param name="iCropBorder">Sometimes too much white space around QR Code and needs cropping</param>
        /// <returns>The (optionally cropped) QRCode Image of the shortened URL from ZXing</returns>
        public static QREncoder GetQREncoderObject(string sLongURL)
        {
            string url = "https://pwqrazurefunctionapp.azurewebsites.net/api/ShortenURLPost?code=bxwoszCCCjEGB2KTqiw/tcNCoNCKMdu9jdcbB2xiXcUoxPhMHWRTDA==";

            try
            {
                // build WSG file link
                var linkToShorten = new ShortenThisURLWithAzure
                {
                    Input = sLongURL
                };

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, url);
                HttpClient connectClient = new HttpClient();
                connectClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                var body = JSONHelper.SerializeJSon<ShortenThisURLWithAzure>(linkToShorten);
                request.Content = new StringContent(body, Encoding.UTF8, "application/json");

                var response = connectClient.SendAsync(request).Result;
                string responseContent = response.Content.ReadAsStringAsync().Result;

                BPSUtilities.WriteLog(responseContent);

                // this is dumb
                ShortenedURLResponseFromAzure shortUrl = JSONHelper.DeserializeJSon<ShortenedURLResponseFromAzure>(responseContent.Replace("[", "").Replace("]", ""));

                if (shortUrl != null)
                {
                    if (!string.IsNullOrEmpty(shortUrl.ShortUrl))
                    {
                        BPSUtilities.WriteLog($"'{shortUrl.LongUrl}' shortened to '{shortUrl.ShortUrl}'");

                        QREncoder enc = new QREncoder();

                        enc.ErrorCorrection = ErrorCorrection.H;

                        // this leads to a 33 x 33 qrcode
                        enc.ModuleSize = 1;
                        enc.QuietZone = 4;

                        enc.Encode(shortUrl.ShortUrl);

                        return enc;
                    }
                }
                else
                {
                    BPSUtilities.WriteLog("Short URL not generated.");
                }
            }
            catch (Exception ex)
            {
                BPSUtilities.WriteLog("Error: " + ex.Message);
                BPSUtilities.WriteLog(ex.StackTrace);
            }

            return null;
        }


        private static void QRCodeImageToClipboard(string sLongURL, int iSize)
        {
            Image img = GetQRCodeImage(sLongURL, iSize, 0);

            if (img != null)
            {
                Clipboard.Clear();
                Clipboard.SetImage(img);
            }
        }

        /// <summary>
        /// Could return bigger bitmap
        /// </summary>
        /// <param name="sAddress"></param>
        /// <param name="iSize"></param>
        /// <returns></returns>
        private static Bitmap CreateQRCode(string sAddress, int iSize)
        {
            // Adding new QRCode encoding system

            int iMinimumSize = 33;

            iSize = Math.Max(iSize, iMinimumSize);

            QREncoder enc = new QREncoder();

            enc.ErrorCorrection = ErrorCorrection.H;

            // this leads to a 33 x 33 qrcode
            enc.ModuleSize = 1;
            enc.QuietZone = 4;

            int iModuleSize = Math.Max(1, (int)Math.Ceiling((double)iSize / (double)iMinimumSize));

            enc.QuietZone = 4 * iModuleSize;
            enc.ModuleSize = iModuleSize;

            enc.Encode(sAddress);

            return enc.CreateQRCodeBitmap();
        }

        /// <summary>
        /// Crops and resizes the image.
        /// </summary>
        /// <param name="img">The image to be processed</param>
        /// <param name="targetWidth">Width of the target</param>
        /// <param name="targetHeight">Height of the target</param>
        /// <param name="x1">The position x1</param>
        /// <param name="y1">The position y1</param>
        /// <param name="x2">The position x2</param>
        /// <param name="y2">The position y2</param>
        /// <param name="imageFormat">The image format</param>
        /// <returns>A cropped and resized image</returns>
        private static Image CropAndResizeImage(Image img, int targetWidth, int targetHeight, int x1, int y1, int x2, int y2, ImageFormat imageFormat)
        {
            var bmp = new Bitmap(targetWidth, targetHeight);
            Graphics g = Graphics.FromImage(bmp);

            g.InterpolationMode = InterpolationMode.NearestNeighbor; // .HighQualityBicubic;
            g.SmoothingMode = SmoothingMode.AntiAlias; // .HighQuality;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.CompositingQuality = CompositingQuality.HighQuality;

            int width = x2 - x1;
            int height = y2 - y1;

            g.DrawImage(img, new Rectangle(0, 0, targetWidth, targetHeight), x1, y1, width, height, GraphicsUnit.Pixel);

            var memStream = new MemoryStream();
            bmp.Save(memStream, imageFormat);
            return Image.FromStream(memStream);
        }

        /// <summary>
        /// Resizes the image.
        /// </summary>
        /// <param name="img">The image to be resized</param>
        /// <param name="targetWidth">Width of the target</param>
        /// <param name="targetHeight">Height of the target</param>
        /// <param name="imageFormat">The image format</param>
        /// <returns>A resized image</returns>
        private static Image ResizeImage(Image img, int targetWidth, int targetHeight, System.Drawing.Imaging.ImageFormat imageFormat)
        {
            return CropAndResizeImage(img, targetWidth, targetHeight, 0, 0, img.Width, img.Height, imageFormat);
        }

        /// <summary>
        /// Crops the image.
        /// </summary>
        /// <param name="img">The image</param>
        /// <param name="x1">The position x1.</param>
        /// <param name="y1">The position y1.</param>
        /// <param name="x2">The position x2.</param>
        /// <param name="y2">The position y2.</param>
        /// <param name="imageFormat">The image format.</param>
        /// <returns>A cropped image.</returns>
        private static Image CropImage(Image img, int x1, int y1, int x2, int y2, System.Drawing.Imaging.ImageFormat imageFormat)
        {
            return CropAndResizeImage(img, x2 - x1, y2 - y1, x1, y1, x2, y2, imageFormat);
        }
    }
}
