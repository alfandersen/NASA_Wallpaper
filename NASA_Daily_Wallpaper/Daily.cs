using Microsoft.Win32;
using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Globalization;

//TODO: Find explanation (how to sort out the links... very regEx mind bending). Possibly open website in browser if the user wants that.(?)

namespace DailyBackground
{
    class Daily
    {
        

        static void Main()
        {
            String baseURL = "https://apod.nasa.gov/apod/";
            Console.WriteLine("Downloading todays picture from "+baseURL+ "astropix.html ...");

            String html = downloadSource(baseURL+"astropix.html");
            String picture;
            Image image;
            if((picture = findPicture(html)) != null && (image = downloadImage(baseURL+picture)) != null)
            {
                String picName = findName(html);
                DateTime date = findDate(html);
                printDateAndName(date, picName);
                String filePath = getBackgroundPath(date.ToString("yyyy-MM-dd") + " " + picName + ".jpg");
                saveBackground(image, filePath);
                setBackground(PicturePosition.Fill, filePath);
            }

            /*
            Console.WriteLine("Press any key to finish.");
            Console.ReadKey();
            //*/
        }

        private static string downloadSource(string page)
        {
            return new WebClient().DownloadString(page);
        }

        private static string findPicture(string html)
        {
            String pattern = @"<img src=""(.*(\.jpg|.png|.gif))""";
            Match m = Regex.Match(html, pattern, RegexOptions.IgnoreCase);
            return m.Groups[1].Value;
        }

        static DateTime findDate(String html)
        {
            String pattern = @"\b(\d{4} \w+ \d{1,2})\b";
            String match = Regex.Match(html, pattern).Groups[1].Value; ;
            DateTime picDateTime;
            if (DateTime.TryParseExact(match, "yyyy MMMM d", new CultureInfo("en-US"), DateTimeStyles.AllowWhiteSpaces, out picDateTime))
                return picDateTime;
            else
                return DateTime.Now;
        }

        static String findName(String html)
        {
            String pattern = @"<center>[\r|\n]+<b>(.*)\<\/";
            String title = Regex.Match(html, pattern, RegexOptions.IgnoreCase).Groups[1].Value;

            if (title.Length == 0) return "No Title";

            String bad = "\\/|<>?*\n" + '"';
            String name = "";
            foreach(char c in title.ToCharArray())
            {
                if (!bad.Contains(c.ToString()))
                {
                    if (c == ':') name += " -";
                    else name += c;
                }
            }
            return name.Trim();
        }

        private static void printDateAndName(DateTime date, string picName)
        {
            String dateString = date.ToString("D", new CultureInfo("en-US")) + ": ";
            Console.Write("\t");
            for (int i = 0; i < dateString.Length + picName.Length; i++) Console.Write("-");
            Console.WriteLine();
            Console.WriteLine("\t" + dateString + picName);
            Console.Write("\t");
            for (int i = 0; i < dateString.Length + picName.Length; i++) Console.Write("-");
            Console.WriteLine();
        }

        static Image downloadImage(String URL)
        {
            HttpWebResponse httpWebReponse;
            Console.WriteLine("Downloading: " + URL);
            Console.WriteLine(" ...");
            try
            {
                HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(URL);
                httpWebReponse = (HttpWebResponse)httpWebRequest.GetResponse();
            }
            catch
            {
                Console.WriteLine("FAILED!!");
                return null;
            }
            Stream stream = httpWebReponse.GetResponseStream();
            Image image = Image.FromStream(stream);
            stream.Close();
            Console.WriteLine("Download done!");
            return image;
        }

        static String getBackgroundPath(String fileName)
        {
            String directory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "/NASA Wallpapers/";
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, fileName);
        }

        static Boolean saveBackground(Image background, String filePath)
        {
            try
            {
                Console.WriteLine("Saving picture as: " + filePath);
                background.Save(filePath, System.Drawing.Imaging.ImageFormat.Jpeg);
                return true;
            }
            catch { return false; }
        }

        public enum PicturePosition
        {
            Tile, Center, Stretch, Fit, Fill
        }

        internal sealed class NativeMethods
        {
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            internal static extern int SystemParametersInfo(
                int uAction,
                int uParam,
                String lpvParam,
                int fuWinIni);
        }

        public static void setBackground(PicturePosition style, String filePath)
        {
            //Console.WriteLine("Setting background...");
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true);
            switch (style)
            {
                case PicturePosition.Tile:
                    key.SetValue(@"PicturePosition", "0");
                    key.SetValue(@"TileWallpaper", "1");
                    break;
                case PicturePosition.Center:
                    key.SetValue(@"PicturePosition", "0");
                    key.SetValue(@"TileWallpaper", "0");
                    break;
                case PicturePosition.Stretch:
                    key.SetValue(@"PicturePosition", "2");
                    key.SetValue(@"TileWallpaper", "0");
                    break;
                case PicturePosition.Fit:
                    key.SetValue(@"PicturePosition", "6");
                    key.SetValue(@"TileWallpaper", "0");
                    break;
                case PicturePosition.Fill:
                    key.SetValue(@"PicturePosition", "10");
                    key.SetValue(@"TileWallpaper", "0");
                    break;
            }
            key.Close();

            const int SET_DESKTOP_BACKGROUND = 20;
            const int UPDATE_INI_FILE = 1;
            const int SEND_WINDOWS_INI_CHANGE = 2;
            NativeMethods.SystemParametersInfo(SET_DESKTOP_BACKGROUND, 0, filePath, UPDATE_INI_FILE | SEND_WINDOWS_INI_CHANGE);
            Console.WriteLine("Wallpaper set and done!");
        }
    }
}