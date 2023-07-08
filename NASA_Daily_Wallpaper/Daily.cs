using Microsoft.Win32;
using System;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Threading.Tasks;

// TODO: Find explanation (how to sort out the links... very regEx mind bending). Possibly open website in browser if the user wants that.(?)

namespace DailyBackground
{
    class Daily
    {
        static async Task Main()
        {
            // Define the base URL for downloading the picture
            string baseURL = "https://apod.nasa.gov/apod/";
            Console.WriteLine("Downloading today's picture from " + baseURL + "astropix.html ...");

            // Download the HTML source code of the webpage
            string html = await DownloadSourceAsync(baseURL + "astropix.html");
            string picture;
            Image image;

            // Check if a picture is found in the HTML source code
            if ((picture = FindPicture(html)) != null)
            {
                // Extract the picture name, date, and explanation from the HTML source code
                string picName = FindName(html);
                DateTime date = FindDate(html);

                // Print the date and name of the picture
                PrintDateAndName(date, picName);

                // Extract the explanation from the HTML source code and format it to fit the desired line width
                string explanation = FindExplanation(html);
                explanation = FitToLineWidth(explanation, 119);

                // Print the formatted explanation
                Console.WriteLine("\nExplanation:\n" + explanation + "\n");

                // Download the image from the URL and save it
                if ((image = await DownloadImageAsync(baseURL + picture)) != null)
                {
                    string filePath = GetBackgroundPath(date.ToString("yyyy-MM-dd") + " " + picName + ".jpg");
                    SaveBackground(image, filePath);
                    SetBackground(PicturePosition.Fill, filePath);
                }
            }
            else
            {
                Console.WriteLine("No picture today!");
            }

            Console.WriteLine("Press any key to finish.");
            Console.ReadKey();
        }

        // Find the explanation from the HTML source code using regular expressions
        private static string FindExplanation(string html)
        {
            string pattern = @"\<b\> Explanation\: \<\/b\>((\n*|.*)*)\<p\>";
            Match m = Regex.Match(html, pattern, RegexOptions.IgnoreCase);
            if (!m.Success)
                return "";

            string explanationHtml = m.Groups[1].Value;
            pattern = @"(\<[^\>]*\>)";
            string explanation = Regex.Replace(explanationHtml, pattern, "");
            pattern = @"(\n)";
            explanation = Regex.Replace(explanation, pattern, " ");
            pattern = @"(\s+)";
            explanation = Regex.Replace(explanation, pattern, " ");
            return explanation.Trim();
        }

        // Fit the explanation text to the specified line width by adding line breaks
        private static string FitToLineWidth(string explanation, int lineWidth)
        {
            string[] words = explanation.Split(' ');
            var sb = new StringBuilder();
            int len = 0;

            foreach (string word in words)
            {
                if (len + word.Length > lineWidth)
                {
                    sb.Replace(' ', '\n', sb.Length - 1, 1);
                    len = 0;
                }

                sb.Append(word).Append(' ');
                len += word.Length + 1;
            }

            return sb.ToString().Trim();
        }

        // Download the source code of a webpage asynchronously
        private static async Task<string> DownloadSourceAsync(string page)
        {
            using (var httpClient = new HttpClient())
            {
                return await httpClient.GetStringAsync(page);
            }
        }

        // Find the picture URL from the HTML source code using regular expressions
        private static string FindPicture(string html)
        {
            string pattern = @"<a href=""(image\/.*(\.jpg|\.png|\.gif))""";
            Match m = Regex.Match(html, pattern, RegexOptions.IgnoreCase);
            if (m.Success) return m.Groups[1].Value;
            return null;
        }

        // Find the date of the picture from the HTML source code using regular expressions
        static DateTime FindDate(string html)
        {
            string pattern = @"\b(\d{4} \w+ \d{1,2})\b";
            string match = Regex.Match(html, pattern).Groups[1].Value; ;
            DateTime picDateTime;
            if (DateTime.TryParseExact(match, "yyyy MMMM d", new CultureInfo("en-US"), DateTimeStyles.AllowWhiteSpaces, out picDateTime))
                return picDateTime;
            else
                return DateTime.Now;
        }

        // Find the name of the picture from the HTML source code using regular expressions
        static string FindName(string html)
        {
            string pattern = @"<center>[\r|\n]+<b>(.*)\<\/";
            string title = Regex.Match(html, pattern, RegexOptions.IgnoreCase).Groups[1].Value;

            if (title.Length == 0) return "No Title";

            string bad = "\\/|<>?*\n" + '"';
            var sb = new StringBuilder();
            foreach (char c in title)
            {
                if (!bad.Contains(c.ToString()))
                {
                    if (c == ':') sb.Append(" -");
                    else sb.Append(c);
                }
            }
            return sb.ToString().Trim();
        }

        // Print the date and name of the picture
        private static void PrintDateAndName(DateTime date, string picName)
        {
            string dateString = date.ToString("D", new CultureInfo("en-US")) + ": ";
            Console.Write("\t");
            for (int i = 0; i < dateString.Length + picName.Length; i++) Console.Write("-");
            Console.WriteLine();
            Console.WriteLine("\t" + dateString + picName);
            Console.Write("\t");
            for (int i = 0; i < dateString.Length + picName.Length; i++) Console.Write("-");
            Console.WriteLine();
        }

        // Download an image from the specified URL asynchronously
        static async Task<Image> DownloadImageAsync(string URL)
        {
            try
            {
                using (var httpClient = new HttpClient())
                {
                    using (var stream = await httpClient.GetStreamAsync(URL))
                    {
                        return Image.FromStream(stream);
                    }
                }
            }
            catch
            {
                Console.WriteLine("FAILED!!");
                return null;
            }
        }

        // Get the path for saving the downloaded background image
        static string GetBackgroundPath(string fileName)
        {
            string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "NASA Wallpapers");
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, fileName);
        }

        // Save the downloaded background image
        static void SaveBackground(Image background, string filePath)
        {
            Console.WriteLine("Saving picture as: " + filePath);
            background.Save(filePath, System.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // Set the background image with the specified style and file path
        public static void SetBackground(PicturePosition style, string filePath)
        {
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

        // Enumeration for different picture positions
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
                string lpvParam,
                int fuWinIni);
        }
    }
}
