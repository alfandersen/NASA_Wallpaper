using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Background
{
    class Program
    {
        static String date;
        static String baseURL = "http://apod.nasa.gov/apod/";
        static String picName;
        static Boolean preMillennial;

        static void Main()
        {
            while (true)
            {
                date = DateTime.Today.ToString("yyMMdd");

                Console.WriteLine("------------------------------------------------------------");
                Console.WriteLine("NASA has posted a new picture - occasionally video which cannot be set as wallpaper - each day since the 16th of June 1995. How awesome is that!!?");
                Console.WriteLine("This program fetches an image of the day from " + baseURL + "archivepix.html");
                Console.WriteLine("------------------------------------------------------------");
                Console.WriteLine("Input commands:");
                //OBS: input format is reversed compared to NASA's. I find DDMMYY easier to use and convert it in isVAlidDate() to NASA's YYMMDD format.
                Console.WriteLine(" - 'DDMMYY'            Eg. 140315 for the 14th of March 2015");
                Console.WriteLine(" - 'r' or 'random'     Gives you a random image (Not yet implemented)");
                Console.WriteLine(" - ''                  No input will give you today's image");
                Console.WriteLine(" - 'q' or 'quit'       Quits the program");

                String input = Console.ReadLine();

                if (input == "")
                {
                    preMillennial = false;
                    Console.WriteLine("Attempting to download today's picture");

                    Boolean success = setPicture(baseURL, date);
                    if (success) break;
                }
                else if (isValidDate(ref input, date))
                {
                    Console.WriteLine("Attempting to download picture for " + input);
                    date = input;

                    Boolean success = setPicture(baseURL, date);
                    if(success) break;
                }
                else if (input == "q" || input == "quit" || input == "Q" || input == "QUIT")
                {
                    Console.WriteLine("Quitting");
                    break;
                }
                else
                {
                    Console.WriteLine(input + " is an invalid input. Try again.");
                }
                Console.WriteLine();
            }
            //Console.ReadLine();
        }

        public static Boolean setPicture(String baseURL, String date)
        {
            String pathExpr = @"\bimage\S+(.jpg|.gif|.png)\b";
            Regex reg = new Regex(pathExpr);
            String htmlCode;
            String filePath;
            String pictureURL;
            WebClient client = new WebClient();
            Console.WriteLine("------URL: " + baseURL + "ap" + date + ".html");
            try
            {
                htmlCode = client.DownloadString(baseURL + "ap" + date + ".html");
                //Console.WriteLine("HTMLcode = " + htmlCode);
                filePath = reg.Match(htmlCode).Value;
                pictureURL = baseURL + filePath;

                String nameExpr = @"\w+\.\b";
                reg = new Regex(nameExpr);
                picName = reg.Match(filePath).Value;
                Console.WriteLine("--------- Filename: " + picName);
                Image background = downloadBackground(pictureURL);


                saveBackground(background);
                setBackground(PicturePosition.Fill);
                return true;
            }
            catch (WebException e)
            {
                Console.WriteLine("Error fetching image: " + e);
                return false;
            }
        }

        public static Boolean isValidDate(ref String input, String today)
        {
            Boolean ok = false;
            if (input.Length == 6)
            {
                int thisYear = Int32.Parse(today[0].ToString() + today[1].ToString());
                int thisMonth = Int32.Parse(today[2].ToString() + today[3].ToString());
                int thisDay = Int32.Parse(today[4].ToString() + today[5].ToString());

                int year;
                int month;
                int day;

                if (!Int32.TryParse((input[0].ToString() + input[1].ToString()), out day)) return false;
                if (!Int32.TryParse((input[2].ToString() + input[3].ToString()), out month)) return false;
                if (!Int32.TryParse((input[4].ToString() + input[5].ToString()), out year)) return false;

                preMillennial = (year >= 95);

                if(year == thisYear)
                {
                         if (month == thisMonth)    if (day <= thisDay) ok = true;
                    else if (month < thisMonth)     if (day <= 31)      ok = true;
                }
                else if(year < thisYear || year >= 96)
                {
                         if (month <= 12)           if (day <= 31)      ok = true;
                }
                else if (year == 95)
                {
                         if (month == 6)            if (day >= 16)      ok = true;
                    else if (month <= 12)           if (day <= 31)      ok = true;
                }

                if (ok) //Convert: DDMMYY --> YYMMDD
                    input = input[4].ToString()
                        + input[5].ToString()
                        + input[2].ToString()
                        + input[3].ToString()
                        + input[0].ToString()
                        + input[1].ToString();
            }
            return ok;
        }

        public static Image downloadBackground(String URL)
        {
            Console.WriteLine("Downloading: " + URL);
            Console.WriteLine(" ...");
            HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(URL);
            HttpWebResponse httpWebReponse = (HttpWebResponse)httpWebRequest.GetResponse();
            Stream stream = httpWebReponse.GetResponseStream();
            Image background = Image.FromStream(stream);
            Console.WriteLine("Downloaded done!");
            return background;
        }

        public static String getBackgroundPath()
        {
            String directory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "/NASA Wallpapers/";
            Directory.CreateDirectory(directory);
            String picDate = preMillennial ? "19" : "20";
            picDate += date[0].ToString() + date[1].ToString() + "-" + date[2].ToString() + date[3].ToString() + "-" + date[4].ToString() + date[5].ToString();
            return Path.Combine(directory, picDate + "_" + picName + "jpg");
        }

        public static Boolean saveBackground(Image background)
        {
            try
            {
                String localPath = getBackgroundPath();
                Console.WriteLine("Saving picture as: " + localPath);
                background.Save(localPath, System.Drawing.Imaging.ImageFormat.Jpeg);
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

        public static void setBackground(PicturePosition style)
        {
            Console.WriteLine("Setting background...");
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
            NativeMethods.SystemParametersInfo(SET_DESKTOP_BACKGROUND, 0, getBackgroundPath(), UPDATE_INI_FILE | SEND_WINDOWS_INI_CHANGE);
            Console.WriteLine("Set background!\n");
        }

    }
}