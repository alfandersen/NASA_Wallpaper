using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Globalization;

//TODO: Find explanation (how to sort out the links... very regEx mind bending). Possibly open website in browser if the user wants that.(?)

namespace DailyBackground
{
    class Daily
    {
        
        static String baseURL = "https://apod.nasa.gov/apod/";
        static String date;
        static String picName;

        static void Main()
        {
            date = DateTime.Today.ToString("yyMMdd");

            Console.WriteLine("Downloading todays picture from "+baseURL+ "astropix.html ...");
            if (setPicture())
                Console.WriteLine("Success!!");
            else
                Console.WriteLine("Failed!!");

            Console.WriteLine("Press any key to finish.");
            Console.ReadKey();
        }
        
        static Boolean setPicture()
        {
            String htmlCode;
            String filePath;
            String pictureURL;
            WebClient client = new WebClient();
            Console.WriteLine("Fetching HTML code ...");
            try
            {
                htmlCode = client.DownloadString(baseURL+ "astropix.html");
                //Console.WriteLine("HTMLcode = " + htmlCode);
                String pathExpr = @"\bimage\S+(.jpg|.gif|.png)\b";
                Regex reg = new Regex(pathExpr);
                Match match = reg.Match(htmlCode);
                if (!match.Success)
                {
                    Console.WriteLine("Unfortunately no image was found today.");
                    return false;
                }
                filePath = match.Value;
                pictureURL = baseURL + filePath;

                picName = findName(htmlCode);

                DateTime picDateTime;
                String dateString = date;
                if (DateTime.TryParseExact(date, "yyMMdd", new CultureInfo("en-US"), DateTimeStyles.None, out picDateTime))
                    dateString = picDateTime.ToString("D", new CultureInfo("en-US"));
                dateString += ": ";
                Console.Write("\t");
                for (int i = 0; i < dateString.Length + picName.Length; i++)
                {
                    Console.Write("-");
                }
                Console.WriteLine();
                Console.WriteLine("\t" + dateString + picName);
                Console.Write("\t");
                for (int i = 0; i < dateString.Length + picName.Length; i++)
                {
                    Console.Write("-");
                }
                Console.WriteLine();

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

        static String findName(String htmlCode)
        {
            Regex reg;
            String name = "";
            String titleExpr = @"\-\s.+"; // The title comes right after "- " in the htmlCode
            reg = new Regex(titleExpr);
            String title = reg.Match(htmlCode).Value;
            reg = new Regex(@"\<.+");
            title = reg.Replace(title, ""); // The older html codes end with </title> on the same line. Get rid of that here.

            int titleLength = title.Length - 1;
            if (title[titleLength] == ' ') titleLength--;

            //title = @"all:them/bad > guys \ are noW ? able : to<be fu**ed uP by regEx!?#-@ " + '"'.ToString();
            String illegalCharsExpr = @"\:|\\|\/|\||\<|\>|\?|\*|\" + '"'.ToString();
            reg = new Regex(illegalCharsExpr);
            List<int> badGuys = new List<int>();
            foreach (Match badGuy in reg.Matches(title))
            {
                badGuys.Add(badGuy.Index);
            }
            for (int i = 1; i <= titleLength; i++)
            {   // Get rid of the "- " in the beginning and potential extra ending space + all the badguys that are not allowed in a file name
                if (!badGuys.Contains(i)) name += title[i].ToString();
            }
            return name;
        }

        static Image downloadBackground(String URL)
        {
            Console.WriteLine("Downloading: " + URL);
            Console.WriteLine(" ...");
            HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(URL);
            HttpWebResponse httpWebReponse = (HttpWebResponse)httpWebRequest.GetResponse();
            Stream stream = httpWebReponse.GetResponseStream();
            Image background = Image.FromStream(stream);
            Console.WriteLine("Download done!");
            return background;
        }

        static String getBackgroundPath()
        {
            String directory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "/NASA Wallpapers/";
            Directory.CreateDirectory(directory);
            String picDate = "20";
            picDate += date[0].ToString() + date[1].ToString() + "-" + date[2].ToString() + date[3].ToString() + "-" + date[4].ToString() + date[5].ToString();
            return Path.Combine(directory, picDate + picName + ".jpg");
        }

        static Boolean saveBackground(Image background)
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
            NativeMethods.SystemParametersInfo(SET_DESKTOP_BACKGROUND, 0, getBackgroundPath(), UPDATE_INI_FILE | SEND_WINDOWS_INI_CHANGE);
            Console.WriteLine("Wallpaper set and done!");
        }
    }
}