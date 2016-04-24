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
                Console.WriteLine(" - 'r' or 'random'     Gives you a random image");
                Console.WriteLine(" - ''                  No input will give you today's image");
                Console.WriteLine(" - 'q' or 'quit'       Quits the program");

                String input = Console.ReadLine();

                if (input == "")
                {
                    preMillennial = false;
                    Console.WriteLine("Attempting to download today's picture");

                    Boolean success = setPicture(baseURL);
                    if (success && !tryAgain()) break;
                }
                else if (isValidDate(ref input))
                {
                    date = input;

                    Boolean success = setPicture(baseURL);
                    if(success && !tryAgain()) break;
                }
                else if (input == "r" || input == "random" || input == "R" || input == "Random" || input == "RANDOM")
                {
                    for (int i = 0; i < 5; i++)
                    {
                        Console.WriteLine("Generating a random date ...");
                        date = randomDate();

                        Boolean success = setPicture(baseURL);
                        if (success) break;
                    }
                    if (!tryAgain()) break;
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
        }

        public static Boolean tryAgain()
        {
            for (int i = 0; i < 5; i++)
            {
                Console.WriteLine("Try again? y/n");
                String again = Console.ReadLine();
                if (again == "y" || again == "yes" || again == "Y" || again == "Yes" || again == "YES") return true;
                if (again == "n" || again == "no" || again == "N" || again == "No" || again == "NO") return false;
                Console.Write("Invalid Input\t");
            }
            return false;
        }

        public static Boolean setPicture(String baseURL)
        {
            String htmlCode;
            String filePath;
            String pictureURL;
            WebClient client = new WebClient();
            //Console.WriteLine("Fetching HTML code from: " + baseURL + "ap" + date + ".html");
            try
            {
                htmlCode = client.DownloadString(baseURL + "ap" + date + ".html");
                //Console.WriteLine("HTMLcode = " + htmlCode);
                String pathExpr = @"\bimage\S+(.jpg|.gif|.png)\b";
                Regex reg = new Regex(pathExpr);
                Match match = reg.Match(htmlCode);
                if (!match.Success)
                {
                    Console.WriteLine("Unfortunately no image was found for this date.");
                    return false;
                }
                filePath = match.Value;
                pictureURL = baseURL + filePath;

                picName = findName(htmlCode);
                Console.WriteLine("\tTitle:" + picName);

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

        public static String randomDate()
        {
            String ranDate;
            int year, month, day;
            Random r = new Random();
            DateTime today = DateTime.Today;

            year = r.Next(1995, today.Year + 1); 
            if (year != today.Year && year != 1995) { // Not this year or 1995: let month and day be anything
                month = r.Next(1, 12+1);
                day = r.Next(1, DateTime.DaysInMonth(year, month)+1);
            }
            else if(year == today.Year){ // This year: only allow untill today
                month = r.Next(1, today.Month + 1);
                if (month != today.Month) {
                    day = r.Next(1, DateTime.DaysInMonth(year, month)+1);
                }
                else {
                    day = r.Next(1, today.Day + 1);
                }
            }
            else { // 1995: only allow from 16th of June
                month = r.Next(6, 13);
                if (month != 6)
                {
                    day = r.Next(1, DateTime.DaysInMonth(year, month)+1);
                }
                else {
                    day = r.Next(16, DateTime.DaysInMonth(year, month) + 1);
                }
            }
            preMillennial = year < 2000;
            DateTime foundDate = new DateTime(year,month, day);
            Console.WriteLine("Random date: " + foundDate.ToString("D", new CultureInfo("en-US")));
            return foundDate.ToString("yyMMdd");
        }

        public static Boolean isValidDate(ref String inputDate)
        {
            DateTime inputDateTime;
            DateTime today = DateTime.Today;
            DateTime oldest;
            DateTime.TryParseExact("160695", "ddMMyy", new CultureInfo("en-US"), DateTimeStyles.None, out oldest);

            if (!DateTime.TryParseExact(inputDate, "ddMMyy", new CultureInfo("en-US"), DateTimeStyles.None, out inputDateTime)) return false;
            if (inputDateTime < oldest || inputDateTime > today) return false;
            Console.WriteLine("Attempting to download picture of " + inputDateTime.ToString("D", new CultureInfo("en-US")));
            preMillennial = inputDateTime.Year < 2000;
            inputDate = inputDateTime.ToString("yyMMdd");
            return true;
        }

        public static String findName(String htmlCode)
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
                if (!badGuys.Contains(i))   name += title[i].ToString();
            }
            return name;
        }

        public static Image downloadBackground(String URL)
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

        public static String getBackgroundPath()
        {
            String directory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + "/NASA Wallpapers/";
            Directory.CreateDirectory(directory);
            String picDate = preMillennial ? "19" : "20";
            picDate += date[0].ToString() + date[1].ToString() + "-" + date[2].ToString() + date[3].ToString() + "-" + date[4].ToString() + date[5].ToString();
            return Path.Combine(directory, picDate + picName + ".jpg");
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