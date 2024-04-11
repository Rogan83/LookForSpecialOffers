using HtmlAgilityPack;
using LookForSpecialOffers.Enums;
using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.DevTools;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Net;
using System.Net.Mail;
using System.Reflection.Metadata;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Newtonsoft.Json;
using static LookForSpecialOffers.WebScraperHelper;
using Newtonsoft.Json.Linq;
using LookForSpecialOffers.Models;
using System.Collections.ObjectModel;
//Bugs:

//todo:
// - Den User eventuell darauf hinweisen, dass die Excel Tabelle geschlossen werden muss, während das Programm läuft, sonst kann sie nicht mit neuen Daten überschrieben werden.
// - Andere Discounter und Supermärkte hinzufügen (Bis jetzt wurde Penny und LIDL hinzugefügt).  
// - Bei LIDL muss noch überprüft werden, ob die aktuelle Angebote von den bereits gespeicherten Angeboten in der Excel Tabelle abweicht, damit 
//   per E-Mail benachrichtigt werden kann, ob neue Angebote vorhanden sind. (Bei Penny wurde dies bereits umgesetzt). 

namespace LookForSpecialOffers
{
    static class Program
    {
        //static string projectPath = AppDomain.CurrentDomain.BaseDirectory;
        static string path = @"H:\Test";

        //static string jsonFilePath = Path.Combine(projectPath, "settings.json");
        static string jsonFilePath = Path.Combine(path, "settings.json");

        static internal List<FavoriteProduct> FavoriteProducts { get; set; } = new();
        static Dictionary<string, bool?> Markets { get; set; } = new Dictionary<string, bool?>();

        internal static string ExcelFilePath { get; set; } = "Angebote.xlsx";

        internal static string Email { get; set; } = string.Empty;
        //internal static string Email { get; set; } = "d.rothweiler@yahoo.de";
        //static string EMail { get; set; } = "hadess90@web.de";
        //static string EMail { get; set; } = "tubadogan.85@googlemail.com";

        internal static string ZipCode { get; set; } = string.Empty;

        static internal Dictionary<MarketEnum, List<Product>> AllProducts = new Dictionary<MarketEnum, List<Product>>();

        internal static bool IsNewOffersAvailable { get; set; } = false;                  // Sind neue Angebote vorhanden? Falls ja, dann soll eine E-Mail verschickt werden

        static void Main(string[] args) 
        {
            // Daten laden
            if (File.Exists(jsonFilePath))
            {
                LoadData();
            }
            else
            {
                FavoriteProducts = new()
                {
                    new FavoriteProduct("Speisequark", 2.60m, 1.30m),
                    new FavoriteProduct("Thunfisch", 5.08m, 1.00m),
                    new FavoriteProduct("Tomate", 2.00m, 1.00m),
                    new FavoriteProduct("Orange", 2.00m, 0.99m),
                    new FavoriteProduct("Buttermilch", 0.99m, 0.49m),
                    new FavoriteProduct("Äpfel", 1.99m, 0.49m),
                    new FavoriteProduct("Hackfleisch", 5.99m, 2.49m)
                };

                Markets = new Dictionary<string, bool?>
                {
                    { "Penny", true },
                    { "Lidl", false },
                    { "Aldi Nord", false },
                    { "Netto", false },
                    { "Kaufland", false }
                };

                ExcelFilePath = "Angebote.xlsx";
            }

            //Test
            //List<Product> p = LoadFromExcel(ExcelFilePath, Discounter.Penny);

            ChromeOptions options = new ChromeOptions();
            //FirefoxOptions options = new();
            //options.AddArgument("--headless");              //öffnet die Seiten im Hintergrund
            using (IWebDriver driver = new ChromeDriver(options))
            //using (IWebDriver driver = new FirefoxDriver(options))
            {
                driver.Manage().Window.Maximize();
                //driver.Manage().Window.Minimize();

                // Extrahiert die Daten wie Artikelnamen, Preis etc. von bestimmten Webseiten von Discountern und anderen Supermärkten.
                if (Markets["Penny"] != null)
                {
                    if (Markets["Penny"].Value == true)
                    {
                        string periodheadline = ExtractHeadlineFromExcel(ExcelFilePath, MarketEnum.Penny);
                        Penny.ExtractOffers(driver, periodheadline);
                    }
                }

                if (Markets["Lidl"] != null)
                {
                    if (Markets["Lidl"].Value == true)
                    {
                        string periodheadline = ExtractHeadlineFromExcel(ExcelFilePath, MarketEnum.Lidl);
                        Lidl.ExtractOffers(driver, periodheadline);
                    }
                }

                if (Markets["Aldi Nord"] != null)
                {
                    if (Markets["Aldi Nord"].Value == true)
                    {
                        string periodheadline = ExtractHeadlineFromExcel(ExcelFilePath, MarketEnum.AldiNord);
                        AldiNord.ExtractOffers(driver, periodheadline);
                    }
                }

                InformPerEMail(IsNewOffersAvailable, AllProducts);
                excelPackage.Dispose();
                driver.Quit();
            }
        }

        static void LoadData()
        {
            string jsonStringFromFile = File.ReadAllText(jsonFilePath);

            JObject data = JObject.Parse(jsonStringFromFile);

            JArray loadedProducts = (JArray)data["FavoriteProducts"];
            FavoriteProducts.Clear();
            foreach (JObject loadedProduct in loadedProducts)
            {
                string name = (string)loadedProduct["Name"];
                decimal priceCapPerKg = (decimal)loadedProduct["PriceCapPerKg"];
                decimal priceCapPerProduct = (decimal)loadedProduct["PriceCapPerProduct"];

                FavoriteProducts.Add(new FavoriteProduct(name, priceCapPerKg, priceCapPerProduct));
            }

            JArray loadedMarkets = (JArray)data["Markets"];
            foreach (JObject loadedMarket in loadedMarkets)
            {
                string? name = (string?)loadedMarket["Name"];
                bool? isSelected = (bool?)loadedMarket["IsSelected"];

                if (name != null && isSelected != null)
                {
                    Markets[name] = isSelected;
                }
            }

            string loadedEmail, loadedPath, loadedZipCode;
            if (data["Email"] != null)
            {
                var temp = (string)data["Email"];
                if (temp != string.Empty)
                    Email = temp;
            }

            if (data["Path"] != null)
            {
                var temp = (string)data["Path"];
                if (temp != string.Empty)
                    ExcelFilePath = temp;
            }

            if (data["ZipCode"] != null)
            {
                ZipCode = (string)data["ZipCode"];
            }
        }

        #region Nicht verwendete Methoden
        static void GoToOffersPage(IWebDriver driver, string pathMainPage)
        {
            string searchName = "//div[contains(@class, 'site-header__wrapper')]";      //Suche nach dem Element, wo alle links von der Kopfzeile vorhanden sind
            HtmlNode? siteHeaderWrapperNode = null;
            try
            {
                siteHeaderWrapperNode = (HtmlNode?)Searching(driver, searchName, KindOfSearchElement.SelectSingleNode);  //Sucht solange nach diesen Element, bis es erschienen ist.
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fehler: " + ex);
            }
            if (siteHeaderWrapperNode != null)
            {
                // XPath-Ausdruck, um das erste a-Element im ersten li-Element mit der angegebenen Klasse zu finden
                string xpathExpression = ".//div[@class='show-for-large']//nav[@class='site-header__nav']//div[@class='main-nav__container']//ul//li[@class='main-nav__item has-submenu'][1]//a[@href]";
                xpathExpression = ".//div[@class='site-header__container']//div[@class='show-for-large']//nav[@class='site-header__nav']//div[@class='main-nav__container']//ul//li[@class='main-nav__item has-submenu'][1]//a[@href]";
                // Das erste passende Element finden, beginnend von siteHeaderWrapperNode
                HtmlNode linkNode = siteHeaderWrapperNode.SelectSingleNode(xpathExpression);

                // Überprüfen, ob ein Element gefunden wurde, und den Wert des href-Attributs abrufen
                if (linkNode != null)
                {
                    string hrefValue = linkNode.Attributes["href"].Value;
                    string pathOffers = String.Concat(pathMainPage, hrefValue);
                    driver.Navigate().GoToUrl(pathOffers);
                    Debug.WriteLine("Der href-Wert des ersten a-Elements ist: " + hrefValue);
                }
                else
                {
                    Debug.WriteLine("Das gewünschte Element wurde nicht gefunden.");
                }
            }
            else
            {
                Debug.WriteLine("Der Node mit der Klasse 'site-header__wrapper' wurde nicht gefunden.");
            }
        }

        /// <summary>
        /// Versucht, ein bestimmtes Element zu finden und versucht es in gewissen Zeitabständen erneut, falls dieses Element nicht gefunden wird.
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="name"></param>
        /// <param name="searchElement"></param>
        /// <param name="interval"></param>
        /// <param name="maxSearchTimeInSeconds"></param>
        /// <returns></returns>
        static object FindObject(IWebDriver driver, string name, KindOfSearchElement searchElement, int interval = 500, int maxSearchTimeInSeconds = 10)
        {
            int maxRepeats = (int)(maxSearchTimeInSeconds / (interval/1000.0f));

            HtmlDocument doc = new HtmlDocument();
            int repeat = 0;
            object element = null;
            while (repeat < maxRepeats)
            {
                if (doc != null && driver != null)
                    doc.LoadHtml(driver.PageSource);

                if (searchElement == KindOfSearchElement.SelectSingleNode && doc != null)
                {
                    element = doc.DocumentNode.SelectSingleNode(name);
                }
                else if (searchElement == KindOfSearchElement.SelectNodes && doc != null)
                {
                    element = doc.DocumentNode.SelectNodes(name);
                    //element = doc.DocumentNode.SelectNodes("//a[contains(@class, 'mat-list-item') and contains(@class, 'mat-focus-indicator') and contains(@class, 'mat-ripple') and contains(@class, 'search-result') and contains(@class, 'mat-list-item-with-avatar') and contains(@class, 'ng-star-inserted')]");
                }
                else if (searchElement == KindOfSearchElement.FindElementByCssSelector && driver != null)
                {
                    try
                    {
                        element = driver.FindElement(By.CssSelector(name));
                    }
                    catch { }
                }
                else if (searchElement == KindOfSearchElement.FindElementsByCssSelector && driver != null)
                {
                    try
                    {
                        element = driver.FindElements(By.CssSelector(name));
                    }
                    catch { }
                }

                if (element != null)
                {
                    if (searchElement == KindOfSearchElement.FindElementsByCssSelector || searchElement == KindOfSearchElement.SelectNodes)
                    {
                        ICollection collection;
                        try
                        {
                            collection = (ICollection)element;
                            if (collection != null && collection.Count > 0)
                            {
                                return element;
                            }
                        }
                        catch
                        {

                        }
                    }
                    else
                    {
                        return element;
                    }
                }

                repeat++;
                Thread.Sleep(interval);
            }
            return null;
        }
        #endregion
    }
}
