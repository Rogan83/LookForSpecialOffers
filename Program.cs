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
using static LookForSpecialOffers.WebScraperHelper;
//Bugs:

//todo:
// - Den User eventuell darauf hinweisen, dass die Excel Tabelle geschlossen werden muss, während das Programm läuft, sonst kann sie nicht mit neuen Daten überschrieben werden.
// - Andere Discounter und Supermärkte hinzufügen (Bis jetzt wurde Penny und LIDL hinzugefügt).  
// - eine grafische Oberfläche mit Einstellmöglichkeiten implementieren (mit .NET Maui). Darüber können die vers.
//   Discounter ausgewählt werden, welche bei der Suche berücksichtigt werden sollen, nach welchen Produkten gesucht werden
//   sollen, welchen Preis sie haben dürfen usw. 

namespace LookForSpecialOffers
{
    static class Program
    {
        // Diese Daten sollen später durch die Eingabe gesetzt werden (entweder durch eine grafischen Oberfläche oder einfach durch eine Datei, in der man diese
        // Daten direkt einträgt).
        #region Testdaten
        static internal List<ProduktFavorite> InterestingProducts { get; set; } = new()
        {
            new ProduktFavorite("Speisequark", 2.60),
            new ProduktFavorite("Thunfisch", 5.08),
            new ProduktFavorite("Tomate", 2.00),
            new ProduktFavorite("Orange", 0.99),
            new ProduktFavorite("Buttermilch", 0.99),
            new ProduktFavorite("Äpfel", 1.99),
            new ProduktFavorite("Hackfleisch", 5.99)
        };
        internal static string ExcelPath { get; set; } = "Angebote.xlsx";

        internal static string EMail { get; set; } = "d.rothweiler@yahoo.de";
        //static string EMail { get; set; } = "hadess90@web.de";
        //static string EMail { get; set; } = "tubadogan.85@googlemail.com";

        internal static string ZipCode { get; set; } = "01239";
        #endregion

        static internal Dictionary<Discounter, List<Product>> AllProducts = new Dictionary<Discounter, List<Product>>();

        internal static bool IsNewOffersAvailable { get; set; } = false;                  // Sind neue Angebote vorhanden? Falls ja, dann soll eine E-Mail verschickt werden

        static void Main(string[] args) 
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless");              //öffnet die Seiten im Hintergrund
            using (IWebDriver driver = new ChromeDriver(options))
            {
                //driver.Manage().Window.Maximize();
                //driver.Manage().Window.Minimize();
                string periodheadline = ExtractHeadlineFromExcel(ExcelPath);

                // Extrahiert die Daten wie Artikelnamen, Preis etc. von bestimmten Webseiten von Discountern und anderen Supermärkten.
                Penny.ExtractOffers(driver, periodheadline);
                //Lidl.ExtractOffers(driver, periodheadline);

                InformPerEMail(IsNewOffersAvailable, AllProducts);

                excelPackage.Dispose();

                driver.Quit();
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
