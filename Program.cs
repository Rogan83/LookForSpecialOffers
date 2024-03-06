using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections;
using System.Diagnostics;
using System.Reflection.Metadata;

namespace LookForSpecialOffers
{
    class Program
    {
        static void Main(string[] args) 
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless");              //öffnet die seiten im hintergrund
            using (IWebDriver driver = new ChromeDriver(options))
            {
                string pathMainPage = "https://www.penny.de";
                driver.Navigate().GoToUrl(pathMainPage);
                
                GoToOffersPage(driver, pathMainPage);

                ScrollToBottom(driver, 30, 10);         // Es könnte sein, dass die Zeit nicht ausreicht. Vllt sollte ich, falls auf ein Element nicht zugegriffen werden kann, diese Methode wiederholen


                string searchName = "//div[contains(@class, 'tabs__content-area')]";      //Suche nach dem Element, wo alle links von der Kopfzeile vorhanden sind
                var articleMainContainer = (HtmlNode)FindObject(driver, searchName, KindOfSearchElement.SelectSingleNode, 100, 10);  //Sucht solange nach diesen Element, bis es erschienen ist.
                HtmlNode mainContainer1;
                if (articleMainContainer != null )
                {
                    var mainSection1 = articleMainContainer.SelectNodes("./section[@class='tabs__content tabs__content--offers t-bg--wild-sand ']")[0];
                    var articleContainer1 = mainSection1.SelectNodes("./div[@class='js-category-section']")[0];
                    var articleContainer1Section1 = articleContainer1.SelectNodes("./section")[0];

                    var weekdayHeadline = articleContainer1Section1.Attributes["id"].Value;

                    var list = articleContainer1Section1.SelectSingleNode("./div[@class='l-container']//ul[@class='tile-list']");

                    var article1 = list.SelectNodes("./li")[3];
                    var count = list.SelectNodes("./li").Count();

                    //var articleName = ((HtmlNode)article1.SelectSingleNode("./div[@class='offer-tile__info-container']//h4[@class= 'tile__hdln offer-tile__headline']//a[@class= 'tile__link--cover']"));
                    var info = article1.SelectSingleNode("./article[@class= 'tile offer-tile']//div[@class='offer-tile__info-container']");
                    var articleName = ((HtmlNode)info.SelectSingleNode("./h4[@class= 'tile__hdln offer-tile__headline']//a[@class= 'tile__link--cover']")).InnerText;
                    var articlePricePerKg = ((HtmlNode)info.SelectSingleNode("./div[@class='offer-tile__unit-price ellipsis']")).InnerText;
                    articlePricePerKg = ExtractPrice(articlePricePerKg);

                    int iii = 9; //[@class= '']
                }



                driver.Quit();
            }

            static string ExtractPrice(string input)
            {
                // Teilen Sie den Eingabetext am "="-Zeichen
                string[] parts = input.Split('=');

                // Überprüfen, ob der Eingabetext das erwartete Format hat
                if (parts.Length == 2)
                {
                    // Extrahieren Sie den Teil nach dem "="-Zeichen und entfernen Sie unnötige Leerzeichen
                    string valuePart = parts[1].Trim();

                    // Ersetzen Sie den Punkt durch ein Komma
                    valuePart = valuePart.Replace('.', ',');


                    // Zeigen Sie den Wert an
                    Console.WriteLine(valuePart);

                    return valuePart;
                }
                else
                {
                    Console.WriteLine("Ungültiges Eingabeformat");
                    return string.Empty;
                }
            }

            static void GoToOffersPage(IWebDriver driver, string pathMainPage)
            {
                string searchName = "//div[contains(@class, 'site-header__wrapper')]";      //Suche nach dem Element, wo alle links von der Kopfzeile vorhanden sind
                var siteHeaderWrapperNode = (HtmlNode)FindObject(driver, searchName, KindOfSearchElement.SelectSingleNode, 100, 10);  //Sucht solange nach diesen Element, bis es erschienen ist.
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
                    Console.WriteLine("Der Node mit der Klasse 'site-header__wrapper' wurde nicht gefunden.");
                }
            }
        }
        /// <summary>
        /// Scroll stufenweise nach unten, damit die Seite komplett geladen wird.
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="delayPerStep"></param>
        /// <param name="steps"></param>
        static void ScrollToBottom(IWebDriver driver, int delayPerStep = 10, int steps = 10)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            long scrollHeight = (long)js.ExecuteScript("return document.body.scrollHeight;");

            long offset = scrollHeight / steps;
            long newPos = 0;

            for (int i = 0; i < steps; i++)
            {
                newPos += offset;

                // Scrolle bis zum Ende der Seite

                ((IJavaScriptExecutor)driver).ExecuteScript($"window.scrollTo(0, {newPos});");
                

                // Warte eine kurze Zeit, um die Seite zu laden
                System.Threading.Thread.Sleep(delayPerStep); // Wartezeit in Millisekunden anpassen
            }
        }

        private static object FindObject(IWebDriver driver, string name, KindOfSearchElement searchElement, int interval = 500, int maxSearchTimeInSeconds = 10)
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
    }

    class Product
    {
        public string Name { get; set; }
        public float Price { get; set;}

        Product(string name, float price)
        {
            Name = name;
            Price = price;
        }
    }
}








