using HtmlAgilityPack;
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
//Bugs:
// - Ermittelt nicht den richtigen Wert für die Höhe von der Webseite
// - Die Standorte werden nicht berücksichtigt, was bedeutet, dass keine lokalen Angebote gespeichert werden. Möglicher Ansatz: Wenn die webseite geladen wird und ganz nach unten gescrollt
//   wird, dann erscheint zum einen ein button, den man drücken muss, um alle cookies zu bestätigen und zum andreen ein Fenster zur Standorteingabe. Dies muss man wohl selbst ausfüllen.

//todo:
// - Den User eventuell darauf hinweisen, dass die Excel Tabelle geschlossen werden muss, während das Programm läuft, sonst kann sie nicht mit neuen Daten überschrieben werden.
// - Andere Discounter hinzufügen.
// - Benachrichtigung per E-Mail implementieren, wenn bestimmte Angebote einen Wert unterschritten haben.   (erledigt)
// - eine grafische Oberfläche mit Einstellmöglichkeiten implementieren (mit .NET Maui). Darüber können die vers. Discounter ausgewählt werden, welche bei der Suche berücksichtigt werden sollen,
//   nach welchen Produkten gesucht werden sollen, welchen Preis sie haben dürfen usw. 
namespace LookForSpecialOffers
{
    class Program
    {
        // Diese Daten sollen später durch die Eingabe gesetzt werden (entweder durch eine grafischen Oberfläche oder einfach durch eine Datei, in der man diese
        // Daten direkt einträgt).
        #region Testdaten
        static public List<ProduktFavorite> interestingProducts = new() 
        { 
            new ProduktFavorite("Quark", 2.60), 
            new ProduktFavorite("Thunfisch", 5.08), 
            new ProduktFavorite("Tomate", 9.50),
            new ProduktFavorite("Orange", 0.99),
            new ProduktFavorite("Buttermilch", 0.99)
        };

        static string ExcelPath { get; set; } = "Angebote.xlsx";

        static string EMail { get; set; } = "d.rothweiler@yahoo.de";
        //static string EMail { get; set; } = "hadess90@web.de";
        //static string EMail { get; set; } = "tubadogan.85@googlemail.com";

        static string zipCode = "68789";
        #endregion

        static void Main(string[] args) 
        {
            ChromeOptions options = new ChromeOptions();
            //options.AddArgument("--headless");              //öffnet die seiten im hintergrund

            //funktioniert leider nicht
            //string profilePath = @"C:\Users\droth\AppData\Local\Google\Chrome\User Data";
            //options.AddArguments($"--user-data-dir={profilePath}");
            //string profilePathFirefox = @"C:\Users\droth\AppData\Local\Mozilla\Firefox\Profiles";
            //profilePathFirefox = @"C:\Users\droth\AppData\Roaming\Mozilla\Firefox\Profiles\h1iwxt69";
            //optionsFirefox.AddArguments($"--user-data-dir={profilePathFirefox}");
            //optionsFirefox.Profile = new FirefoxProfile("h1iwxt69.default");

            using (IWebDriver driver = new ChromeDriver(options))
            {

                string periodheadline = ExtractHeadlineFromExcel(ExcelPath);
                ExtractOffersFromPenny(driver, periodheadline);

                driver.Quit();
            }

            

            // Extrahiert alle Sonderangebote vom Penny und speichert diese formatiert in eine Excel Tabelle ab.
            static void ExtractOffersFromPenny(IWebDriver driver, string oldPeriodHeadline)
            {
                string pathMainPage = "https://www.penny.de/angebote";
                bool isNewOffersAvailable = false;                  // Sind neue Angebote vom Penny vorhanden? Falls ja, dann soll eine E-Mail verschickt werden

                driver.Navigate().GoToUrl(pathMainPage);

                //Thread.Sleep(1000);


                //GoToOffersPage(driver, pathMainPage);      //Scheint jetzt richtig zu gehen
                EnterZipCode(driver, zipCode);

                ScrollToBottom(driver, 50, 20, 1000);         // Es scheint so, dass es wichtig ist, dass man das herunterscrollen in vielen Steps einteilen wichtig ist


                string searchName = "//div[contains(@class, 'tabs__content-area')]";      //Suche nach dem Element, wo alle links von der Kopfzeile vorhanden sind



                HtmlNode mainContainer = (HtmlNode)Find(driver,searchName, KindOfSearchElement.SelectSingleNode);   //Sucht solange nach diesen Element, bis es erschienen ist.

                //var mainContainer = (HtmlNode)FindObject(driver, searchName, KindOfSearchElement.SelectSingleNode, 200, 10);  //Sucht solange nach diesen Element, bis es erschienen ist.
                List<Product> products = new();

                if (mainContainer != null)     //Der maincontainer enthält alles relevantes
                {
                    var mainSection = mainContainer.SelectSingleNode("./section[@class='tabs__content tabs__content--offers t-bg--wild-sand ']");
                    var articleDivContainers = mainSection.SelectNodes("./div[@class='js-category-section']");

                    foreach (var articleDivContainer in articleDivContainers)
                    {
                        var offerStartDate = articleDivContainer.Attributes["id"].Value.Replace('-', ' ');  //Ab wann gilt dieses Angebot
                        var articleSectionContainers = articleDivContainer.SelectNodes("./section");

                        foreach (var articleSectionContainer in articleSectionContainers)
                        {
                            var weekdayHeadline = articleSectionContainer.Attributes["id"].Value;

                            var list = articleSectionContainer.SelectSingleNode("./div[@class='l-container']//ul[@class='tile-list']");
                            var items = list.SelectNodes("./li");
                            //var items = list.SelectNodes("./li[contains(@class, 'tile-list')]");
                            foreach (var item in items)
                            {
                                var badgeContainer = item.SelectSingleNode("./article//div[@class = 'badge--split t-bg--blue-petrol t-color--white']");
                                var badge = string.Empty;       // Die Plakette, die zusätzliche Infos angibt
                                if (badgeContainer != null)
                                {
                                    badge = ((HtmlNode)badgeContainer.SelectSingleNode("./span")).InnerHtml;
                                }

                                var info = item.SelectSingleNode("./article//div[@class='offer-tile__info-container']");
                                if (info == null) { continue; }         // das erste item hat keinen Artikel mit der Klasse. Deswegen muss dieser übersprungen werden

                                var articleName = WebUtility.HtmlDecode(((HtmlNode)info.SelectSingleNode("./h4[@class= 'tile__hdln offer-tile__headline']//a[@class= 'tile__link--cover']")).InnerText);

                                var articlePricePerKg = ((HtmlNode)info.SelectSingleNode("./div[@class='offer-tile__unit-price ellipsis']")).InnerText;
                                string description = articlePricePerKg.Split('(')[0].Trim();        //extrahiert die Beschreibung 
                                double articlePricePerKg1 = 0;
                                double articlePricePerKg2 = 0;

                                if (articlePricePerKg != null)
                                {
                                    var price1 = ExtractPrices(articlePricePerKg)[0];
                                    if (price1 != null)
                                        articlePricePerKg1 = Math.Round(double.Parse(price1, CultureInfo.InvariantCulture), 2);
                                    var price2 = ExtractPrices(articlePricePerKg)[1];
                                    if (price2 != null)
                                    {
                                        articlePricePerKg2 = Math.Round(double.Parse(price2, CultureInfo.InvariantCulture), 2);
                                    }
                                }

                                var priceContainer = item.SelectSingleNode("./article" +
                                    "//div[contains(@class, 'bubble offer-tile')]" +
                                    "//div");

                                string oldPriceText = "", newPriceText = "";
                                double oldPrice = 0, newPrice = 0;

                                var price = priceContainer.SelectSingleNode("./div//span[@class='value']");
                                if (price != null)
                                {
                                    oldPriceText = price.InnerText.Replace('–', ' ');
                                    oldPrice = double.Parse(oldPriceText, CultureInfo.InvariantCulture);
                                    oldPrice = Math.Round(oldPrice, 2);
                                }

                                price = priceContainer.SelectSingleNode("./span[@class='ellipsis bubble__price']");
                                if (price != null)
                                {
                                    newPriceText = (priceContainer.SelectSingleNode("./span[@class='ellipsis bubble__price']")).InnerText.
                                        Replace('–', ' ').Replace('*', ' ');
                                    if (!double.TryParse(newPriceText, CultureInfo.InvariantCulture, out newPrice))
                                        Debug.WriteLine($"folgende Zeichenkette konnte nicht umgewandelt werden: {newPriceText}");
                                    newPrice = Math.Round(newPrice, 2);
                                }

                                products.Add(new Product(articleName, description, oldPrice, newPrice, articlePricePerKg1, articlePricePerKg2, badge, offerStartDate));
                            }
                        }
                    }

                    // Der Zeitraum, von wann bis wann die Angebote gelten
                    var period = ((HtmlNode)mainSection.SelectSingleNode("./div[@class = 'category-menu']" +
                        "//div[@class = 'category-menu__header-wrapper']" +
                        "//div[@class = 'category-menu__header l-container']" +
                        "//div[@class = 'category-menu__header-container']" +
                        "//div//div//div")).Attributes["data-startend"].Value;



                    // Wenn keine Datei vorhanden ist, dann kann auch nicht verglichen werden, ob das Datum von den Angeboten,
                    // die in der Excel Tabelle stehen und das Datum von den aktuellen Angeboten übereinstimmen. In diesen Fall
                    // muss davon ausgegangen werden, dass noch nicht über die neuen Angebote informiert wurde.
                    if (oldPeriodHeadline == string.Empty || !oldPeriodHeadline.Contains(period))
                    {
                        isNewOffersAvailable = true;
                    }

                    InFormPerEMail(isNewOffersAvailable, products);


                    SaveToExcel(products, period, ExcelPath);
                }

                static string[] ExtractPrices(string input)
                {
                    string[] prices = new string[2];
                    if (!input.Contains("("))           // Der Preis (falls vorhanden) ist immer in der Klammer enthalten. Wenn kein Preis vorhanden ist, dann interessiert diese Info nicht und gibt einen leeren String zurück.
                        return prices;

                    // Teilen Sie den Eingabetext am "="-Zeichen

                    string[] parts = input.Split('=');

                    // Überprüfen, ob der Eingabetext das erwartete Format hat
                    if (parts.Length == 2)
                    {
                        // Extrahieren Sie den Teil nach dem "="-Zeichen und entfernen Sie unnötige Leerzeichen
                        string valuePart = parts[1].Trim();

                        if (valuePart.Contains('/'))
                        {
                            prices = valuePart.Split('/');
                            prices[1] = prices[1].Replace(')', ' ');
                        }
                        else
                        {
                            // Die schließende Klammer wird entfernt
                            prices[0] = valuePart.Replace(')', ' ');
                        }

                        return prices;
                    }
                    else
                    {
                        Debug.WriteLine("ungültiges input format");
                        return prices;          // in diesen Fall ist prices jeweils leer
                    }
                }

                // Informiert jedesmal, wenn neue Angebote verfügbar sind, per E-Mail, wenn diese einen bestimmten festgelegten Preis unterstreiten.
                static void InFormPerEMail(bool isNewOffersAvailable, List<Product> products)
                {
                    if (isNewOffersAvailable)
                    {
                        int interestingOfferCount = 0;
                        string offers = string.Empty;
                        // Als nächstes soll untersucht werden, ob von den interessanten Angeboten der Preis auch niedrig genug ist.
                        foreach (var product in products)
                        {
                            foreach (var interestingProduct in interestingProducts)
                            {
                                if (product.Name.ToLower().Trim().Contains(interestingProduct.Name.ToLower().Trim()))
                                {
                                    // falls bei beiden Kg bzw. Liter Preise nichts drin steht, dann könnte das bedeuten, dass entweder die Menge schon ein Kilo entspricht oder dass es einzel Preise sind
                                    if (product.PricePerKgOrLiter1 == 0 && product.PricePerKgOrLiter2 == 0)
                                    {
                                        if (product.NewPrice <= interestingProduct.PriceCap)
                                        {
                                            offers += $" {product.Name} für nur {product.NewPrice} €.\n";
                                            interestingOfferCount++;
                                        }
                                    }
                                    else if ((product.PricePerKgOrLiter1 <= interestingProduct.PriceCap && product.PricePerKgOrLiter1 != 0) ||
                                        (product.PricePerKgOrLiter2 <= interestingProduct.PriceCap && product.PricePerKgOrLiter2 != 0))
                                    {
                                        offers += $" {product.Name} für nur {product.NewPrice} €.\n";
                                        interestingOfferCount++;
                                    }
                                }
                            }
                        }
                        string body = string.Empty;
                        string subject = string.Empty;
                        if (interestingOfferCount > 1)
                        {
                            subject = "Interessantes Angebote gefunden!";
                            body = $"Gute Nachricht! Folgende Angebote, welche deine preislichen Vorstellungen entspricht, wurden gefunden: \n\n{offers}";
                        }
                        else if (interestingOfferCount == 1)
                        {
                            subject = "Es wurde ein interessantes Angebot gefunden!";
                            body = $"Gute Nachricht! Folgendes Angebot, welches deine preisliche Vorstellung entspricht, wurde gefunden: \n\n{offers}";
                        }

                        if (interestingOfferCount > 0)
                        {
                            body += "\nHier ist der Link: https://www.penny.de/angebote \nLass es dir schmecken!";
                            SendEMail(EMail, subject, body);
                        }
                    }
                }

                static void EnterZipCode(IWebDriver driver, string zipCode)
                {
                    // Der Button befindet sich innerhalb der Shadow Root. Dieses kapselt die Elemente, die darin enthalten sind,
                    // d.h. dass die Elemente, die sich darin befinden, von außen geschützt sind. Deswegen muss erst auf den
                    // Shadow Root zugegriffen werden und von dort aus kann nach den Elementen darin gesucht werden.
                    var parent = driver.FindElement(By.XPath("//*[@id='usercentrics-root']"));
                    ShadowRoot shadowRoot = (ShadowRoot)((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].shadowRoot", parent);
                    // Zuerst muss der Cookie Button geklickt werden, weil dieser im Vordergrund ist und das anklicken des Button verhindert, mit dem alle Angebote eingesehen werden können. 
                    Thread.Sleep(1000);
                    var CookieAcceptBtn = shadowRoot.FindElement(By.CssSelector(".sc-dcJsrY.iWikWl"));
                    //var CookieAcceptBtn = (IWebElement)Find(driver, ".sc-dcJsrY.iWikWl", KindOfSearchElement.FindElementByCssSelector);

                    CookieAcceptBtn.Click();                         //bestätigt die Cookies

                    var ShowSearchForMarketBtn = driver.FindElement(By.ClassName("market-tile__btn-not-selected"));

                    Actions actions = new Actions(driver);
                    actions.MoveToElement(ShowSearchForMarketBtn).Perform();
                    // Wenn auf diesen Button geklickt wird, wird ein Eingabefeld angezeigt, wo die PLZ eingegeben werden kann
                    ShowSearchForMarketBtn.Click();



                    // Instanziieren der Actions-Klasse
                    actions = new Actions(driver);

                    // Schrittweise Eingabe des Textes
                    foreach (char c in zipCode)
                    {
                        // Drücken der Taste für den aktuellen Buchstaben
                        actions.SendKeys(c.ToString()).Build().Perform();
                    }


                    IWebElement wrapper = driver.FindElement(By.XPath("//*[@class='market-modal__wrapper']"));
                    //Thread.Sleep(1000);
                    IWebElement chooseMarketBtn = null;
                    chooseMarketBtn = SearchForChooseMarketBtn(driver, wrapper, chooseMarketBtn, 100, 5000);

                    if (chooseMarketBtn != null)
                        ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", chooseMarketBtn);

                    // Sucht solange nach diesen Button, bis er erscheint.
                    static IWebElement SearchForChooseMarketBtn(IWebDriver driver, IWebElement wrapper, IWebElement chooseMarketBtn, int delayStep, int maxDelayTotal)
                    {
                        int maxCount = maxDelayTotal / delayStep;
                        int count = 0;
                        while (count < maxCount)
                        {
                            try
                            {
                                chooseMarketBtn = wrapper.FindElement(By.XPath(".//main//div[@class='market-modal__results']//ul//li[1]//article//div//div//a"));
                                break;
                            }
                            catch
                            {
                                chooseMarketBtn = null;
                                Thread.Sleep(delayStep);
                            }
                            count++;
                        }

                        //var href = aTag.GetAttribute("href");
                        return chooseMarketBtn;
                    }
                }
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
                Debug.WriteLine("Der Node mit der Klasse 'site-header__wrapper' wurde nicht gefunden.");
            }
        }

        /// <summary>
        /// Scroll stufenweise nach unten, damit die Seite komplett geladen wird.
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="delayPerStep"></param>
        /// <param name="steps"></param>

        static void ScrollToBottom(IWebDriver driver, int delayPerStep = 200, int steps = 10, int delayDetermineScrollHeigth = 500)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            
            long oldScrollHeight = (long)js.ExecuteScript("return document.body.scrollHeight;");
            long newScrollHeight = 0;
            //Es dauert eine Weile, bis die Scrollheight ermittelt wird. Deswegen wird die schleife so lange wiederholt, bis sind die Scrollheight nicht mehr verändert, was bedeutet, dass diese den entgültigen wert ermittelt haben muss
            while (true)
            {
                Thread.Sleep(delayDetermineScrollHeigth);
                newScrollHeight = (long)js.ExecuteScript("return document.body.scrollHeight;");     

                if (newScrollHeight == oldScrollHeight)
                    break;
                else
                {
                    oldScrollHeight = newScrollHeight;
                }
            }

            long offset = oldScrollHeight / steps;
            long newPos = 0;

            for (int i = 0; i < steps; i++)
            {
                newPos += offset;

                // Scrolle bis zum Ende der Seite

                ((IJavaScriptExecutor)driver).ExecuteScript($"window.scrollTo(0, {newPos});");
                

                // Warte eine kurze Zeit, um die Seite zu laden
                Thread.Sleep(delayPerStep); // Wartezeit in Millisekunden anpassen
            }
        }
        /// <summary>
        /// Überprüft, ob der Begriff in dem anderen Begriff vorkommt.
        /// </summary>
        /// <param name="substring"></param>
        /// <param name="product"></param>
        /// <returns></returns>
        static bool IsContains(string substring, string product)
        {
            substring = substring.Trim().ToLower();
            product = product.Trim().ToLower();

            return product.Contains(substring);
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
        /// <summary>
        /// Macht eig. das selbe wie die FindObject Methode, bloß eleganter.
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="searchName"></param>
        /// <param name="searchElement"></param>
        /// <param name="pollingIntervalTime"></param>
        /// <param name="maxWaitTime"></param>
        /// <returns></returns>
        static object Find(IWebDriver driver, string searchName, KindOfSearchElement searchElement, int pollingIntervalTime = 500, int maxWaitTime = 10)
        {
            object element = null;
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTime))
            {
                PollingInterval = TimeSpan.FromMilliseconds(pollingIntervalTime),
            };
            wait.IgnoreExceptionTypes(typeof(ElementNotInteractableException));


            var main = wait.Until(d =>
            {
                switch (searchElement)
                {
                    case KindOfSearchElement.SelectSingleNode:
                        try
                        {
                            HtmlDocument doc = new HtmlDocument();
                            doc.LoadHtml(driver.PageSource);
                            if (doc != null)
                                element = doc.DocumentNode.SelectSingleNode(searchName);

                            return element != null;
                        }
                        catch (NoSuchElementException)
                        {
                            return false;
                        }

                    case KindOfSearchElement.SelectNodes:
                        try
                        {
                            HtmlDocument doc = new HtmlDocument();
                            doc.LoadHtml(driver.PageSource);
                            if (doc != null)
                                element = doc.DocumentNode.SelectNodes(searchName);

                            return element != null;
                        }
                        catch (NoSuchElementException)
                        {
                            return false;
                        }

                    case KindOfSearchElement.FindElementByCssSelector:
                        try
                        {
                            element = driver.FindElement(By.CssSelector(searchName));

                            return element != null;
                        }
                        catch (NoSuchElementException)
                        {
                            return false;
                        }

                    case KindOfSearchElement.FindElementsByCssSelector:
                        try
                        {
                            element = driver.FindElements(By.CssSelector(searchName));

                            return element != null;
                        }
                        catch (NoSuchElementException)
                        {
                            return false;
                        }
                    default:
                        return false;
                }

                
            });

            return element;
        }

        static void SaveToExcel(List<Product> products, string period, string path)
        {
            if (products == null)
            {
                Debug.WriteLine("Keine Daten zum speichern vorhanden.");
                return;
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string excelFilePath = path;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Penny");

                int columnCount = 8;                                  // Anzahl der Spalten

                var cellHeadline = worksheet.Cells[1, 1];
                worksheet.Cells["A1:H2"].Merge = true;                  // Bereich verbinden

                cellHeadline.Value = $"Die Angebote vom Penny vom {period}";

                // Überschrift formatieren
                worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells["A1"].Style.Font.Size = 14;
                cellHeadline.Style.Font.Color.SetColor(Color.Wheat);

                HighLightCells(1, 1, 2, columnCount, Color.Red, worksheet);

                worksheet.Cells[3, 1].Value = "Name";
                worksheet.Cells[3, 2].Value = "Bezeichnung";
                worksheet.Cells[3, 3].Value = "Vorheriger Preis";
                worksheet.Cells[3, 4].Value = "Neuer Preis";
                worksheet.Cells[3, 5].Value = "Preis Pro Kg oder Liter (erstes Angebot)";
                worksheet.Cells[3, 6].Value = "Preis Pro Kg oder Liter (zweites Angebot)";
                worksheet.Cells[3, 7].Value = "Info";
                worksheet.Cells[3, 8].Value = "Beginn";

                // Beschriftung formatieren
                HighLightCells(3, 1, 3, columnCount, Color.Gray, worksheet);

                int offsetRow = 4;

                string euroFormat = "#,##0.00 €";           // Währungsformat

                for (int i = 0; i < products.Count; i++)
                {
                    worksheet.Cells[i + offsetRow, 1].Value = products[i].Name;
                    worksheet.Cells[i + offsetRow, 2].Value = products[i].Description;
                    if (products[i].OldPrice != 0)
                    {
                        worksheet.Cells[i + offsetRow, 3].Value = products[i].OldPrice;
                        worksheet.Cells[i + offsetRow, 3].Style.Numberformat.Format = euroFormat; ; 
                    }
                    if (products[i].NewPrice != 0)
                    { 
                        worksheet.Cells[i + offsetRow, 4].Value = products[i].NewPrice;
                        worksheet.Cells[i + offsetRow, 4].Style.Numberformat.Format = euroFormat; 
                    }
                    if (products[i].PricePerKgOrLiter1 != 0)
                    {
                        worksheet.Cells[i + offsetRow, 5].Value = products[i].PricePerKgOrLiter1;
                        worksheet.Cells[i + offsetRow, 5].Style.Numberformat.Format = euroFormat; 
                    }
                    if (products[i].PricePerKgOrLiter2 != 0)
                    {
                        worksheet.Cells[i + offsetRow, 6].Value = products[i].PricePerKgOrLiter2;
                        worksheet.Cells[i + offsetRow, 6].Style.Numberformat.Format = euroFormat; 
                    }
                    worksheet.Cells[i + offsetRow, 7].Value = products[i].Badge;
                    worksheet.Cells[i + offsetRow, 8].Value = products[i].OfferStartDate;

                    if (i % 2 == 1)
                    {
                        HighLightCells(i + offsetRow, 1, i + offsetRow, columnCount, Color.LightGray, worksheet);
                    }
                    
                    // Überprüfe, ob eines der interessanten Produkten mit dabei ist. Falls ja, dann verändere die Hintergrundfarbe
                    foreach (var interestingProduct in interestingProducts)
                    {
                        string produktFullName = products[i].Name;
                        if (IsContains(interestingProduct.Name, produktFullName))
                        {

                            if (products[i].PricePerKgOrLiter1 == 0 && products[i].PricePerKgOrLiter2 == 0)
                            {
                                if (products[i].NewPrice <= interestingProduct.PriceCap)
                                {
                                    HighLightCells(i + offsetRow, 1, i + offsetRow, columnCount, Color.LightCoral, worksheet);
                                }
                            }
                            else if ((products[i].PricePerKgOrLiter1 <= interestingProduct.PriceCap && products[i].PricePerKgOrLiter1 != 0) ||
                                     (products[i].PricePerKgOrLiter2 <= interestingProduct.PriceCap && products[i].PricePerKgOrLiter2 != 0))         // es existieren teilweise 2 kg Preise, je nach Produktausführung.
                            {
                                HighLightCells(i + offsetRow, 1, i + offsetRow, columnCount, Color.LightCoral, worksheet);
                            }
                            else
                            {
                                HighLightCells(i + offsetRow, 1, i + offsetRow, columnCount, Color.Yellow, worksheet);
                            }
                        }
                    }

                    //static void HighLightCells(int row, int offsetRow, int columnCount, Color color, ExcelWorksheet worksheet)  
                    
                }
                static void HighLightCells(int fromRow, int fromCol, int toRow, int toCol, Color color, ExcelWorksheet worksheet)
                {
                    //var style = worksheet.Cells[row + offsetRow, 1, row + offsetRow, columnCount].Style;            // Bereich auswählen, welcher farblich geändert werden soll
                    var style = worksheet.Cells[fromRow, fromCol, toRow, toCol].Style;            // Bereich auswählen, welcher farblich geändert werden soll
                    style.Fill.PatternType = ExcelFillStyle.Solid;                                          // Bereich wird mit einer einheitlichen Farbe ohne Farbverlauf oder Muster eingefärbt
                    style.Fill.BackgroundColor.SetColor(color); 
                }
                //Spaltenbreite automatisch anpassen
                for (int i = 1; i <= columnCount; i++)
                {
                    worksheet.Column(i).AutoFit();
                }

                FileInfo excelFile = new FileInfo(excelFilePath);
                try
                {
                    excelPackage.SaveAs(excelFile);
                }
                catch
                {
                    Console.WriteLine("Speichern fehlgeschlagen.");
                }
            }
        }
        /// <summary>
        /// Lade die alte Angebots Datei und extrahiere die Überschrift (enthält von wann bis wann die Angebote gültig sind).
        /// </summary>
        /// <param name="excelPath"></param>
        /// <returns></returns>
        static string ExtractHeadlineFromExcel(string excelPath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Excel-Paket erstellen und die Excel-Datei laden
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excelPath)))
            {
                if (!File.Exists(excelPath))
                    return string.Empty;

                // Das erste Arbeitsblatt auswählen (Index beginnt bei 0). Es kann aber auch der Name vom Arbeitsblatt angegeben werden
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["Penny"];

                if (worksheet != null)
                    return (string)worksheet.Cells["A1"].Value;
                else
                    return string.Empty;
            }
        }

        static void SendEMail(string mailAdress, string subject, string body)
        {
            // E-Mail-Einstellungen
            string senderEmail = "d.rothweiler83@gmx.de";
            string receiverEmail = mailAdress;
            //string receiverEmail = "d.rothweiler83@gmx.de";

            // SMTP-Server-Einstellungen
            //string smtpServer = "smtp.mail.yahoo.com";
            string smtpServer = "mail.gmx.net";
            int smtpPort = 587; // Standard-Port für SMTP ist 587
            //string smtpUsername = "d.rothweiler@yahoo.de";
            string smtpUsername = "d.rothweiler83@gmx.de";
            //string smtpPassword = "41149512-Dominic";
            string smtpPassword = "41149512dominic";

            // Erstellen Sie eine neue SMTP-Clientinstanz
            SmtpClient client = new SmtpClient(smtpServer, smtpPort);
            client.EnableSsl = true; // SSL aktivieren, falls erforderlich
            client.Credentials = new NetworkCredential(smtpUsername, smtpPassword);

            // Erstellen Sie eine neue E-Mail-Nachricht
            MailMessage message = new MailMessage(senderEmail, receiverEmail, subject, body);

            try
            {
                // Senden Sie die E-Mail
                client.Send(message);
                Console.WriteLine("E-Mail-Benachrichtigung erfolgreich gesendet.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fehler beim Senden der E-Mail-Benachrichtigung: " + ex.Message);
            }
        }
    }
    
    class Product
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public double OldPrice { get; set;}
        public double NewPrice { get; set;}
        public double PricePerKgOrLiter1 { get; set;}
        public double PricePerKgOrLiter2 { get; set;}
        public string Badge { get; set;}
        public string OfferStartDate { get; set;}

        public Product(string name, string description, double oldPrice, double newPrice, double pricePerKgOrLiter1, double pricePerKgOrLiter2, string badge, string offerStartDate)
        {
            Name = name;
            Description = description;
            OldPrice = oldPrice;
            NewPrice = newPrice;
            PricePerKgOrLiter1 = pricePerKgOrLiter1;
            PricePerKgOrLiter2 = pricePerKgOrLiter2;
            Badge = badge;
            OfferStartDate = offerStartDate;
        }
    }

    class ProduktFavorite
    {
        public string Name { get; set; }
        public double PriceCap { get; set; }

        public ProduktFavorite(string name, double pricePerKg)
        {
            Name = name;
            PriceCap = pricePerKg;
        }
    }
}
