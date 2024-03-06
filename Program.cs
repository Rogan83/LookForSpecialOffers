using HtmlAgilityPack;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections;
using System.Diagnostics;
using System.Drawing;
using System.Reflection.Metadata;
using System.Text.RegularExpressions;

namespace LookForSpecialOffers
{
    class Program
    {
        static void Main(string[] args) 
        {
            ChromeOptions options = new ChromeOptions();
            //options.AddArgument("--headless");              //öffnet die seiten im hintergrund
            using (IWebDriver driver = new ChromeDriver(options))
            {
                string pathMainPage = "https://www.penny.de";
                driver.Navigate().GoToUrl(pathMainPage);
                
                GoToOffersPage(driver, pathMainPage);      //Scheint jetzt richtig zu gehen

                ScrollToBottom(driver, 200, 10);         // Es könnte sein, dass die Zeit nicht ausreicht. Vllt sollte ich, falls auf ein Element nicht zugegriffen werden kann, diese Methode wiederholen

                string searchName = "//div[contains(@class, 'tabs__content-area')]";      //Suche nach dem Element, wo alle links von der Kopfzeile vorhanden sind
                var mainContainer = (HtmlNode)FindObject(driver, searchName, KindOfSearchElement.SelectSingleNode, 100, 10);  //Sucht solange nach diesen Element, bis es erschienen ist.
                List<Product> products = new();

                if (mainContainer != null )     //Der maincontainer enthält alles relevantes
                {
                    var mainSection = mainContainer.SelectSingleNode("./section[@class='tabs__content tabs__content--offers t-bg--wild-sand ']");
                    var articleContainers = mainSection.SelectNodes("./div[@class='js-category-section']");

                    foreach(var articleContainer in articleContainers)
                    {
                        var offerStartDate = articleContainer.Attributes["id"].Value.Replace('-',' ');  //Ab wann gilt dieses Angebot
                        var articleContainerSections = articleContainer.SelectNodes("./section");

                        foreach (var articleContainerSection in articleContainerSections)
                        {
                            var weekdayHeadline = articleContainerSection.Attributes["id"].Value;

                            var list = articleContainerSection.SelectSingleNode("./div[@class='l-container']//ul[@class='tile-list']");
                            var items = list.SelectNodes("./li");

                            foreach (var item in items)
                            {
                                //string notAvailableText = "nicht vorhanden bzw. angegeben";         // Dieser Text soll abgespeichert werden, wenn ein Preis Feld leer ist.

                                var info = item.SelectSingleNode("./article//div[@class='offer-tile__info-container']");
                                if (info == null) { continue; }         // das erste item hat keinen Artikel mit der Klasse. Deswegen muss dieser übersprungen werden

                                var articleName = ((HtmlNode)info.SelectSingleNode("./h4[@class= 'tile__hdln offer-tile__headline']//a[@class= 'tile__link--cover']")).InnerText;

                                var articlePricePerKg = ((HtmlNode)info.SelectSingleNode("./div[@class='offer-tile__unit-price ellipsis']")).InnerText;
                                string description = articlePricePerKg.Split('(')[0].Trim();        //extrahiert die Beschreibung 
                                string articlePricePerKg1 = string.Empty;
                                string articlePricePerKg2 = string.Empty;

                                if (articlePricePerKg != null)
                                {
                                    //articlePricePerKg1 = float.Parse((ExtractPrices(articlePricePerKg))[0]);
                                    //var price1 = ExtractPrices(articlePricePerKg)[0];
                                    articlePricePerKg1 = ExtractPrices(articlePricePerKg)[0];
                                    //if (price1 != null)
                                    //    articlePricePerKg2 = float.Parse(price1);
                                    //var price2 = ExtractPrices(articlePricePerKg)[1];
                                    articlePricePerKg2 = ExtractPrices(articlePricePerKg)[1];
                                    //if (price2 != null)
                                    //    articlePricePerKg2 = float.Parse(price2);
                                }
                                //if (articlePricePerKg1 == String.Empty || articlePricePerKg == null)
                                //    articlePricePerKg1 = notAvailableText;
                                //if (articlePricePerKg2 == String.Empty || articlePricePerKg == null)
                                //    articlePricePerKg2 = notAvailableText;


                                var priceContainer = item.SelectSingleNode("./article" +
                                    "//div[contains(@class, 'bubble offer-tile')]" +
                                    "//div");

                                string oldPriceText = "", newPriceText = "";
                                float oldPrice = 0, newPrice = 0;

                                var price = priceContainer.SelectSingleNode("./div//span[@class='value']");
                                if (price != null)
                                {
                                    oldPriceText = price.InnerText.Replace(',', ' ').Replace('–', ' ').Replace('.',',');
                                    //oldPrice = float.Parse(oldPriceText);
                                }

                                price = priceContainer.SelectSingleNode("./span[@class='ellipsis bubble__price']");
                                if (price != null)
                                {
                                    newPriceText = (priceContainer.SelectSingleNode("./span[@class='ellipsis bubble__price']")).InnerText.
                                        Replace(',', ' ').Replace('–', ' ').Replace('*', ' ').Replace('.', ',');
                                    //newPrice = float.Parse(newPriceText);
                                }

                                products.Add(new Product(articleName, description, oldPriceText, newPriceText, articlePricePerKg1, articlePricePerKg2, offerStartDate));
                            }
                        }
                    }

                    // Der Zeitraum, von wann bis wann die Angebote gelten
                    var period = ((HtmlNode)mainSection.SelectSingleNode("./div[@class = 'category-menu']" +
                        "//div[@class = 'category-menu__header-wrapper']" +
                        "//div[@class = 'category-menu__header l-container']" +
                        "//div[@class = 'category-menu__header-container']" +
                        "//div//div//div")).Attributes["data-startend"].Value;


                    SaveToExcel(products, period);

                    int iii = 9; //[@class = '']
                }

                driver.Quit();
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
                        prices[0] = prices[0].Replace('.', ',');
                        prices[1] = prices[1].Replace('.', ',').Replace(')', ' ');
                    }
                    else
                    {
                        // Ersetzt den Punkt durch ein Komma und die Klammer wird entfernt
                        prices[0] = valuePart.Replace('.', ',').Replace(')', ' ');
                    }

                    return prices;
                }
                else
                {
                    Debug.WriteLine("Invalid input format");
                    return prices;          // in diesen Fall ist prices jeweils leer
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
            
            long oldScrollHeight = (long)js.ExecuteScript("return document.body.scrollHeight;");
            long newScrollHeight = 0;
            //Es dauert eine Weile, bis die Scrollheight ermittelt wird. Deswegen wird die schleife so lange wiederholt, bis sind die Scrollheight nicht mehr verändert, was bedeutet, dass diese den entgültigen wert ermittelt haben muss
            while (true)
            {
                Thread.Sleep(100);
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
                System.Threading.Thread.Sleep(delayPerStep); // Wartezeit in Millisekunden anpassen
            }
        }

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

        static void SaveToExcel(List<Product> data, string period)
        {
            if (data == null)
            {
                Debug.WriteLine("No Date to save");
                return;
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string excelFilePath = "Angebote.xlsx";

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Penny");

                worksheet.Cells[1, 1].Value = $"Die Angebote vom Penny vom {period}";

                worksheet.Cells[3, 1].Value = "Name";
                worksheet.Cells[3, 2].Value = "Bezeichnung";
                worksheet.Cells[3, 3].Value = "Vorheriger Preis";
                worksheet.Cells[3, 4].Value = "Neuer Preis";
                worksheet.Cells[3, 5].Value = "Preis Pro Kg oder Liter erstes Angebot";
                worksheet.Cells[3, 6].Value = "Preis Pro Kg oder Liter zweites Angebot";
                worksheet.Cells[3, 7].Value = "Begin";

                for (int i = 0; i < data.Count; i++)
                {
                    worksheet.Cells[i + 4, 1].Value = data[i].Name;
                    worksheet.Cells[i + 4, 2].Value = data[i].Description;
                    worksheet.Cells[i + 4, 3].Value = data[i].OldPrice;
                    worksheet.Cells[i + 4, 4].Value = data[i].NewPrice;
                    worksheet.Cells[i + 4, 5].Value = data[i].PricePerKgOrLiter1;
                    worksheet.Cells[i + 4, 6].Value = data[i].PricePerKgOrLiter2;
                    worksheet.Cells[i + 4, 7].Value = data[i].OfferStartDate;
                    
                }

                //Spaltenbreite automatisch anpassen
                for (int i = 1; i <= 7; i++)
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
                    Console.WriteLine("Saving is failed");
                }
            }
        }
    }

    class Product
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string OldPrice { get; set;}
        public string NewPrice { get; set;}
        public string PricePerKgOrLiter1 { get; set;}
        public string PricePerKgOrLiter2 { get; set;}
        public string OfferStartDate { get; set;}

        public Product(string name, string description, string oldPrice, string newPrice, string pricePerKgOrLiter1, string pricePerKgOrLiter2, string offerStartDate)
        {
            Name = name;
            Description = description;
            OldPrice = oldPrice;
            NewPrice = newPrice;
            PricePerKgOrLiter1 = pricePerKgOrLiter1;
            PricePerKgOrLiter2 = pricePerKgOrLiter2;
            OfferStartDate = offerStartDate;
        }
    }
}








