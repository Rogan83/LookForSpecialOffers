using HtmlAgilityPack;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace LookForSpecialOffers
{
    internal static class Penny
    {
        // Extrahiert alle Sonderangebote vom Penny und speichert diese formatiert in eine Excel Tabelle ab.
        internal static void ExtractOffers(IWebDriver driver, string oldPeriodHeadline)
        {
            string pathMainPage = "https://www.penny.de/angebote";
            bool isNewOffersAvailable = false;                  // Sind neue Angebote vom Penny vorhanden? Falls ja, dann soll eine E-Mail verschickt werden

            driver.Navigate().GoToUrl(pathMainPage);

            ClickCookieButton(driver);
            EnterZipCode(driver, Program.ZipCode);
            WebScraperHelper.ScrollToBottom(driver, 50, 20, 1000);         // Es scheint so, dass es wichtig ist, dass man das herunterscrollen in vielen kleinen Steps einzuteilen, wichtig ist

            //Suche nach dem Element, wo alle links von der Kopfzeile vorhanden sind
            HtmlNode? mainContainer = (HtmlNode?)WebScraperHelper.Find(driver, "//div[contains(@class, 'tabs__content-area')]", KindOfSearchElement.SelectSingleNode);   //Sucht solange nach diesen Element, bis es erschienen ist.
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
                        foreach (var item in items)
                        {
                            var badgeContainer = item.SelectSingleNode("./article//div[@class = 'badge--split t-bg--blue-petrol t-color--white']");
                            var badge = string.Empty;       // Die Plakette, die zusätzliche Infos angibt
                            if (badgeContainer != null)
                            {
                                badge = ((HtmlNode)badgeContainer.SelectSingleNode("./span")).InnerHtml;
                            }

                            var info = item.SelectSingleNode("./article//div[@class='offer-tile__info-container']");
                            if (info == null) { continue; }

                            var articleName = WebUtility.HtmlDecode(((HtmlNode)info.SelectSingleNode("./h4[@class= 'tile__hdln offer-tile__headline']//a[@class= 'tile__link--cover']")).InnerText);

                            var articlePricePerKg = ((HtmlNode)info.SelectSingleNode("./div[@class='offer-tile__unit-price ellipsis']")).InnerText;
                            string description = articlePricePerKg.Split('(')[0].Trim();        //extrahiert die Beschreibung 
                            double articlePricePerKg1 = 0;
                            double articlePricePerKg2 = 0;

                            if (articlePricePerKg != null)
                            {
                                var price1 = ExtractPrices(articlePricePerKg)[0];
                                if (price1 != null)
                                {
                                    articlePricePerKg1 = Math.Round(double.Parse(price1, CultureInfo.InvariantCulture), 2);
                                }
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
                                newPriceText = priceContainer.SelectSingleNode("./span[@class='ellipsis bubble__price']").InnerText.
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

                InformPerEMail(isNewOffersAvailable, products);
                WebScraperHelper.SaveToExcel(products, period, Program.ExcelPath);
            }

            static string[] ExtractPrices(string input)
            {
                string[] prices = new string[2];
                if (!input.Contains("("))           // Der Preis (falls vorhanden) ist immer in der Klammer enthalten. Wenn kein Preis vorhanden ist, dann interessiert diese Info nicht und gibt einen leeren String zurück.
                    return prices;

                // Teile den Eingabetext am "="-Zeichen
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
            static void InformPerEMail(bool isNewOffersAvailable, List<Product> products)
            {
                if (isNewOffersAvailable)
                {
                    int interestingOfferCount = 0;
                    string offers = string.Empty;
                    // Als nächstes soll untersucht werden, ob von den interessanten Angeboten der Preis auch niedrig genug ist.
                    foreach (var product in products)
                    {
                        foreach (var interestingProduct in Program.InterestingProducts)
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
                        subject = "Interessante Angebote gefunden!";
                        body = $"Gute Nachricht! Folgende Angebote, welche deinen preislichen Vorstellungen entspricht, wurden gefunden: \n\n{offers}";
                    }
                    else if (interestingOfferCount == 1)
                    {
                        subject = "Es wurde ein interessantes Angebot gefunden!";
                        body = $"Gute Nachricht! Folgendes Angebot, welches deine preisliche Vorstellung entspricht, wurde gefunden: \n\n{offers}";
                    }

                    if (interestingOfferCount > 0)
                    {
                        body += "\nHier ist der Link: https://www.penny.de/angebote \nLass es dir schmecken!";
                        WebScraperHelper.SendEMail(Program.EMail, subject, body);
                    }
                }
            }

            // Sucht nach den Cookie Button und bestätigt diesen.
            static void ClickCookieButton(IWebDriver driver)
            {
                // Der Button befindet sich innerhalb der Shadow Root. Dieses kapselt die Elemente, die darin enthalten sind,
                // d.h. dass die Elemente, die sich darin befinden, von außen geschützt sind. Deswegen muss erst auf den
                // Shadow Root zugegriffen werden und von dort aus kann nach den Elementen darin gesucht werden.
                IWebElement? parent = null;
                try
                {
                    //parent = (WebElement)driver.FindElement(By.XPath("//*[@id='usercentrics-root']"));
                    parent = (IWebElement?)WebScraperHelper.Find(driver, "//*[@id='usercentrics-root']", KindOfSearchElement.FindElementByXPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"error: {ex}");
                    return;
                }

                ShadowRoot shadowRoot = (ShadowRoot)((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].shadowRoot", parent);

                // Zuerst muss der Cookie Button geklickt werden, weil dieser im Vordergrund ist und das anklicken des Button verhindert, mit dem alle Angebote eingesehen werden können. 
                //Thread.Sleep(1000);
                //var CookieAcceptBtn = shadowRoot.FindElement(By.CssSelector(".sc-dcJsrY.iWikWl"));

                IWebElement? cookieAcceptBtn = null;
                try
                {
                    cookieAcceptBtn = (IWebElement?)WebScraperHelper.Find(shadowRoot, driver, ".sc-dcJsrY.iWikWl", KindOfSearchElement.FindElementByCssSelector);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"error: {ex}");
                    return;
                }

                if (!WebScraperHelper.CheckIfInteractable(cookieAcceptBtn))
                    return;
            }


            // Gibt die Postleitzahl ein, damit regionale Angebote angezeigt werden kann
            static void EnterZipCode(IWebDriver driver, string zipCode)
            {
                //var ShowSearchForMarketBtn = driver.FindElement(By.ClassName("market-tile__btn-not-selected"));
                IWebElement? ShowSearchForMarketBtn = null;
                try
                {
                    ShowSearchForMarketBtn = (IWebElement?)WebScraperHelper.Find(driver, "market-tile__btn-not-selected", KindOfSearchElement.FindElementByClassName);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex}");
                }

                Actions actions = new Actions(driver);
                if (ShowSearchForMarketBtn != null)
                {
                    actions.MoveToElement(ShowSearchForMarketBtn).Perform();
                    // Wenn auf diesen Button geklickt wird, wird ein Eingabefeld angezeigt, wo die PLZ eingegeben werden kann
                    ShowSearchForMarketBtn.Click();
                }

                // Instanziieren der Actions-Klasse
                actions = new Actions(driver);

                // Schrittweise Eingabe des Textes
                foreach (char c in zipCode)
                {
                    // Drückt die Taste
                    actions.SendKeys(c.ToString()).Build().Perform();
                }

                //IWebElement wrapper = driver.FindElement(By.XPath("//*[@class='market-modal__wrapper']"));
                IWebElement? wrapper = null;
                try
                {
                    wrapper = (IWebElement?)WebScraperHelper.Find(driver, "//*[@class='market-modal__wrapper']", KindOfSearchElement.FindElementByXPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex}");
                }

                IWebElement? chooseMarketBtn = null;

                if (wrapper != null)
                    chooseMarketBtn = SearchForChooseMarketBtn(driver, wrapper, chooseMarketBtn, 100, 5000);

                if (chooseMarketBtn != null)
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", chooseMarketBtn);

                // Sucht solange nach diesen Button, bis er erscheint.
                static IWebElement? SearchForChooseMarketBtn(IWebDriver driver, IWebElement wrapper, IWebElement? chooseMarketBtn, int delayStep, int maxDelayTotal)
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
}
