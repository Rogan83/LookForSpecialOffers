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
using LookForSpecialOffers.Enums;
using System.Collections.ObjectModel;
using static LookForSpecialOffers.WebScraperHelper;
using System.Text.RegularExpressions;
using LookForSpecialOffers.Models;

namespace LookForSpecialOffers
{
    internal static class Penny
    {
        internal static List<Product> Products { get; set; } = new();
        // Extrahiert alle Sonderangebote vom Penny und speichert diese formatiert in eine Excel Tabelle ab.
        internal static void ExtractOffers(IWebDriver driver, string oldPeriodHeadline)
        {
            string pathMainPage = "https://www.penny.de/angebote";

            driver.Navigate().GoToUrl(pathMainPage);

            ClickCookieButton(driver);
            EnterZipCode(driver, Program.ZipCode);
            ScrollThroughPage(driver, 300, 2000, 500);         // Es scheint so, dass es wichtig ist, dass man das herunterscrollen in vielen kleinen Steps einzuteilen, wichtig ist

            //Suche nach dem Element, wo alle links von der Kopfzeile vorhanden sind
            HtmlNode? mainContainer = (HtmlNode?)Searching(driver, "//div[contains(@class, 'tabs__content-area')]", KindOfSearchElement.SelectSingleNode);   //Sucht solange nach diesen Element, bis es erschienen ist.
            
            if (mainContainer != null)     //Der maincontainer enthält alles relevantes
            {
                var mainSection = mainContainer.SelectSingleNode("./section[@class='tabs__content tabs__content--offers t-bg--wild-sand ']");
                var articleDivContainers = mainSection.SelectNodes("./div[@class='js-category-section']");

                foreach (var articleDivContainer in articleDivContainers)
                {
                    //Ab wann gilt dieses Angebot
                    var offerStartDate = articleDivContainer.Attributes["id"].Value.Replace('-', ' ').
                        Replace("ae","ä").Replace("oe","ö").Replace("ue","ü").Replace("montag","Montag").
                        Replace("dienstag","Dienstag").Replace("mittwoch","Mittwoch").
                        Replace("donnerstag","Donnerstag").Replace("freitag","Freitag");                          

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

                            HtmlNode articleNode = info.SelectSingleNode("./h4[@class= 'tile__hdln offer-tile__headline']//a[@class= 'tile__link--cover']");
                            string articleName = string.Empty;
                            if (articleNode != null)
                                articleName = WebUtility.HtmlDecode(articleNode.InnerText).Replace("*"," ");

                            var articlePricesPerKgNode = ((HtmlNode)info.SelectSingleNode
                                ("./div[@class='offer-tile__unit-price ellipsis']"));
                            string articlePricesPerKg = string.Empty;
                            if (articlePricesPerKgNode != null)
                                articlePricesPerKg = articlePricesPerKgNode.InnerText;

                            string description = articlePricesPerKg.Split('(')[0].Trim();        //extrahiert die Beschreibung, welche vor dem Kilo Preis steht

                            var priceContainer = item.SelectSingleNode("./article" +
                                "//div[contains(@class, 'bubble offer-tile')]" +
                                "//div");

                            decimal oldPrice = 0, newPrice = 0;

                            if (priceContainer != null)
                            {
                                var priceElement = priceContainer.SelectSingleNode("./div//span[@class='value']");
                                if (priceElement != null)
                                {
                                    var prices = ExtractPrices(priceElement.InnerText);
                                    if (prices != null && prices.Count >= 1)
                                        oldPrice = prices[0];
                                }

                                priceElement = priceContainer.SelectSingleNode("./span[@class='ellipsis bubble__price']");
                                if (priceElement != null)
                                {
                                    var prices = ExtractPrices(priceElement.InnerText);
                                    if (prices != null && prices.Count >= 1)
                                        newPrice = prices[0];
                                }
                            }

                            var pricesPerKg = ExtractPrices(articlePricesPerKg, newPrice);

                            foreach (var pricePerKg in pricesPerKg)
                            {
                                var product = new Product(articleName, description, oldPrice, newPrice, pricePerKg, badge, offerStartDate);
                                Products.Add(product);
                            }
                        }
                    }
                }

                // Der Zeitraum, von wann bis wann die Angebote gelten
                var period = ((HtmlNode)mainSection.SelectSingleNode("./div[@class = 'category-menu']" +
                    "//div[@class = 'category-menu__header-wrapper']" +
                    "//div[@class = 'category-menu__header l-container']" +
                    "//div[@class = 'category-menu__header-container']" +
                    "//div//div//div")).Attributes["data-startend"].Value;


                // Die Produkte von der Excel Tabelle
                //var loadedProducts = LoadFromExcel(Program.ExcelFilePath, Discounter.Penny);
                // Überprüft, ob die beiden Listen identisch sind.
                //var isEpual = loadedProducts.SequenceEqual(Products);
                //if (!isEpual)
                //{
                //    // Wenn die Listen nicht identisch sind, dann gibt es neue Angebote und der User soll über diese per
                //    // E-mail informiert werden.
                //    Program.IsNewOffersAvailable = true;
                //}

                if (oldPeriodHeadline == string.Empty || !oldPeriodHeadline.Contains(period))
                {
                    Program.IsNewOffersAvailable = true;
                }

                SaveToExcel(Products, period, Program.ExcelFilePath, MarketEnum.Penny);
            }

            Program.AllProducts[MarketEnum.Penny] = new List<Product>(Products);

            #region Nested Methods
            static List<decimal> ExtractPrices(string input, decimal newPrice = 0)
            {
                string pattern = @"(\d+\,\d+)|(\d+\.\d+)|(\.\d+)|(\d+)";
                // Regulären Ausdruck erstellen
                Regex regex = new Regex(pattern);

                List<decimal> prices = new List<decimal>();
                
                // Teile den Eingabetext am "="-Zeichen, wenn vorhanden
                if (input.ToLower().Contains("="))
                {
                    string[] parts = input.Split('=');

                    // Extrahieren Sie den Teil nach dem "="-Zeichen und entferne unnötige Leerzeichen
                    string valuePart = parts[1].Trim();

                    if (valuePart.Contains('/'))
                    {
                        MatchCollection matches = regex.Matches(input);

                        foreach (Match match in matches)
                        {
                            decimal amount = 0;
                            if (match.Success)
                            {
                                if (!decimal.TryParse(match.Value.Replace(",", "."), CultureInfo.InvariantCulture, out amount))
                                {
                                    Console.WriteLine($"Der extrahierte Wert: {match.Value} konnte nicht als Zahl umgewandelt werden");
                                }
                                else
                                {
                                    amount = Math.Round(amount, 2);
                                }
                            }
                            prices.Add(amount);
                        }
                    }
                    else
                    {
                        decimal price = ExtractSinglePriceOrValue(valuePart, regex);
                        prices.Add(price);

                    }
                }
                else if (input.Contains("kg") || (input.ToLower().Contains("l") && input.ToLower().Contains("je")))
                {
                    decimal value = ExtractSinglePriceOrValue(input, regex);
                    if (value == 0) { value = 1; }
                    if (newPrice != 0)
                    {
                        prices.Add(Math.Round(newPrice / value, 2));
                    }
                }
                else
                {
                    decimal price = ExtractSinglePriceOrValue(input, regex);
                    prices.Add(price);
                }

                return prices;

                static decimal ExtractSinglePriceOrValue(string input, Regex regex)
                {
                    Match match = regex.Match(input);
                    string priceText = string.Empty;
                    if (match.Success)
                    {
                        priceText = match.Value.Replace(",", ".");
                    }
                    decimal price = 0;

                    if (!decimal.TryParse(match.Value, CultureInfo.InvariantCulture, out price))
                    {
                        Console.WriteLine($"Der extrahierte Wert: {match.Value} konnte nicht als Zahl umgewandelt werden");
                    }
                    else
                    {
                        price = Math.Round(price, 2);
                    }
                    return price;
                }
            }

            // Sucht nach den Cookie Button und bestätigt diesen.
            static void ClickCookieButton(IWebDriver driver)
            {
                // Der Button befindet sich innerhalb der Shadow Root. Dieses kapselt die Elemente, die darin enthalten sind,
                // d.h. dass die Elemente, die sich darin befinden, von außen geschützt sind. Deswegen muss erst auf den
                // Shadow Root zugegriffen werden und von dort aus kann nach den Elementen darin gesucht werden.
                string id = "usercentrics-root";
                IWebElement? parent = (IWebElement?)Searching(driver, $"//*[@id='{id}']", KindOfSearchElement.FindElementByXPath,500,10, 
                    $"Das Element mit der Id {id} wurde nicht gefunden");

                ShadowRoot shadowRoot = (ShadowRoot)((IJavaScriptExecutor)driver).ExecuteScript("return arguments[0].shadowRoot", parent);

                // Zuerst muss der Cookie Button geklickt werden, weil dieser im Vordergrund ist und das anklicken des Button verhindert, mit dem alle Angebote eingesehen werden können. 
                Thread.Sleep(1000);
                //cookieAcceptBtn = shadowRoot.FindElement(By.CssSelector(".sc-dcJsrY.hCLQdG"));          // der klassenname ändert sich andauernd vom cookie btn. Das muss noch geändert werden.
                //ReadOnlyCollection<IWebElement?> cookieBtns = shadowRoot.FindElements(By.CssSelector(".sc-dcJsrY"));          // anscheinend ist die erste Klasse von den beiden Klassen immer gleich und nur die 2. Klasse ändert sich
                string className = ".sc-dcJsrY";
                ReadOnlyCollection<IWebElement?>? cookieBtns = (ReadOnlyCollection<IWebElement?>?)Searching(shadowRoot, driver, className, 
                    KindOfSearchElement.FindElementsByCssSelector, 500, 3, $"kein Cookie Button gefunden");

                if (cookieBtns == null) { return; }

                if (cookieBtns.Count >= 2)
                {
                    IWebElement? acceptBtn = cookieBtns[2];
                    if (acceptBtn == null) { return; }

                    if (!ClickButtonIfPossible(acceptBtn))
                    {
                        Console.WriteLine("Das Element kann nicht geklickt werden");
                        return;
                    }
                }
            }

            // Gibt die Postleitzahl ein, damit regionale Angebote angezeigt werden kann
            static void EnterZipCode(IWebDriver driver, string zipCode)
            {
                if (zipCode.Length != 5)
                    return;

                IWebElement? ShowSearchForMarketBtn = (IWebElement?)Searching(driver, "market-tile__btn-not-selected", KindOfSearchElement.FindElementByClassName);

                Actions actions = new Actions(driver);
                if (ShowSearchForMarketBtn == null) { return; }

                actions.MoveToElement(ShowSearchForMarketBtn).Perform();
                // Wenn auf diesen Button geklickt wird, wird ein Eingabefeld angezeigt, wo die PLZ eingegeben werden kann
                ShowSearchForMarketBtn.Click();

                actions = new Actions(driver);

                // Schrittweise Eingabe des Textes
                foreach (char c in zipCode)
                {
                    // Drückt die Taste
                    actions.SendKeys(c.ToString()).Build().Perform();
                }

                IWebElement? wrapper = (IWebElement?)Searching(driver, "//*[@class='market-modal__wrapper']", KindOfSearchElement.FindElementByXPath);
                
                //IWebElement? chooseMarketBtn = null;

                if (wrapper == null) { return; }

                //chooseMarketBtn = SearchForChooseMarketBtn(driver, wrapper, chooseMarketBtn, 100, 5000);
                string xPath = ".//main//div[@class='market-modal__results']//ul//li[1]//article//div//div//a";
                IWebElement? chooseMarketBtn = (IWebElement?)Searching(wrapper, driver, xPath, KindOfSearchElement.FindElementByXPath, 500, 3,
                    "Der Button, mit welchen der Markt ausgewählt werden kann, wurde nicht gefunden.");

                if (chooseMarketBtn != null)
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", chooseMarketBtn);
            }
            #endregion
        }
    }
}