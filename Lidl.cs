using HtmlAgilityPack;
using LookForSpecialOffers.Enums;
using Microsoft.Extensions.FileSystemGlobbing;
using Microsoft.Extensions.FileSystemGlobbing.Internal;
using Microsoft.Extensions.Primitives;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Security.Claims;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static LookForSpecialOffers.WebScraperHelper;


//Bug:
// - Wenn das Fenster minimiert ist funktioniert es nicht richtig. Ansonsten
// scheint es zu funktionieren.

//Todo:

namespace LookForSpecialOffers
{
    internal class Lidl
    {
        internal static List<Product> Products = new();
        static string pathMainPage = "https://www.lidl.de/store";
        internal static void ExtractOffers(IWebDriver driver, string oldPeriodHeadline)
        {
            List<DetailPage>? detailPages = new();

            driver.Navigate().GoToUrl(pathMainPage);

            ClickAcceptCookieBtn(driver, 3);
            ClickShowMoreBtn(driver, 3);

            ScrollThroughPage(driver, 300, 500, 500);

            detailPages = CollectDetailPageData();

            if (detailPages != null)
            {
                for (int i = 0; i < detailPages.Count; i++)
                {
                    driver.Navigate().GoToUrl(detailPages[i].Url);

                    // Alle Produktdaten von den Unterseiten extrahieren
                    ExtractSubPage(driver, detailPages[i].Url, detailPages[i].Info);
                }
            }
            //Filtert identische Einträge heraus
            Products = Products.Distinct().ToList();
            string period = string.Empty;
            SaveToExcel(Products, period, Program.ExcelPath, Discounter.Lidl);

            Program.AllProducts[Discounter.Lidl] = new List<Product>(Products);

            #region Nested Methods

            // Sucht die URL und die Infos von jeder Unterseite heraus, welche den Beginn des Angebots idr. beinhalten.
            List<DetailPage>? CollectDetailPageData()
            {
                List<DetailPage> detailPages = new List<DetailPage>();

                IWebElement? main = (IWebElement?)Searching(driver,
                    "[class*='AHeroStageGroup__Body AHeroStageGroup__Body-Current_Sales_Week']",
                    KindOfSearchElement.FindElementByCssSelector, 500, 1, "Haupt Container nicht gefunden.");

                if (main == null) { return null; }

                ReadOnlyCollection<IWebElement?>? mainDivContainers = (ReadOnlyCollection<IWebElement?>?)Searching(main, driver,
                    "./div/div", KindOfSearchElement.FindElementsByXPath, 500, 1, "Die Haupt Container, wo sich alle Unterseiten befinden, wurden nicht gefunden.");

                if (mainDivContainers == null) { return null; }

                for (int i = 0; i < mainDivContainers.Count; i++)
                {
                    if (mainDivContainers[i] == null) { continue; }

                    ReadOnlyCollection<IWebElement?>? list = null;

                    IWebElement? ol = (IWebElement?)Searching(mainDivContainers[i], driver, "ol.AHeroStageItems__List",
                        KindOfSearchElement.FindElementByCssSelector, 500, 2);
                    if (ol == null) { return null; };

                    list = (ReadOnlyCollection<IWebElement?>?)Searching(ol, driver, "li",
                        KindOfSearchElement.FindElementsByTagName, 500, 1);

                    Console.WriteLine($"Die Anzahl der Unterseiten vom Container mit dem Klassennamen: " +
                        $"{mainDivContainers[i].GetAttribute("class")} und der ID: {mainDivContainers[i].GetAttribute("id")} " +
                        $"beträgt: " + list?.Count);

                    if (list == null) { return null; }

                    //for (int pageNr = 0; pageNr < 1; pageNr++)
                    for (int pageNr = 0; pageNr < list.Count; pageNr++)
                    {
                        Console.WriteLine("Seitennummer: " + pageNr);
                        if (list[pageNr] == null) { continue; }

                        // Ab wann startet das Angebot
                        string startDate = ((IWebElement?)Searching(list[pageNr], driver,
                        "./a/div[contains(@class,'Details')]/p", KindOfSearchElement.FindElementByXPath, 500, 1,
                        "Datum vom Beginn des Angebots nicht gefunden."))?.Text ?? "";

                        IWebElement? aTag = (IWebElement?)Searching(list[pageNr], driver, "./a", KindOfSearchElement.FindElementByXPath, 500, 1);

                        string? url = aTag?.GetAttribute("href");

                        if (!string.IsNullOrEmpty(url))
                        {
                            detailPages.Add(new DetailPage(url, startDate));
                        }
                    }
                }

                return detailPages;
            }
            // Extrahiere die Seite, wo jeweils alle Produkte stehen
            static void ExtractSubPage(IWebDriver driver, string url, string startDate)
            {
                ScrollThroughPage(driver, 300, 1000, 100);

                IWebElement? mainDivContainer = null;       //Hauptcontainer
                string className = "ATheCampaign__Page";
                
                mainDivContainer = (IWebElement?)Searching(driver, $"//div[@class = '{className}']",
                    KindOfSearchElement.FindElementByXPath, 500, 2, $"Der Div Container mit dem Klassennamen {className} wurde nicht gefunden.");  //Sucht solange nach diesen Element, bis es erschienen ist oder die max. Zeit überschritten wurde

                if (mainDivContainer == null)
                {
                    return;
                }
                // ist wohl zu speziell, da auf der deluxe Seite die Sections etwas anders heißen.
                //string searchname = ".//section[contains(@class, 'ATheCampaign__SectionWrapper') " +
                //    "and contains(@class, 'APageRoot__Section') " +
                //    "and contains(@class, 'ATheCampaign__SectionWrapper--relative')]";

                string searchname = ".//section[contains(@class, 'ATheCampaign__SectionWrapper') " +
                     "and contains(@class, 'APageRoot__Section')]";
                
                ReadOnlyCollection<IWebElement?>? sections = (ReadOnlyCollection<IWebElement?>?)Searching(mainDivContainer,
                    driver, searchname, KindOfSearchElement.FindElementsByXPath, 500, 2);
                
                //Debug
                Console.WriteLine($"anzahl sections: {sections?.Count()}");
                //info
                // Die "ol" mit der klasse "ACampaignGrid" beinhaltet ein li Element mit der Überschrift:
                // "beste Preis zum wochenstart",
                // aber auch li Elemente mit Produkten mit den Klassennamen:  "ACampaignGrid__item ACampaignGrid__item--product" 
                // Letzeres ist interessant

                if (sections == null) return;

                foreach (var section in sections)
                {
                    if (section == null) continue;
                    
                    IWebElement? ol = (IWebElement?)Searching(section, driver, "./div/div/ol", KindOfSearchElement.FindElementByXPath, 500, 1, 
                        $"ol Element von der section mit der id: {section.GetAttribute("id")} nicht gefunden.");

                    if (ol == null) 
                    {
                        continue;
                    }
                    else
                    {
                        Console.WriteLine($"ol Element von der section mit der id: {section.GetAttribute("id")} gefunden.");
                    }

                    ReadOnlyCollection<IWebElement?>? liElements = (ReadOnlyCollection<IWebElement?>?)Searching(ol, driver,
                        "li.ACampaignGrid__item.ACampaignGrid__item--product",
                        KindOfSearchElement.FindElementsByCssSelector, 500, 1);


                    if (liElements == null) { Console.WriteLine("produkt liste ist null."); return; }
                    else if (liElements.Count() == 0) { Console.WriteLine("produkt liste ist leer."); return; }

                    foreach (var liElement in liElements)
                    {
                        if (liElement == null) { continue; }
                        
                        IWebElement? productInfoContainer = (IWebElement?)Searching(liElement, driver,
                            "./div/div/div[contains(@class, 'product-grid-box grid-box')]",
                            KindOfSearchElement.FindElementByXPath, 500, 2);

                        //enthält infos wie artikelname und die Beschreibung.
                        IWebElement? content = null;
                       
                        if (productInfoContainer == null) { continue; }

                        content = (IWebElement?)Searching(productInfoContainer, driver,
                            "./div[contains(@class, 'content')]",
                            KindOfSearchElement.FindElementByXPath, 500, 2);

                        string articleName = WebUtility.HtmlDecode(content?.FindElement(By.XPath
                            ("./h2"))?.Text ?? "");

                        string badge = WebUtility.HtmlDecode(content?.FindElement(By.XPath
                            ("./div[contains(@class, 'text')]"))?.Text ?? "");

                        //test
                        if (articleName.ToLower().Contains("rotkäppchen"))
                        {
                        }
                        ///

                        double newPrice = 0, oldPrice = 0;
                        List<double> articlePricesPerKg = new List<double>();
                        //bool isPriceInCent = false;

                        string oldPriceText = string.Empty, newPriceText = string.Empty, articlePricePerKgText = string.Empty,
                            description = string.Empty;

                        // suche nach dem neuen (aktuellen) Preis
                        List<double> temp = new List<double>();
                        if (productInfoContainer == null) { continue; }

                        temp = ConvertPrices(productInfoContainer, ".m-price__price.m-price__price--small", newPriceText);
                        if (temp.Count > 0)
                            newPrice = temp[0];  // Es kommt nur 1 aktueller Preis vor
                                                    // suche nach dem vorherigen Preis, welcher durchgestrichen dargestellt wird.
                        temp = ConvertPrices(productInfoContainer, ".strikethrough.m-price__rrp", oldPriceText);
                        if (temp.Count > 0)
                            oldPrice = temp[0];  // Es kommt nur 1 vorheriger Preis vor

                        articlePricesPerKg = ConvertPrices(productInfoContainer, ".price-footer", articlePricePerKgText, newPrice, true);

                        try
                        {
                            description = productInfoContainer.FindElement(By.CssSelector(".product-grid-box__amount")).Text.Trim();
                        }
                        catch
                        {
                            //Console.WriteLine("Beschreibung nicht vorhanden");
                        }
                        
                        // Es kann sein, dass keine kg Preise ermittelt bzw. gefunden werden konnten.
                        
                        if (articlePricesPerKg.Count > 0)
                        {
                            foreach (double articlePricePerKg in articlePricesPerKg)
                            {
                                var product = new Product(articleName, description, oldPrice, newPrice, articlePricePerKg, badge, startDate);
                                Products.Add(product);
                            }
                        }
                        else
                        {
                            var product = new Product(articleName, description, oldPrice, newPrice, 0, badge, startDate);
                            Products.Add(product);
                        }
                    }
                }

                static List<double> ConvertPrices(IWebElement divProduct, string cssSelector, string priceText, double newPrice = 0, bool isKgPriceText = false)
                {
                    List<double> prices = new List<double>();
                    bool isPriceInCent = false;
                    try
                    {
                        priceText = divProduct.FindElement(By.CssSelector(cssSelector)).Text;
                    }
                    catch
                    {
                        priceText = string.Empty;
                    }

                    if (isKgPriceText)
                    {
                        // Wenn das = vorhanden ist, dann steht der Kg Preis dahinter
                        // (bis jetzt bezieht sich dieser Preis dann auf 1 Kg oder 1 Liter. Falls das nicht der Fall 
                        // sein sollte, muss hier noch Anpassungen gemacht werden).
                        if (priceText.Contains("=") && (priceText.ToLower().Contains("l") || priceText.ToLower().Contains("kg")))
                        {
                            prices = ExtractPricesBehindEqualChar(priceText);
                        }
                        //Wenn kein = vorhanden ist, dann steht der Preis dort nicht pro Kg oder pro Liter drin
                        //dann könnte man nachschauen, wie viel das Produkt selbst wiegt, indem die Zahl selbst heraus
                        // gefiltert wird und dann mit Hilfe vom Stück Preis und dem Gewicht wird dann der Preis pro
                        // Kg bzw.Liter berechnet. Vorher sollte aber herausgefunden werden, ob die Bezeichnung Kg
                        // oder "Gramm" (g) drin
                        else if (priceText.Contains("kg") || (priceText.ToLower().Contains("l") && priceText.ToLower().Contains("je")))
                        {
                            List<double> unitAmount = ExtractPricesOrValues(priceText);
                            //Wenn keine Zahl gefunden wurde, liegt es wohl daran, dass dort sowas wie 'kg-Preis'
                            //nur drin steht, was ja bedeutet, dass die Menge 1 Kg sein muss. In diesen Fall wird ja die 
                            //Zahl 0 zurückgegeben
                            for (int i = 0; i < unitAmount.Count; i++)
                            {
                                if (unitAmount[i] == 0)
                                    unitAmount[i] = 1;

                                prices.Add(Math.Round(newPrice / unitAmount[i], 2));  // die kg Preise bestimmen
                            }
                        }
                    }
                    else
                    {
                        double price = 0;

                        if (priceText.Contains("-") && priceText.Contains("."))
                        {
                            isPriceInCent = true;
                            priceText = priceText.Replace("-", " ").Replace(".", " ");
                        }

                        if (!double.TryParse(priceText, CultureInfo.InvariantCulture, out price))
                        {
                            //Console.WriteLine($"folgender Preis konnte nicht umgewandelt werden: {priceText}");
                        }
                        if (isPriceInCent)
                        {
                            price /= 100d;
                        }
                        price = Math.Round(price, 2);
                        prices.Add(price);
                    }

                    return prices;

                    static List<double> ExtractPricesBehindEqualChar(string input)
                    {
                        List<double> prices = new List<double>();

                        // Teile den Eingabetext am "="-Zeichen
                        int index = input.IndexOf("=");
                        string textBehindEqualChar = string.Empty;
                        if (index != -1) // Überprüfen, ob das = gefunden wurde
                        {
                            textBehindEqualChar = input.Substring(index + 1); // Den Rest ab der Position des ersten = extrahieren
                        }

                        // Extrahiert den Teil nach dem ersten vorkommenden "="-Zeichen und wandelt diese in eine oder mehrere Zahlen um
                        prices = ExtractPricesOrValues(textBehindEqualChar);

                        return prices;
                    }


                    static List<double> ExtractPricesOrValues(string input)
                    {
                        List<double> amounts = new List<double>();

                        // Muster, um Zahlen zu extrahieren
                        //extrahiert alle zahlen im format mit folgenden Formatbeispielen
                        //2,4   6.4  .4   6
                        string pattern = @"(\d+\,\d+)|(\d+\.\d+)|(\.\d+)|(\d+)";
                        // Regulären Ausdruck erstellen
                        Regex regex = new Regex(pattern);

                        // Wenn ein / im input vor kommt, aber kein =, dann folgen mehrere kg preise
                        // Diese sollen alle einzeln extrahiert werden und in seperaten 
                        // Zeilen in die Tabelle jeweils eingetragen werden.
                        int numberOfPrices = 0;
                        if (input.Contains("/") && !input.Contains("="))
                        {
                            numberOfPrices = input.Count(c => c == '/');
                            MatchCollection matches = regex.Matches(input);

                            foreach (Match match in matches)
                            {
                                double amount = 0;
                                if (match.Success)
                                {
                                    if (!double.TryParse(match.Value, CultureInfo.InvariantCulture, out amount))
                                    {
                                        Console.WriteLine($"Der extrahierte Wert: {match.Value} konnte nicht als Zahl umgewandelt werden");
                                    }
                                }
                                amounts.Add(amount);
                            }
                        }
                        //ansonsten müsste nur ein relevanter Preis drin stehen, den man einzeln extrahieren muss
                        else
                        {
                            // Übereinstimmungen finden

                            Match match = regex.Match(input);

                            string amountText = string.Empty;

                            if (match.Success)
                                amountText = match.Value.Replace(",", ".");

                            double amount = 0;

                            if (!double.TryParse(amountText, CultureInfo.InvariantCulture, out amount))
                            {
                                Console.WriteLine($"Der Betrag konnte nicht umgewandelt werden: {amountText}");
                            }
                            amounts.Add(amount);
                        }
                        return amounts;
                    }
                }
            }

            // Suche falls vorhanden den Button, welcher alle Unterseiten anzeigen lässt.
            // Dieser erscheint nur dann, wenn besonders viele Unterseiten vorhanden sind.
            static void ClickShowMoreBtn(IWebDriver driver, int waitTime = 2)
            {
                IWebElement? showMoreBtn = (IWebElement?)Searching(driver, ".AMoreHeroStageItems__ToggleButton-label",
                    KindOfSearchElement.FindElementByCssSelector, 500, waitTime);
                
                if (showMoreBtn == null) { return; }
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", showMoreBtn);
            }

            static void ClickAcceptCookieBtn(IWebDriver driver, int waitTime = 2)
            {
                IWebElement? cookieAcceptBtn = (IWebElement?)Searching(driver, "onetrust-accept-btn-handler",
                KindOfSearchElement.FindElementByID, 500, waitTime);
               
                cookieAcceptBtn?.Click();
            }
            #endregion
        }
    }
}