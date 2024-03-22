using HtmlAgilityPack;
using LookForSpecialOffers.Enums;
using Microsoft.Extensions.FileSystemGlobbing;
using Microsoft.Extensions.FileSystemGlobbing.Internal;
using Microsoft.Extensions.Primitives;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static LookForSpecialOffers.WebScraperHelper;


//Bug:
// - Die Produkte von der Kategorie Deluxe wurden nicht mit hinzugefügt
// - Die liste mit den einzelnen Produkten ist teilweise leer und teilweise hat sie elemente, ka. wieso. 
// - Der Preis von Ariel wird in 2 vers. Kg Preisen unterteilt, weil die Zeichenkette hinter dem = ein / enthält.

//Todo:
// - Der Beginn von jeden Artikel und ob der Artikel nur mit der App verfügbar ist, wenn möglich noch in die Tabelle speichern.
//   Außerdem noch von wann bis wann diese Angebote gültig sind, wenn möglich (Notfalls von Penny übernehmen)

namespace LookForSpecialOffers
{
    internal class Lidl
    {
        static List<Product> products = new();
        static string pathMainPage = "https://www.lidl.de/store";
        internal static void ExtractOffers(IWebDriver driver, string oldPeriodHeadline)
        {
            
            bool isNewOffersAvailable = false;                  // Sind neue Angebote vom Penny vorhanden? Falls ja, dann soll eine E-Mail verschickt werden

            driver.Navigate().GoToUrl(pathMainPage);

            //driver.FindElement(By.Id("onetrust-accept-btn-handler"));
            //akzeptiere den Cookie Button
            //var cookieAcceptBtn = (IWebElement?)Searching(driver, "onetrust-accept-btn-handler", KindOfSearchElement.FindElementByID);
            //cookieAcceptBtn?.Click();

            // Gehe zu jeder Unterseite ...
            IWebElement? mainDivContainer = null;

            try
            {
                mainDivContainer = (IWebElement?)Searching(driver, "//div[@class = 'AHeroStageItems']",
                    KindOfSearchElement.FindElementByXPath,500,3);  //Sucht solange nach diesen Element, bis es erschienen ist.
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error{ex.Message}");
            }

            if (mainDivContainer != null)
            {
                ReadOnlyCollection<IWebElement?>? list = null;
                try
                {
                    //list = (mainDivContainer.FindElement(By.XPath(".//ol"))).FindElements(By.TagName("li"));
                    IWebElement? ol = (IWebElement?)Searching(mainDivContainer, driver, ".//ol", KindOfSearchElement.FindElementByXPath, 500, 3);
                    list = (ReadOnlyCollection<IWebElement?>?)Searching(ol, driver, "li", KindOfSearchElement.FindElementsByTagName); 
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex}");
                    return;
                }

                Console.WriteLine("anzahl Unterseiten: "  + list?.Count);

                if (list == null) { return; }
                {
                    //foreach (var li in list)
                    for (int page = 0; page < list.Count; page++)
                    {
                        Console.WriteLine("Seitennummer: " + page);
                        if (list[page] == null) { continue; }
                        //verursachte eine Fehlermeldung (vermutlich wurde keine elemente mit dem Tag "a" gefunden, aber genau weiß ich es nicht, da diese nur selten auftaucht
                        //var aTag = li.FindElement(By.XPath((".//a")));
                        var aTag = (IWebElement?)Searching(list[page], driver, ".//a", KindOfSearchElement.FindElementByXPath,500,3);

                        string url = string.Empty;
                        if (aTag != null)
                        {
                            url = aTag.GetAttribute("href");
                        }
                        else
                        {
                            Console.WriteLine("Das HTML Element 'a' wurde nicht gefunden");
                        }

                        if (!string.IsNullOrEmpty(url))
                        {
                            driver.Navigate().GoToUrl(url);

                            // Unterseite extrahieren
                            ExtractSubPage(driver, url);

                            driver.Navigate().Back();
                        }
                    }
                }
            }
            string period = string.Empty;
            WebScraperHelper.SaveToExcel(products, period, Program.ExcelPath, Discounter.Lidl);

            #region verschachtelte Methode(n)
            static void ExtractSubPage(IWebDriver driver, string url)
            {
                ScrollToBottom(driver, 300, 1000, 500);
                //((IJavaScriptExecutor)driver).ExecuteScript($"window.scrollTo(0, 0);");

                IWebElement? mainDivContainer = null;       //Hauptcontainer

                try
                {
                    mainDivContainer = (IWebElement?)Searching(driver, "//div[@class = 'ATheCampaign__Page']",
                        KindOfSearchElement.FindElementByXPath);  //Sucht solange nach diesen Element, bis es erschienen ist oder die max. Zeit überschritten wurde
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error{ex.Message}");
                }

                if (mainDivContainer == null) 
                {
                    Console.WriteLine($"Main Container auf folgender Unterseite nicht gefunden: {url}");
                    return; 
                }

                //Thread.Sleep(100);
                //ReadOnlyCollection<IWebElement> sections = mainDivContainer.FindElements(By.XPath
                //    (".//section[contains(@class, 'ATheCampaign__SectionWrapper') " +
                //    "and contains(@class, 'APageRoot__Section') " +
                //    "and contains(@class, 'ATheCampaign__SectionWrapper--relative')]"));

                //Console.WriteLine("datentyp von sections: "+sections.GetType());
                //Console.WriteLine("anzahl: "+sections.Count);

                string searchname = ".//section[contains(@class, 'ATheCampaign__SectionWrapper') " +
                    "and contains(@class, 'APageRoot__Section') " +
                    "and contains(@class, 'ATheCampaign__SectionWrapper--relative')]";

                ReadOnlyCollection<IWebElement?>? sections = (ReadOnlyCollection<IWebElement?>?)Searching(mainDivContainer,
                    driver, searchname, KindOfSearchElement.FindElementsByXPath, 500, 3);
                //Debug
                Console.WriteLine($"anzahl sections: {sections?.Count()}");

                if (sections == null) return;

                foreach (var section in sections)
                {
                    if (section == null) continue;

                    IWebElement? ol = null;
                    try
                    {
                        //ol = section.FindElement(By.XPath(".//div//div//ol"));
                        ol = (IWebElement?)Searching(section, driver, ".//div//div//ol", KindOfSearchElement.FindElementByXPath);
                    }
                    catch 
                    {
                        return;
                    }

                    if (ol == null)
                    {
                        Console.WriteLine($"ol Element auf folgender Seite nicht gefunden: {url}");
                        return;
                    }

                    ReadOnlyCollection<IWebElement?>? list = null;
                    //Thread.Sleep(500);
                    // Bug. Die liste ist teilweise leer und teilweise hat sie elemente, ka. wieso. 
                    
                    {
                        list = (ReadOnlyCollection<IWebElement?>?)Searching(ol, driver, "li.ACampaignGrid__item.ACampaignGrid__item--product div div div.product-grid-box.grid-box",
                            KindOfSearchElement.FindElementsByCssSelector);
                        int listAmount = list.Count();
                        listAmount = list != null ? list.Count() : 0;
                        int count = 0;
                        int maxCount = 20;
                        while (listAmount == 0)
                        {
                            if (count >= maxCount)
                            {
                                Console.WriteLine($"Es wurden nach {count} erneuten Laden der Seite keine Listen Elemente gefunden!");
                                break;
                            }
                            count++;

                            listAmount = list != null? list.Count() : 0;
                            Console.WriteLine("anzahl listen items: " + listAmount);

                            driver.Navigate().GoToUrl(pathMainPage);
                            Thread.Sleep(100);
                            driver.Navigate().GoToUrl(url);
                            ScrollToBottom(driver, 300, 1000, 500);
                            ((IJavaScriptExecutor)driver).ExecuteScript($"window.scrollTo(0, 0);");
                            Thread.Sleep(100);
                            //list = (ReadOnlyCollection<IWebElement?>?)Searching(ol, driver, "li.ACampaignGrid__item.ACampaignGrid__item--product div div div.product-grid-box.grid-box",
                            //KindOfSearchElement.FindElementsByCssSelector);
                            //list = ol.FindElements(By.CssSelector
                            //    ("li.ACampaignGrid__item.ACampaignGrid__item--product div div div.product-grid-box.grid-box"));
                            list = FindListOfProducts("li.ACampaignGrid__item.ACampaignGrid__item--product div div div.product-grid-box.grid-box");
                            listAmount = list != null ? list.Count() : 0;
                        }
                        //list = ol.FindElements(By.CssSelector("li.ACampaignGrid__item.ACampaignGrid__item--product"));
                        ReadOnlyCollection<IWebElement?>?FindListOfProducts(string name)
                        {
                            try
                            {
                                // Versuche, das Element zu finden
                                ReadOnlyCollection<IWebElement?>? elements = ol.FindElements(By.CssSelector("li.ACampaignGrid__item.ACampaignGrid__item--product div div div.product-grid-box.grid-box"));

                                // Überprüfe, ob die Liste leer ist
                                if (elements.Count == 0)
                                {
                                    // Das Element wurde nicht gefunden, gib null zurück
                                    return null;
                                }

                                // Das Element wurde gefunden, gib die Liste der Elemente zurück
                                return elements;
                            }
                            catch 
                            {
                                // Falls eine NoSuchElementException auftritt, gib null zurück
                                return null;
                            }
                        }
                    }
                    //catch (Exception ex)
                    //{
                    //    Console.WriteLine($"Fehler: {ex}");
                    //    return;
                    //}

                    if (list == null) { Console.WriteLine("produkt liste ist null."); return; }
                    else if (list.Count() == 0) { Console.WriteLine("produkt liste ist leer."); return; }

                    foreach (var li in list)
                    {
                        if (li == null) { continue; }

                        IWebElement? divProduct = null;
                        try
                        {
                            //divProduct = li.FindElement(By.CssSelector(".product-grid-box.grid-box"));
                            divProduct = li;
                            //divProduct = (IWebElement?)Searching(li, driver, ".product-grid-box.grid-box", KindOfSearchElement.FindElementByCssSelector,500,15);
                            Console.WriteLine("typ vom div: "+divProduct.GetType());
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"div container für die Produkte nicht gefunden. Fehlermeldung: {ex}");
                            return;
                        }

                        // Extrahiere alle Informationen
                        string articleName = WebUtility.HtmlDecode(divProduct.FindElement(By.XPath
                            (".//a")).GetAttribute("aria-label"));


                        //test
                        if (articleName.ToLower().Contains("dash"))
                        {
                            int test = 0;
                        }
                        ///

                        double newPrice = 0, oldPrice = 0;
                        List<double> articlePricesPerKg;
                        bool isPriceInCent = false;

                        string oldPriceText = string.Empty, newPriceText = string.Empty, articlePricePerKgText = string.Empty;

                        List<double> temp = ConvertPrices(divProduct, ".m-price__price.m-price__price--small", newPriceText);
                        if (temp.Count > 0)
                            newPrice = temp[0];  // Es kommt nur 1 aktueller Preis vor

                        temp = ConvertPrices(divProduct, ".strikethrough.m-price__rrp", oldPriceText);
                        if (temp.Count > 0)
                            oldPrice = temp[0];  // Es kommt nur 1 vorheriger Preis vor

                        articlePricesPerKg = ConvertPrices(divProduct, ".price-footer", articlePricePerKgText, newPrice, true);

                        string description = string.Empty;
                        try 
                        {
                            description = divProduct.FindElement(By.CssSelector(".product-grid-box__amount")).Text.Trim();
                        }
                        catch
                        {
                            Console.WriteLine("Beschreibung nicht vorhanden");
                        }

                        // Es kann sein, dass keine kg Preise ermittelt bzw. gefunden werden konnten.
                        if (articlePricesPerKg.Count > 0)
                        {
                            foreach (double articlePricePerKg in articlePricesPerKg)
                            {
                                products.Add(new Product(articleName, description, oldPrice, newPrice, articlePricePerKg, string.Empty, string.Empty));
                            }
                        }
                        else
                        {
                            products.Add(new Product(articleName, description, oldPrice, newPrice, 0, string.Empty, string.Empty));
                        }
                            
                    }
                }

                static List<double> ConvertPrices(IWebElement divProduct, string cssSelector, string priceText, double newPrice = 0, bool isKgPriceText = false)
                {
                    List <double> prices = new List<double>();
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
            #endregion
        }
    }
}