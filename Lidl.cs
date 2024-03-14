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


//Bug:
// - Die Produkte von der Kategorie Deluxe wurden nicht mit hinzugefügt
// - Der Preis pro Kg hat noch teilweise ein "-" vorne dran. Diese Zeichen müssen noch entfernt werden.

//Todo:
// - Der Beginn von jeden Artikel und ob der Artikel nur mit der App verfügbar ist, wenn möglich noch in die Tabelle speichern
//   Außerdem noch von wann bis wann diese Angebote gültig sind, wenn möglich (Notfalls von Penny übernehmen)

namespace LookForSpecialOffers
{
    internal class Lidl
    {
        static List<Product> products = new();
        
        internal static void ExtractOffers(IWebDriver driver, string oldPeriodHeadline)
        {
            string pathMainPage = "https://www.lidl.de/store";
            bool isNewOffersAvailable = false;                  // Sind neue Angebote vom Penny vorhanden? Falls ja, dann soll eine E-Mail verschickt werden


            //test
            //string pattern = @"(\d+\,\d+)|(\d+\.\d+)|(\.\d+)|(\d+)";
            ////string pattern = @"(\d+)(?!\d)";
            //string input = "10,58; leterasdf 3.48";
            //Regex regex = new Regex(pattern);

            //// Übereinstimmungen finden
            //MatchCollection matches = regex.Matches(input);

            
            //string amountText = matches[0].Value;
            ///test ende



            driver.Navigate().GoToUrl(pathMainPage);

            //driver.FindElement(By.Id("onetrust-accept-btn-handler"));
            //akzeptiere den Cookie Button
            var cookieAcceptBtn = (IWebElement?)WebScraperHelper.Find(driver, "onetrust-accept-btn-handler", KindOfSearchElement.FindElementByID);
            cookieAcceptBtn?.Click();

            // Gehe zu jeder Unterseite ...
            IWebElement? mainDivContainer = null;

            try
            {
                mainDivContainer = (IWebElement?)WebScraperHelper.Find(driver, "//div[@class = 'AHeroStageItems']",
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
                    list = (mainDivContainer.FindElement(By.XPath(".//ol"))).FindElements(By.TagName("li"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex}");
                    return;
                }

                Console.WriteLine("count: "  + list.Count);

                if (list == null) { return; }
                {
                    foreach (var li in list)
                    {
                        if (li == null) { continue; }
                        //verursachte eine Fehlermeldung (vermutlich wurde keine elemente mit dem Tag "a" gefunden, aber genau weiß ich es nicht, da diese nur selten auftaucht
                        //var aTag = li.FindElement(By.XPath((".//a")));
                        var aTag = (IWebElement?)WebScraperHelper.Find(li, driver, ".//a", KindOfSearchElement.FindElementByXPath,500,3);

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
                WebScraperHelper.ScrollToBottom(driver, 50, 30, 1000);
                ((IJavaScriptExecutor)driver).ExecuteScript($"window.scrollTo(0, 0);");

                IWebElement? mainDivContainer = null;       //Hauptcontainer

                try
                {
                    mainDivContainer = (IWebElement?)WebScraperHelper.Find(driver, "//div[@class = 'ATheCampaign__Page']",
                        KindOfSearchElement.FindElementByXPath);  //Sucht solange nach diesen Element, bis es erschienen ist.
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

                var sections = mainDivContainer.FindElements(By.XPath
                    (".//section[contains(@class, 'ATheCampaign__SectionWrapper') " +
                    "and contains(@class, 'APageRoot__Section') " +
                    "and contains(@class, 'ATheCampaign__SectionWrapper--relative')]"));

                foreach (var section in sections)
                {
                    IWebElement? ol = null;
                    try
                    {
                        ol = section.FindElement(By.XPath(".//div//div//ol"));
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
                    try
                    {
                        list = ol.FindElements(By.CssSelector("li.ACampaignGrid__item.ACampaignGrid__item--product"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Fehler: {ex}");
                        return;
                    }

                    if (list == null) { return; }
                    {
                        foreach (var li in list)
                        {
                            if (li == null) { continue; }

                            IWebElement? divProduct = null;
                            try
                            {
                                divProduct = li.FindElement(By.CssSelector(".product-grid-box.grid-box"));
                            }
                            catch
                            {
                                return;
                            }

                            // Extrahiere alle Informationen
                            string articleName = WebUtility.HtmlDecode(divProduct.FindElement(By.XPath
                                (".//a")).GetAttribute("aria-label"));

                            double oldPrice = 0, newPrice = 0, articlePricePerKg;
                            bool isPriceInCent = false;

                            string oldPriceText = string.Empty, newPriceText = string.Empty, articlePricePerKgText = string.Empty;

                            newPrice            = ConvertPrice(divProduct, ".m-price__price.m-price__price--small", newPriceText);
                            oldPrice            = ConvertPrice(divProduct, ".strikethrough.m-price__rrp", oldPriceText);
                            articlePricePerKg   = ConvertPrice(divProduct, ".price-footer", articlePricePerKgText, newPrice, true);

                            string description = string.Empty;
                            try 
                            {
                                description = divProduct.FindElement(By.CssSelector(".product-grid-box__amount")).Text.Trim();
                            }
                            catch
                            {
                                Console.WriteLine("Beschreibung nicht vorhanden");
                            }

                            products.Add(new Product(articleName, description, oldPrice, newPrice, articlePricePerKg, 0, string.Empty, string.Empty));
                        }
                    }
                }

                static double ExtractPrice(string input)
                {
                    double price = 0;

                    // Teile den Eingabetext am "="-Zeichen
                    string[] parts = input.Split('=');

                    // Überprüfen, ob der Eingabetext das erwartete Format hat
                    //if (parts.Length == 2)
                    {
                        // Extrahieren Sie den Teil nach dem "="-Zeichen und entfernen Sie unnötige Leerzeichen
                        price = ExtractValue(parts[1].Trim());
                        
                        return price;

                    }
                    //else
                    //{
                    //    Debug.WriteLine("ungültiges input format");
                    //    return price;          // in diesen Fall ist prices jeweils leer
                    //}
                }

                static double ExtractValue(string input)
                {
                    //Problem: Wenn ein / im input vor kommt, dann folgen mehrere kg preise
                    // in diesen Fall wird nur der erste berücksichtigt. Später soll aber jeder von diesen
                    //Preisen in eine extra Spalte gespeichert werden.

                    // Muster, um Zahlen zu extrahieren
                    //extrahiert alle zahlen im format mit folgenden Formatbeispielen
                    //2,4   6.4  .4   6
                    string pattern = @"(\d+\,\d+)|(\d+\.\d+)|(\.\d+)|(\d+)";

                    // Regulären Ausdruck erstellen
                    Regex regex = new Regex(pattern);

                    // Übereinstimmungen finden
                    //MatchCollection matches = regex.Matches(input);
                    Match match = regex.Match(input);

                    string amountText = string.Empty;

                    if (match.Success) 
                        amountText = match.Value.Replace(",",".");

                    double amount = 0;

                    if (!double.TryParse(amountText, CultureInfo.InvariantCulture, out amount))
                    {
                        Console.WriteLine($"Der Betrag konnte nicht umgewandelt werden: {amountText}");
                    }
                    return amount;
                }

                static double ConvertPrice(IWebElement divProduct, string cssSelector, string priceText, double newPrice = 0, bool isKgPriceText = false)
                {
                    double price = 0;
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
                        // sein sollte, muss hier noch Anpassungen gemacht werden.
                        if (priceText.Contains("="))
                        {
                            price = ExtractPrice(priceText);
                        }
                        //Wenn kein = vorhanden ist, dann steht der Preis dort nicht pro Kg oder pro Liter drin
                        //dann könnte man nachschauen, wie viel das Produkt selbst wiegt, indem die Zahl selbst heraus
                        // gefiltert wird und dann mit Hilfe vom Stück Preis und dem Gewicht wird dann der Preis pro
                        // Kg bzw.Liter berechnet. Vorher sollte aber herausgefunden werden, ob die Bezeichnung Kg
                        // oder "Gramm" (g) drin
                        else if (priceText.Contains("kg"))
                        {
                            double unitAmount = ExtractValue(priceText);
                            //Wenn keine Zahl gefunden wurde, liegt es wohl daran, dass dort sowas wie 'kg-Preis'
                            //nur drin steht, was ja bedeutet, dass die Menge 1 Kg sein muss. In diesen Fall wird ja die 
                            //Zahl 0 zurückgegeben
                            if (unitAmount == 0)
                                unitAmount = 1;
                            price = newPrice / unitAmount;
                            price = Math.Round(price, 2);
                        }
                    }
                    else
                    {
                        if (priceText.Contains("-") && priceText.Contains("."))
                        {
                            isPriceInCent = true;
                            priceText = priceText.Replace("-", " ").Replace(".", " ");
                        }

                        if (!double.TryParse(priceText, CultureInfo.InvariantCulture, out price))
                            Console.WriteLine($"folgender Preis konnte nicht umgewandelt werden: {priceText}");
                        if (isPriceInCent)
                        {
                            price /= 100d;
                        }
                        price = Math.Round(price, 2);
                    }
                    
                    return price;
                }
            }
            #endregion
        }
    }
}