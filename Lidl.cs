using HtmlAgilityPack;
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
using System.Threading.Tasks;

namespace LookForSpecialOffers
{
    internal class Lidl
    {
        static List<Product> products = new();

        internal static void ExtractOffers(IWebDriver driver, string oldPeriodHeadline)
        {
            string pathMainPage = "https://www.lidl.de/store";
            bool isNewOffersAvailable = false;                  // Sind neue Angebote vom Penny vorhanden? Falls ja, dann soll eine E-Mail verschickt werden


            driver.Navigate().GoToUrl(pathMainPage);

            // Gehe zu jeder Unterseite ...
            IWebElement? mainDivContainer = null;

            try
            {
                mainDivContainer = (IWebElement?)WebScraperHelper.Find(driver, "//div[@class = 'AHeroStageItems']",
                    KindOfSearchElement.FindElementByXPath);  //Sucht solange nach diesen Element, bis es erschienen ist.
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
                        var aTag = li.FindElement(By.XPath((".//a")));

                        string url = string.Empty;
                        if (aTag != null)
                        {
                            url = aTag.GetAttribute("href");
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

            static void ExtractSubPage(IWebDriver driver, string url)
            {

                //WebScraperHelper.ScrollToBottom(driver, 50, 20, 1000);  

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

                            double oldPrice = 0, newPrice = 0;
                            bool isPriceInCent = false;

                            string oldPriceText = string.Empty, newPriceText = string.Empty;

                            newPrice = ConvertPrice(divProduct, ".m-price__price.m-price__price--small", newPriceText);
                            oldPrice = ConvertPrice(divProduct, ".strikethrough.m-price__rrp", oldPriceText);

                            string description = string.Empty;
                            try 
                            {
                                description = divProduct.FindElement(By.CssSelector(".product-grid-box__amount")).Text.Trim();
                            }
                            catch
                            {
                                Console.WriteLine("Beschreibung nicht vorhanden");
                            }

                            string articlePricePerKgText = divProduct.FindElement(By.CssSelector(".price-footer")).Text;
                            double articlePricePerKg = 0;
                            if (!double.TryParse(ExtractPrice(articlePricePerKgText), CultureInfo.InvariantCulture, out articlePricePerKg))
                                Console.WriteLine($"folgende Zeichenkette konnte nicht umgewandelt werden: {newPriceText}");

                            products.Add(new Product(articleName,description,oldPrice,newPrice, articlePricePerKg,0,"",""));

                            int i = 0;
                        }
                    }
                }

                static string ExtractPrice(string input)
                {
                    string price = string.Empty;

                    // Teile den Eingabetext am "="-Zeichen
                    string[] parts = input.Split('=');

                    // Überprüfen, ob der Eingabetext das erwartete Format hat
                    if (parts.Length == 2)
                    {
                        // Extrahieren Sie den Teil nach dem "="-Zeichen und entfernen Sie unnötige Leerzeichen
                        return parts[1].Trim();

                    }
                    else
                    {
                        Debug.WriteLine("ungültiges input format");
                        return price;          // in diesen Fall ist prices jeweils leer
                    }
                }

                static double ConvertPrice(IWebElement divProduct, string cssSelector,  string newPriceText)
                {
                    double price = 0;
                    bool isPriceInCent = false;
                    try
                    {
                        newPriceText = divProduct.FindElement(By.CssSelector(cssSelector)).Text;
                    }
                    catch
                    {
                        newPriceText = string.Empty;
                    }

                    if (newPriceText.Contains("-") && newPriceText.Contains("."))
                    {
                        isPriceInCent = true;
                        newPriceText = newPriceText.Replace("-", " ").Replace(".", " ");
                    }

                    if (!double.TryParse(newPriceText, CultureInfo.InvariantCulture, out price))
                        Console.WriteLine($"folgende Zeichenkette konnte nicht umgewandelt werden: {newPriceText}");
                    if (isPriceInCent)
                    {
                        price /= 100d;
                    }
                    price = Math.Round(price, 2);
                    return price;
                }
            }

        }
    }
}