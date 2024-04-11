using LookForSpecialOffers.Enums;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static LookForSpecialOffers.WebScraperHelper;

namespace LookForSpecialOffers
{
    internal static class AldiNord
    {

        internal static void ExtractOffers(IWebDriver driver, string oldPeriodHeadline)
        {
            string pathMainPage = "https://www.aldi-nord.de/angebote.html";

            driver.Navigate().GoToUrl(pathMainPage);

            ClickCookieButton(driver);
            //EnterZipCode(driver, Program.ZipCode);       wenn man die plz eingibt, scheint es keine Änderungen von den Produkten zu geben, aber ich bin nicht ganz sicher.
            ScrollThroughPage(driver, 100, 1000, 500);

            //Alle Angebote extrahieren
            // Jeder einzelne Container enthält alle relevanten Infos
            string className = "mod-article-tile__content";
            ReadOnlyCollection<IWebElement?>? productInfoContainers = (ReadOnlyCollection<IWebElement?>?)Searching(driver,
                className, KindOfSearchElement.FindElementsByClassName,500,2);

            if (productInfoContainers == null)
            {
                Console.WriteLine("Es wurden keine Produkte gefunden.");
                return;
            }
            for (int i = 0; i < productInfoContainers.Count; i++)
            {
                //Product(articleName, description, oldPrice, newPrice, articlePricePerKg, badge, startDate);

                string description, articleName, badge,
                       startDate;
                decimal oldPrice = 0, newPrice = 0, articlePricePerKg = 0;

                className = "mod-article-tile__title";
                articleName = (string?)GetProductInfo(driver, productInfoContainers[i], className) ?? "";

                className = "mod-article-tile__brand";
                description = (string?)GetProductInfo(driver, productInfoContainers[i], className) ?? "";

                //className = "price__main";
                className = "price__previous";
                var temp = GetProductInfo(driver, productInfoContainers[i], className);
                if (temp != null)
                {
                    temp = temp.Replace(".", ",");
                    decimal.TryParse(temp, out oldPrice);
                }

                className = "price__wrapper";
                temp = GetProductInfo(driver, productInfoContainers[i], className);
                if (temp != null)
                {
                    temp = temp.Replace(".", ",");
                    decimal.TryParse(temp, out newPrice);
                }
            }

            int test = 0;

            #region Nested Methods
            static void ClickCookieButton(IWebDriver driver)
            {
                // Der Button befindet sich innerhalb der Shadow Root. Dieses kapselt die Elemente, die darin enthalten sind,
                // d.h. dass die Elemente, die sich darin befinden, von außen geschützt sind. Deswegen muss erst auf den
                // Shadow Root zugegriffen werden und von dort aus kann nach den Elementen darin gesucht werden.
                string id = "usercentrics-root";
                IWebElement? parent = (IWebElement?)Searching(driver, $"//*[@id='{id}']", KindOfSearchElement.FindElementByXPath, 500, 3,
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

                if (cookieBtns.Count >= 3)
                {
                    IWebElement? acceptBtn = cookieBtns[2];
                    var c = acceptBtn.GetAttribute("class");

                    if (acceptBtn == null) { return; }

                    if (!ClickButtonIfPossible(acceptBtn))
                    {
                        Console.WriteLine("Das Element kann nicht geklickt werden");
                        return;
                    }
                }
            }

            static void EnterZipCode(IWebDriver driver, string zipCode)
            {
                if (zipCode.Length != 5)
                    return;

                IWebElement? ShowSearchForMarketBtn = (IWebElement?)Searching(driver, "mod-store-picker-flyout__title", 
                    KindOfSearchElement.FindElementByClassName);

                if (ShowSearchForMarketBtn == null) { return; }

                // Wenn auf diesen Button geklickt wird, wird ein Eingabefeld angezeigt, wo die PLZ eingegeben werden kann
                //ShowSearchForMarketBtn.Click();
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", ShowSearchForMarketBtn);


                // Eingabefeld finden und PLZ einfügen
                string className = "geosuggest__input";
                IWebElement? inputFieldZipCode = (IWebElement?)Searching(driver, className,
                    KindOfSearchElement.FindElementByClassName);

                // löscht den input und gibt die PLZ ein
                inputFieldZipCode.Clear();
                inputFieldZipCode.SendKeys(zipCode);

                // Suche nach den Button zum bestätigen der PLZ
                string cssName = ".button-base.ubsf_store-finder-button";
                IWebElement? submitBtn = (IWebElement?)Searching(driver, cssName,
                    KindOfSearchElement.FindElementByCssSelector);

                if (submitBtn != null)
                {
                    submitBtn.Click();
                }
                else
                {
                    Console.WriteLine("Submit Button für die Eingabe der PLZ nicht gefunden.");
                    return;
                }

                // Wähle den Markt aus, der am nächsten dran ist und lade diese Seite
                cssName = ".button-base.ubsf_location-list-item-cta.ubsf_store-finder-button";
                IWebElement? chooseMarket = (IWebElement?)Searching(driver, cssName,
                    KindOfSearchElement.FindElementByCssSelector);
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].click();", chooseMarket);
            }

            static string? GetProductInfo(IWebDriver driver, IWebElement? productInfoContainer 
                ,string className, int maxWaitTime = 1)
            {
                IWebElement? infoContainer = (IWebElement?)Searching(productInfoContainer, driver,
                                    className, KindOfSearchElement.FindElementByClassName, 500, maxWaitTime);
                string? info = null;
                if (infoContainer != null)
                {
                    info = infoContainer.Text;
                }
                if (info != null)
                {
                    string[] parts = info.Split('\r');
                    return parts[0];
                }
                return null;
            }
            #endregion
        }


    }
}
