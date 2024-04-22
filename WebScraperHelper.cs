using HtmlAgilityPack;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;
using LookForSpecialOffers.Enums;
using OpenQA.Selenium.DevTools.V120.Debugger;
using System.Runtime.InteropServices;
using System.Collections.ObjectModel;
using OpenQA.Selenium.DevTools.V120.Preload;
using LookForSpecialOffers.Models;
using System.IO;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace LookForSpecialOffers
{
    internal static class WebScraperHelper
    {
        internal static ExcelPackage excelPackage = new ExcelPackage();

        /// <summary>
        /// Überprüft, ob ein Element interagierbar ist
        /// </summary>
        /// <param name="cookieAcceptBtn"></param>
        internal static bool ClickButtonIfPossible(IWebElement cookieAcceptBtn)
        {
            bool isClickable = false;
            int waitTimeInSeconds = 10;
            int delayStep = 10;
            int maxCount = waitTimeInSeconds * 1000 / delayStep;
            int count = 0;

            while (!isClickable && count < maxCount)
            {
                if (cookieAcceptBtn.Enabled && cookieAcceptBtn.Displayed)
                {
                    isClickable = true;
                }
                Thread.Sleep(delayStep);
                count++;
            }
            try
            {
                cookieAcceptBtn.Click();                         //bestätigt die Cookies
                return true;
            }
            catch (ElementNotInteractableException ex)
            {
                Console.WriteLine($"Element not interactable: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected Exception: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Scroll stufenweise nach unten, damit die Seite komplett geladen wird.
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="delayPerStep"></param>
        /// <param name="stepsCount"></param>
        internal static void ScrollThroughPage(IWebDriver driver, int delayPerStep = 50, int scrollStep = 2000, int delayDetermineScrollHeigth = 500)
        {
            // Öffnet einen neuen Tab
            string newWindowHandle = driver.WindowHandles.Last();

            // Wechsel zu den neuen Tab
            driver.SwitchTo().Window(newWindowHandle);
            
            Random rand = new Random();

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            long oldScrollHeight = (long)js.ExecuteScript("return document.body.scrollHeight;");
            long newScrollHeight = 0;
            // Als erstes muss die Höhe der Webseite ermittelt werden. Diese verändert sich, während die Webseite Inhalt nach lädt über die Zeit.
            // Es dauert eine Weile, bis die Scrollheight ermittelt wird. Deswegen wird die schleife so lange wiederholt, bis sind die Scrollheight
            // nicht mehr verändert, was bedeuten müsste, dass diese den entgültigen wert ermittelt hat
            long newPos = 0;
            bool isDeterminedScrollHeight1Time = false;
            while (true)
            {
                // Soll nur einmalig ausgeführt werden
                while (!isDeterminedScrollHeight1Time)
                {
                    Thread.Sleep(delayDetermineScrollHeigth);
                    newScrollHeight = (long)js.ExecuteScript("return document.body.scrollHeight;");

                    if (newScrollHeight == oldScrollHeight)
                        isDeterminedScrollHeight1Time = true;
                    else
                    {
                        oldScrollHeight = newScrollHeight;
                    }
                }

                while (true)
                {
                    int min, max, randomOffset, newScrollStep, newDelayPerStep;

                    min = -(scrollStep / 5);
                    max = scrollStep / 5;
                    randomOffset = rand.Next(min, max + 1);
                    newScrollStep = scrollStep + randomOffset;

                    newPos += newScrollStep;

                    if (newPos > newScrollHeight)
                    {
                        newPos = newScrollHeight;
                        ((IJavaScriptExecutor)driver).ExecuteScript($"window.scrollTo(0, {newPos});");
                        break;
                    }

                    // Scrolle stufenweise bis zum Ende der Seite
                    ((IJavaScriptExecutor)driver).ExecuteScript($"window.scrollTo(0, {newPos});");

                    // Warte eine kurze zufällige Zeit, um ein Stück weiter herunter zu scrollen.
                    min = -(delayPerStep / 2); 
                    max = delayPerStep / 2;
                    randomOffset = rand.Next(min, max + 1);
                    newDelayPerStep = delayPerStep + randomOffset ;
                    Thread.Sleep(newDelayPerStep); // Wartezeit in Millisekunden anpassen
                }

                //Erneute Prüfung, ob jetzt eine höhere max. Scrollhöhe ermittelt wurde, die man zurücklegen kann
                newScrollHeight = (long)js.ExecuteScript("return document.body.scrollHeight;");
                if (newScrollHeight == oldScrollHeight)
                    break;
                else
                    oldScrollHeight = newScrollHeight;

            }
            ((IJavaScriptExecutor)driver).ExecuteScript($"window.scrollTo(0, 0);");

            //((IJavaScriptExecutor)driver).ExecuteScript($"window.scrollTo(0, {newPos});");

            //Scrolle noch den Rest runter
            //newScrollHeight = (long)js.ExecuteScript("return document.body.scrollHeight;");
            //((IJavaScriptExecutor)driver).ExecuteScript($"window.scrollTo(0, {newScrollHeight});");
        }

        //internal static void S()
        //{
        //    SendInput.SendKeyDown(Keys.Down);
        //}


        /// <summary>
        /// Überprüft, ob der Begriff in dem anderen Begriff vorkommt.
        /// </summary>
        /// <param name="substring"></param>
        /// <param name="product"></param>
        /// <returns></returns>

        internal static bool IsContains(string substring, string product)
        {
            substring = substring.Trim().ToLower();
            product = product.Trim().ToLower();

            return product.Contains(substring);
        }
        /// <summary>
        /// Versucht, ein bestimmtes Element zu finden und versucht es in gewissen Zeitabständen erneut, falls dieses Element nicht gefunden wird.
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="searchName"></param>
        /// <param name="searchElement"></param>
        /// <param name="pollingIntervalTime"></param>
        /// <param name="maxWaitTime"></param>
        /// <returns></returns>

        internal static object? Searching(IWebDriver driver, string searchName, KindOfSearchElement searchElement, int pollingIntervalTime = 500, int maxWaitTime = 5, string noSuchFoundExMsg = "")
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTime))
            {
                PollingInterval = TimeSpan.FromMilliseconds(pollingIntervalTime),
            };
            //wait.IgnoreExceptionTypes(typeof(ElementNotInteractableException));
            //wait.IgnoreExceptionTypes(typeof(NoSuchElementException));


            HtmlDocument? doc = null; 
            try
            {
                object? element = wait.Until<object?>(driver =>
                {
                    switch (searchElement)
                    {
                        case KindOfSearchElement.SelectSingleNode:
                            doc = new HtmlDocument();
                            doc.LoadHtml(driver.PageSource);
                            try 
                            { 
                                return doc.DocumentNode.SelectSingleNode(searchName);
                            }
                            catch { return null; }

                        case KindOfSearchElement.SelectNodes:
                            doc = new HtmlDocument();
                            doc.LoadHtml(driver.PageSource);
                            try
                            {
                                return doc.DocumentNode.SelectNodes(searchName);
                            }
                            catch { return null; }

                        case KindOfSearchElement.FindElementByCssSelector:
                            try
                            {
                                return driver.FindElement(By.CssSelector(searchName));
                            }
                            catch { return null; }

                        case KindOfSearchElement.FindElementsByCssSelector:
                            try
                            {
                                return driver.FindElements(By.CssSelector(searchName));
                            }
                            catch { return null; }

                        case KindOfSearchElement.FindElementByClassName:
                            try
                            {
                                return driver.FindElement(By.ClassName(searchName));
                            }
                            catch { return null; }

                        case KindOfSearchElement.FindElementsByClassName:
                            try
                            {
                                return driver.FindElements(By.ClassName(searchName));
                            }
                            catch { return null; }

                        case KindOfSearchElement.FindElementByXPath:
                            try
                            {
                                return driver.FindElement(By.XPath(searchName));
                            }
                            catch { return null; }

                        case KindOfSearchElement.FindElementByID:
                            try
                            {
                                return driver.FindElement(By.Id(searchName));
                            }
                            catch { return null; }

                        default:
                            throw new ArgumentException("Invalid KindOfSearchElement.");
                    }
                });

                return element;
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Element not found.");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{noSuchFoundExMsg} Error: {ex.Message}");
                return null;
            }
        }

        //internal static object? Searching3(IWebDriver driver, string searchName, KindOfSearchElement searchElement, int pollingIntervalTime = 500, int maxWaitTime = 5, string noSuchFoundExMsg = "")
        //{
        //    object? element = null;
        //    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTime))
        //    {
        //        PollingInterval = TimeSpan.FromMilliseconds(pollingIntervalTime),
        //    };
        //    wait.IgnoreExceptionTypes(typeof(ElementNotInteractableException));

        //    #region wait until
        //    var main = wait.Until(d =>
        //    {
        //        switch (searchElement)
        //        {
        //            case KindOfSearchElement.SelectSingleNode:
        //                try
        //                {
        //                    HtmlDocument doc = new HtmlDocument();
        //                    doc.LoadHtml(driver.PageSource);
        //                    if (doc != null)
        //                        element = doc.DocumentNode.SelectSingleNode(searchName);

        //                    return element != null;
        //                }
        //                catch
        //                {
        //                    return false;
        //                }

        //            case KindOfSearchElement.SelectNodes:
        //                try
        //                {
        //                    HtmlDocument doc = new HtmlDocument();
        //                    doc.LoadHtml(driver.PageSource);
        //                    if (doc != null)
        //                        element = doc.DocumentNode.SelectNodes(searchName);

        //                    return element != null;
        //                }
        //                catch
        //                {
        //                    return false;
        //                }

        //            case KindOfSearchElement.FindElementByCssSelector:
        //                try
        //                {
        //                    element = driver.FindElement(By.CssSelector(searchName));

        //                    return element != null;
        //                }
        //                catch
        //                {
        //                    return false;
        //                }

        //            case KindOfSearchElement.FindElementsByCssSelector:
        //                try
        //                {
        //                    element = driver.FindElements(By.CssSelector(searchName));

        //                    return element != null;
        //                }
        //                catch
        //                {
        //                    return false;
        //                }
        //            case KindOfSearchElement.FindElementByClassName:
        //                try
        //                {
        //                    element = driver.FindElement(By.ClassName(searchName));

        //                    return element != null;
        //                }
        //                catch
        //                {
        //                    return false;
        //                }
        //            case KindOfSearchElement.FindElementByXPath:
        //                try
        //                {
        //                    element = driver.FindElement(By.XPath(searchName));

        //                    return element != null;
        //                }
        //                catch
        //                {
        //                    Console.WriteLine(noSuchFoundExMsg);
        //                    return false;
        //                }
        //            case KindOfSearchElement.FindElementByID:
        //                try
        //                {
        //                    element = driver.FindElement(By.Id(searchName));

        //                    return element != null;
        //                }
        //                catch
        //                {
        //                    return false;
        //                }
        //            default:
        //                return false;
        //        }
        //    });
        //    #endregion

        //    return element;
        //}

        /// <summary>
        /// Versucht, ein bestimmtes Element zu finden und versucht es in gewissen Zeitabständen erneut, falls dieses Element nicht gefunden wird.
        /// </summary>
        /// <param name="iWebElement"></param>
        /// <param name="driver"></param>
        /// <param name="searchName"></param>
        /// <param name="searchElement"></param>
        /// <param name="pollingIntervalTime"></param>
        /// <param name="maxWaitTime"></param>
        /// <returns></returns>
        internal static object? Searching(IWebElement iWebElement, IWebDriver driver, string searchName, 
            KindOfSearchElement searchElement, int pollingIntervalTime = 500, int maxWaitTime = 10, string noSuchFoundExMsg = "")
        {
            object? element = null;
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTime))
            {
                PollingInterval = TimeSpan.FromMilliseconds(pollingIntervalTime), 
            };
            //wait.IgnoreExceptionTypes(typeof(ElementNotInteractableException));
            wait.IgnoreExceptionTypes(typeof(NoSuchElementException));
            element = WaitUntil(iWebElement, driver, searchName, searchElement, element, wait, 500, noSuchFoundExMsg);

            return element;

            object? WaitUntil(IWebElement iWebElement, IWebDriver driver, string searchName, 
                KindOfSearchElement searchElement, object? element, WebDriverWait wait, int waitTimeCountItems, string errorMsg)
            {
                int? oldCounterElements = 0, newCounterElements = 0;
                do
                {
                    oldCounterElements = newCounterElements;
                    try
                    {
                        wait.Until(d =>
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
                                    catch
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
                                    catch
                                    {
                                        return false;
                                    }
                                case KindOfSearchElement.FindElementByCssSelector:
                                    try
                                    {
                                        element = iWebElement.FindElement(By.CssSelector(searchName));

                                        return element != null;
                                    }
                                    catch
                                    {
                                        return false;
                                    }
                                case KindOfSearchElement.FindElementsByCssSelector:
                                    try
                                    {
                                        element = iWebElement.FindElements(By.CssSelector(searchName));

                                        return element != null;
                                    }
                                    catch
                                    {
                                        return false;
                                    }
                                case KindOfSearchElement.FindElementByClassName:
                                    try
                                    {
                                        element = iWebElement.FindElement(By.ClassName(searchName));

                                        return element != null;
                                    }
                                    catch
                                    {
                                        return false;
                                    }
                                case KindOfSearchElement.FindElementsByClassName:
                                    try
                                    {
                                        element = iWebElement.FindElements(By.ClassName(searchName));

                                        return element != null;
                                    }
                                    catch
                                    {
                                        return false;
                                    }
                                case KindOfSearchElement.FindElementByXPath:
                                    try
                                    {
                                        element = iWebElement.FindElement(By.XPath(searchName));

                                        return element != null;
                                    }
                                    catch
                                    {
                                        Console.WriteLine("");
                                        return false;
                                    }
                                case KindOfSearchElement.FindElementsByXPath:
                                    try
                                    {
                                        element = iWebElement.FindElements(By.XPath(searchName));

                                        return element != null;
                                    }
                                    catch
                                    {
                                        return false;
                                    }
                                case KindOfSearchElement.FindElementByTagName:
                                    try
                                    {
                                        element = iWebElement.FindElement(By.TagName(searchName));

                                        return element != null;
                                    }
                                    catch
                                    {
                                        return false;
                                    }
                                case KindOfSearchElement.FindElementsByTagName:
                                    try
                                    {
                                        element = iWebElement.FindElements(By.TagName(searchName));

                                        return element != null;
                                    }
                                    catch
                                    {
                                        return false;
                                    }
                                case KindOfSearchElement.FindElementByID:
                                    try
                                    {
                                        element = driver.FindElement(By.Id(searchName));

                                        return element != null;
                                    }
                                    catch
                                    {
                                        return false;
                                    }
                                case KindOfSearchElement.FindElementsByID:
                                    try
                                    {
                                        element = driver.FindElements(By.Id(searchName));

                                        return element != null;
                                    }
                                    catch
                                    {
                                        return false;
                                    }
                                default:
                                    return false;
                            }
                        });
                    }
                    //ReadOnlyCollection<IWebElement?> ?
                    catch (NoSuchElementException)
                    {
                        Console.WriteLine("Element not found.");
                        return null;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"{noSuchFoundExMsg} Error: {ex.Message}");
                        return null;
                    }
                    // Werden nach mehreren Elementen gesucht?
                    if (searchElement == KindOfSearchElement.FindElementsByCssSelector || searchElement == KindOfSearchElement.FindElementsByClassName ||
                        searchElement == KindOfSearchElement.FindElementsByTagName || searchElement == KindOfSearchElement.FindElementsByXPath)
                    {
                        ReadOnlyCollection<IWebElement?>? temp = null;
                        try
                        {
                            temp = (ReadOnlyCollection<IWebElement?>?)element;
                        }
                        catch (Exception ex) { Console.WriteLine("fehler: " + ex); }

                        newCounterElements = temp?.Count();
                        Thread.Sleep(waitTimeCountItems);
                    }
                }
                while (newCounterElements != oldCounterElements);
                return element;
            }
        }

        /// <summary>
        /// Versucht, ein bestimmtes Element zu finden und versucht es in gewissen Zeitabständen erneut, falls dieses Element nicht gefunden wird.
        /// </summary>
        /// <param name="shadowRoot"></param>
        /// <param name="driver"></param>
        /// <param name="searchName"></param>
        /// <param name="searchElement"></param>
        /// <param name="pollingIntervalTime"></param>
        /// <param name="maxWaitTime"></param>
        /// <returns></returns>
        internal static object? Searching(ShadowRoot shadowRoot, IWebDriver driver, string searchName, KindOfSearchElement searchElement, 
            int pollingIntervalTime = 500, int maxWaitTime = 10, string noSuchFoundExMsg = "")
        {
            object? element = null;
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTime))
            {
                PollingInterval = TimeSpan.FromMilliseconds(pollingIntervalTime),
            };
            wait.IgnoreExceptionTypes(typeof(ElementNotInteractableException));

            try
            {
                wait.Until(d =>
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
                                element = shadowRoot.FindElement(By.CssSelector(searchName));

                                return element != null;
                            }
                            catch (NoSuchElementException)
                            {
                                return false;
                            }
                        case KindOfSearchElement.FindElementsByCssSelector:
                            try
                            {
                                element = shadowRoot.FindElements(By.CssSelector(searchName));

                                return element != null;
                            }
                            catch (NoSuchElementException)
                            {
                                return false;
                            }
                        case KindOfSearchElement.FindElementByClassName:
                            try
                            {
                                element = shadowRoot.FindElement(By.ClassName(searchName));

                                return element != null;
                            }
                            catch (NoSuchElementException)
                            {
                                return false;
                            }
                        case KindOfSearchElement.FindElementsByClassName:
                            try
                            {
                                element = shadowRoot.FindElements(By.ClassName(searchName));

                                return element != null;
                            }
                            catch (NoSuchElementException)
                            {
                                return false;
                            }
                        case KindOfSearchElement.FindElementByXPath:
                            try
                            {
                                element = shadowRoot.FindElement(By.XPath(searchName));

                                return element != null;
                            }
                            catch (NoSuchElementException)
                            {
                                return false;
                            }
                        case KindOfSearchElement.FindElementsByXPath:
                            try
                            {
                                element = shadowRoot.FindElements(By.XPath(searchName));

                                return element != null;
                            }
                            catch (NoSuchElementException)
                            {
                                return false;
                            }
                        case KindOfSearchElement.FindElementByTagName:
                            try
                            {
                                element = shadowRoot.FindElement(By.TagName(searchName));

                                return element != null;
                            }
                            catch (NoSuchElementException)
                            {
                                return false;
                            }
                        case KindOfSearchElement.FindElementsByTagName:
                            try
                            {
                                element = shadowRoot.FindElements(By.TagName(searchName));

                                return element != null;
                            }
                            catch (NoSuchElementException)
                            {
                                return false;
                            }
                        case KindOfSearchElement.FindElementByID:
                            try
                            {
                                element = shadowRoot.FindElement(By.Id(searchName));

                                return element != null;
                            }
                            catch (NoSuchElementException)
                            {
                                return false;
                            }
                        case KindOfSearchElement.FindElementsByID:
                            try
                            {
                                element = shadowRoot.FindElements(By.Id(searchName));

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
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Element not found.");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{noSuchFoundExMsg} Error: {ex.Message}");
                return null;
            }
            return element;
        }

        internal static void SaveToExcel(List<Product> products, string period, string excelFilePath, MarketEnum discounter)
        {
            if (products == null)
            {
                Debug.WriteLine("Keine Daten zum speichern vorhanden.");
                return;
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(discounter.ToString());

            int columnCount = 7;                                  // Anzahl der Spalten

            var cellHeadline = worksheet.Cells[1, 1];
            worksheet.Cells["A1:G2"].Merge = true;                  // Bereich verbinden

            if (discounter == MarketEnum.Penny)
                cellHeadline.Value = $"Die Angebote vom Penny vom {period}";
            else if (discounter == MarketEnum.Lidl)
                cellHeadline.Value = $"Die Angebote vom Lidl {period}";
            else if (discounter == MarketEnum.Netto)
                cellHeadline.Value = $"Die Angebote vom Netto";
            else if (discounter == MarketEnum.Kaufland)
                cellHeadline.Value = $"Die Angebote vom Kaufland";
            else if (discounter == MarketEnum.AldiNord)
                cellHeadline.Value = $"Die Angebote vom Aldi Nord {period}";


            // Überschrift formatieren
            worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["A1"].Style.Font.Size = 14;
            cellHeadline.Style.Font.Color.SetColor(Color.Wheat);

            if (discounter == MarketEnum.Penny)
                HighLightCells(1, 1, 2, columnCount, worksheet, Color.Red, Color.LightPink);
            else if (discounter == MarketEnum.Lidl)
                HighLightCells(1, 1, 2, columnCount, worksheet, Color.Green, Color.LightGreen);
            else if (discounter == MarketEnum.Netto)
                HighLightCells(1, 1, 2, columnCount, worksheet, Color.Yellow, Color.LightYellow);
            else if (discounter == MarketEnum.Kaufland)
                HighLightCells(1, 1, 2, columnCount, worksheet, Color.LightPink, Color.DarkRed);
            else if (discounter == MarketEnum.AldiNord)
                HighLightCells(1, 1, 2, columnCount, worksheet, Color.SkyBlue, Color.DarkBlue);

            worksheet.Cells[3, 1].Value = "Name";
            worksheet.Cells[3, 2].Value = "Bezeichnung";
            worksheet.Cells[3, 3].Value = "Vorheriger Preis";
            worksheet.Cells[3, 4].Value = "Neuer Preis";
            worksheet.Cells[3, 5].Value = "Preis Pro Kg oder Liter";
            worksheet.Cells[3, 6].Value = "Info";
            worksheet.Cells[3, 7].Value = "Beginn";

            // Beschriftung formatieren
            HighLightCells(3, 1, 3, columnCount, worksheet, Color.Gray);

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
                if (products[i].PricePerKgOrLiter != 0)
                {
                    worksheet.Cells[i + offsetRow, 5].Value = products[i].PricePerKgOrLiter;
                    worksheet.Cells[i + offsetRow, 5].Style.Numberformat.Format = euroFormat;
                }
                worksheet.Cells[i + offsetRow, 6].Value = products[i].Badge;
                worksheet.Cells[i + offsetRow, 7].Value = products[i].OfferStartDate;

                if (i % 2 == 1)
                {
                    HighLightCells(i + offsetRow, 1, i + offsetRow, columnCount, worksheet, Color.LightGray);
                }

                // Überprüfe, ob eines der interessanten Produkten mit dabei ist. Falls ja, dann verändere die Hintergrundfarbe
                foreach (var interestingProduct in Program.FavoriteProducts)
                {
                    string produktFullName = products[i].Name;
                    if (IsContains(interestingProduct.Name, produktFullName))
                    {
                        if (products[i].PricePerKgOrLiter == 0)
                        {
                            if (products[i].NewPrice <= interestingProduct.PriceCapPerKg && products[i].NewPrice != 0)
                            {
                                HighLightCells(i + offsetRow, 1, i + offsetRow, columnCount, worksheet, Color.LightCoral);
                            }
                        }
                        else if (products[i].PricePerKgOrLiter <= interestingProduct.PriceCapPerKg && products[i].PricePerKgOrLiter != 0)      
                        {
                            HighLightCells(i + offsetRow, 1, i + offsetRow, columnCount, worksheet, Color.LightCoral);
                        }
                        else
                        {
                            HighLightCells(i + offsetRow, 1, i + offsetRow, columnCount, worksheet, Color.Yellow);
                        }
                    }
                }

                //static void HighLightCells(int row, int offsetRow, int columnCount, Color color, ExcelWorksheet worksheet)  

            }
            static void HighLightCells(int fromRow, int fromCol, int toRow, int toCol, ExcelWorksheet worksheet, Color backgroundColor, Color foregroundColor = default)
            {
                //var style = worksheet.Cells[row + offsetRow, 1, row + offsetRow, columnCount].Style;            // Bereich auswählen, welcher farblich geändert werden soll
                var style = worksheet.Cells[fromRow, fromCol, toRow, toCol].Style;            // Bereich auswählen, welcher farblich geändert werden soll
                style.Fill.PatternType = ExcelFillStyle.Solid;                                          // Bereich wird mit einer einheitlichen Farbe ohne Farbverlauf oder Muster eingefärbt
                style.Fill.BackgroundColor.SetColor(backgroundColor);
                style.Font.Color.SetColor(foregroundColor);
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

        internal static List<Product> LoadFromExcel(string excelFilePath, MarketEnum discounter)
        {
            List<Product> productsFromExcel = new List<Product>();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Öffnen der Excel-Datei
            FileInfo file = new FileInfo(excelFilePath);

            excelPackage = new ExcelPackage(file);
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[discounter.ToString()];

            if (worksheet == null)
            {
                Console.WriteLine("Excel Worksheet wurde nicht gefunden");
                return productsFromExcel;
            }

            // Iteriere durch die Zellen und füge den Wert der Liste hinzu
            for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
            {
                string name = "", description = "", badge = "", startDate = "";
                decimal? oldPrice = 0, newPrice = 0, pricePerKgOrLiter = 0;
                
                try
                {
                    var nameValue = (string)worksheet.Cells[row, 1].Value;
                    if (nameValue != null) { name = nameValue.ToString(); }
                    var descriptionValue = (string)worksheet.Cells[row, 2].Value;
                    if (descriptionValue != null) { description = descriptionValue.ToString(); }
                    var oldPriceValue = worksheet.Cells[row, 3].Value;
                    if (oldPriceValue != null) { oldPrice = decimal.Parse(oldPriceValue.ToString()); }
                    var newPriceValue = worksheet.Cells[row, 4].Value;
                    if (newPriceValue != null) { newPrice = decimal.Parse(newPriceValue.ToString()); }
                    var pricePerKgOrLiterValue = worksheet.Cells[row, 5].Value;
                    if (pricePerKgOrLiterValue != null) { pricePerKgOrLiter = decimal.Parse(pricePerKgOrLiterValue.ToString()); }
                    var badgeValue = worksheet.Cells[row, 6].Value;
                    if (badgeValue != null) { badge = badgeValue.ToString(); }
                    var startDateValue = (string)worksheet.Cells[row, 7].Value;
                    if (startDateValue != null) { startDate = startDateValue.ToString(); }

                    productsFromExcel.Add(new Product(name, description, oldPrice, newPrice, pricePerKgOrLiter, badge, startDate));
                }
                catch { Console.WriteLine("Fehler beim umwandeln von den Werten von der Excel Tabelle."); }
            }

            return productsFromExcel;
        }

        /// <summary>
        /// Lade die alte Angebots Datei und extrahiere die Überschrift (enthält von wann bis wann die Angebote gültig sind).
        /// </summary>
        /// <param name="excelPath"></param>
        /// <returns></returns>
        internal static string ExtractHeadlineFromExcel(string excelPath, MarketEnum marketEnum)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Excel-Paket erstellen und die Excel-Datei laden
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(excelPath)))
            {
                if (!File.Exists(excelPath))
                    return string.Empty;

                // Das erste Arbeitsblatt auswählen (Index beginnt bei 0). Es kann aber auch der Name vom Arbeitsblatt angegeben werden
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[marketEnum.ToString()];

                if (worksheet != null)
                    return (string)worksheet.Cells["A1"].Value;
                else
                    return string.Empty;
            }
        }

        /// <summary>
        /// Sendet eine E-Mail, falls folgende Kriterien erfüllt werden:
        /// 1. Es müssen neue Angebote vorhanden sein.
        /// 2. Es muss in den Sonderangeboten min. ein Produkte vorhanden sein, welches in der Auswahl mit aufgenommen wurde.
        /// 3. Die gefundenen Produkte bzw. das gefundene Produkt muss den angegebenen Preis unterschritten haben oder gleich hoch sein.
        /// </summary>
        /// <param name="mailAdress"></param>
        /// <param name="subject"></param>
        /// <param name="body"></param>
        

        /// <summary>
        /// Informiert jedesmal, wenn neue Angebote verfügbar sind, per E-Mail, wenn diese einen bestimmten festgelegten Preis unterstreiten.
        /// </summary>
        /// <param name="isNewOffersAvailable"></param>
        /// <param name="allProducts"></param>
        internal static void InformPerEMail(bool isNewOffersAvailable, Dictionary<MarketEnum, List<Product>> allProducts)
        {
            if (isNewOffersAvailable)
            {
                int interestingOfferCounter = 0;
                string offers = string.Empty;
                List<string> marketsWithInterestingOffers = new();
                // Als nächstes soll untersucht werden, ob von den interessanten Angeboten der Preis auch niedrig genug ist.
                foreach (var products in allProducts)
                {
                    foreach (var product in products.Value)
                    {
                        foreach (var interestingProduct in Program.FavoriteProducts)
                        { 
                            if (product.Name.ToLower().Trim().Contains(interestingProduct.Name.ToLower().Trim()))
                            {
                                string marketName = products.Key.ToString();
                                // falls bei beiden Kg bzw. Liter Preise der Wert "0" drin steht, dann könnte das bedeuten,
                                // dass entweder die Menge schon ein Kilo entspricht oder dass sich der Preis pro Produkt bezieht.
                                if (product.PricePerKgOrLiter == 0)
                                {
                                    if (product.NewPrice <= interestingProduct.PriceCapPerProduct && product.NewPrice != 0)
                                    {
                                        offers += $" {product.Name} {product.Description} vom Anbieter {marketName} {product.OfferStartDate}" +
                                            $" für nur {product.NewPrice}€.\n";
                                        interestingOfferCounter++;

                                        if (!marketsWithInterestingOffers.Contains(marketName))
                                        {
                                            marketsWithInterestingOffers.Add(marketName);
                                        }
                                    }
                                }
                                else if (product.PricePerKgOrLiter <= interestingProduct.PriceCapPerKg && product.PricePerKgOrLiter != 0)
                                {
                                    offers += $" {product.Name} {product.Description} vom Anbieter {marketName} {product.OfferStartDate}" +
                                        $" für nur {product.NewPrice}€ ({product.PricePerKgOrLiter}€ pro Kg).\n";
                                    interestingOfferCounter++;

                                    if (!marketsWithInterestingOffers.Contains(marketName))
                                    {
                                        marketsWithInterestingOffers.Add(marketName);
                                    }
                                }
                            }
                        }
                    }
                    offers += "\n";
                }

                string body = string.Empty;
                string subject = string.Empty;
                if (interestingOfferCounter > 1)
                {
                    subject = "Interessante Angebote gefunden!";
                    body = $"Gute Nachricht! Folgende Angebote, die dich interessieren und günstig genug sind, wurden gefunden: \n\n\n{offers}";
                }
                else if (interestingOfferCounter == 1)
                {
                    subject = "Es wurde ein interessantes Angebot gefunden!";
                    body = $"Gute Nachricht! Folgendes Angebot, welches dich interessiert und günstig genug ist, wurde gefunden: \n\n\n{offers}";
                }

                if (interestingOfferCounter > 0)
                {
                    if (marketsWithInterestingOffers.Count > 1)
                        body += "\nHier sind die Links zu den Anbietern: \n\n";
                    else
                        body += "\nHier ist der Link zum Anbieter: \n\n";

                    foreach (var market in marketsWithInterestingOffers)
                    {
                        if (market.ToLower() == "penny")
                            body += "https://www.penny.de/angebote \n";
                        if (market.ToLower() == "lidl")
                            body += "https://www.lidl.de/store \n";
                        if (market.ToLower() == "aldinord")
                            body += "https://www.aldi-nord.de/angebote.html \n";
                    }
                    
                    body += "\nViel Spaß beim Einkaufen!";
                    //"https://www.penny.de/angebote \n" +
                    //"https://www.lidl.de/store \n\n" +
                    //"Viel Spaß beim Einkaufen!";

                    SendEMail(Program.Email, subject, body);
                }
            }
            void SendEMail(string mailAdress, string subject, string body)
            {
                // E-Mail-Einstellungen
                string senderEmail = "****";
                string receiverEmail = mailAdress;

                // SMTP-Server-Einstellungen
                string smtpServer = "mail.gmx.net";
                int smtpPort = 587; // Standard-Port für SMTP ist 587
                string smtpUsername = "****";
                string smtpPassword = "****";

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
    }
}
