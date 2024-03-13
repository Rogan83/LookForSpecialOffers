﻿using HtmlAgilityPack;
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

namespace LookForSpecialOffers
{
    internal static class WebScraperHelper
    {
        internal static ExcelPackage excelPackage = new ExcelPackage();

        /// <summary>
        /// Überprüft, ob ein Element interagierbar ist
        /// </summary>
        /// <param name="cookieAcceptBtn"></param>
        internal static bool CheckIfInteractable(IWebElement cookieAcceptBtn)
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
        internal static void ScrollToBottom(IWebDriver driver, int delayPerStep = 200, int stepsCount = 10, int delayDetermineScrollHeigth = 500)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            long oldScrollHeight = (long)js.ExecuteScript("return document.body.scrollHeight;");
            long newScrollHeight = 0;
            // Als erstes muss die Höhe der Webseite ermittelt werden. Diese verändert sich, während die Webseite Inhalt nach lädt über die Zeit.
            // Es dauert eine Weile, bis die Scrollheight ermittelt wird. Deswegen wird die schleife so lange wiederholt, bis sind die Scrollheight
            // nicht mehr verändert, was bedeuten müsste, dass diese den entgültigen wert ermittelt hat
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

            //Nachdem die Höhe der Seite ermittelt wurde, soll nun Stufenweise die Seite herunter gescrollt werden, damit der Inhalt Stück für Stück
            //von der Seite nachgeladen werden kann.
            long offset = oldScrollHeight / stepsCount;
            long newPos = 0;

            for (int i = 0; i < stepsCount; i++)
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

        internal static bool IsContains(string substring, string product)
        {
            substring = substring.Trim().ToLower();
            product = product.Trim().ToLower();

            return product.Contains(substring);
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

        internal static object? Find(IWebDriver driver, string searchName, KindOfSearchElement searchElement, int pollingIntervalTime = 500, int maxWaitTime = 10)
        {
            object? element = null;
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
                    case KindOfSearchElement.FindElementByClassName:
                        try
                        {
                            element = driver.FindElement(By.ClassName(searchName));

                            return element != null;
                        }
                        catch (NoSuchElementException)
                        {
                            return false;
                        }
                    case KindOfSearchElement.FindElementByXPath:
                        try
                        {
                            element = driver.FindElement(By.XPath(searchName));

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

        internal static object? Find(ShadowRoot shadowRoot, IWebDriver driver, string searchName, KindOfSearchElement searchElement, int pollingIntervalTime = 500, int maxWaitTime = 10)
        {
            object? element = null;
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTime))
            {
                PollingInterval = TimeSpan.FromMilliseconds(pollingIntervalTime),
            };
            wait.IgnoreExceptionTypes(typeof(ElementNotInteractableException));


            var main = wait.Until(d =>
            {
                switch (searchElement)
                {
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
                    default:
                        return false;
                }


            });

            return element;
        }

        internal static void SaveToExcel(List<Product> products, string period, string path, Discounter discounter)
        {
            if (products == null)
            {
                Debug.WriteLine("Keine Daten zum speichern vorhanden.");
                return;
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string excelFilePath = path;

            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(discounter.ToString());

            int columnCount = 8;                                  // Anzahl der Spalten

            var cellHeadline = worksheet.Cells[1, 1];
            worksheet.Cells["A1:H2"].Merge = true;                  // Bereich verbinden

            if (discounter == Discounter.Penny)
                cellHeadline.Value = $"Die Angebote vom Penny vom {period}";
            else if (discounter == Discounter.Lidl)
                cellHeadline.Value = $"Die Angebote vom Lidl";
            else if (discounter == Discounter.Netto)
                cellHeadline.Value = $"Die Angebote vom Netto";
            else if (discounter == Discounter.Kaufland)
                cellHeadline.Value = $"Die Angebote vom Kaufland";
            else if (discounter == Discounter.Aldi)
                cellHeadline.Value = $"Die Angebote vom Aldi";


            // Überschrift formatieren
            worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheet.Cells["A1"].Style.Font.Size = 14;
            cellHeadline.Style.Font.Color.SetColor(Color.Wheat);

            if (discounter == Discounter.Penny)
                HighLightCells(1, 1, 2, columnCount, Color.Red, worksheet);
            else if (discounter == Discounter.Lidl)
                HighLightCells(1, 1, 2, columnCount, Color.Green, worksheet);
            else if (discounter == Discounter.Netto)
                HighLightCells(1, 1, 2, columnCount, Color.Yellow, worksheet);
            else if (discounter == Discounter.Kaufland)
                HighLightCells(1, 1, 2, columnCount, Color.LightPink, worksheet);
            else if (discounter == Discounter.Aldi)
                HighLightCells(1, 1, 2, columnCount, Color.LightSkyBlue, worksheet);

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
                foreach (var interestingProduct in Program.InterestingProducts)
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

        /// <summary>
        /// Lade die alte Angebots Datei und extrahiere die Überschrift (enthält von wann bis wann die Angebote gültig sind).
        /// </summary>
        /// <param name="excelPath"></param>
        /// <returns></returns>
        internal static string ExtractHeadlineFromExcel(string excelPath)
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

        /// <summary>
        /// Sendet eine E-Mail, falls folgende Kriterien erfüllt werden:
        /// 1. Es müssen neue Angebote vorhanden sein.
        /// 2. Es muss in den Sonderangeboten min. ein Produkte vorhanden sein, welches in der Auswahl mit aufgenommen wurde.
        /// 3. Die gefundenen Produkte bzw. das gefundene Produkt muss den angegebenen Preis unterschritten haben oder gleich hoch sein.
        /// </summary>
        /// <param name="mailAdress"></param>
        /// <param name="subject"></param>
        /// <param name="body"></param>
        internal static void SendEMail(string mailAdress, string subject, string body)
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
}
