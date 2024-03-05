using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections;

namespace LookForSpecialOffers
{
    class Program
    {
        static void Main(string[] args) 
        {
            using (IWebDriver driver = new ChromeDriver())
            {
                driver.Navigate().GoToUrl("https://www.penny.de/");
                string searchName = "//div[contains(@class, 'site-header__wrapper')]";      //Suche nach dem Element, wo alle links von der Kopfzeile vorhanden sind
                var headerWrapper = FindObject(driver, searchName, KindOfSearchElement.SelectSingleNode, 100, 50);

            }
        }







        private static object FindObject(IWebDriver driver, string name, KindOfSearchElement searchElement, int interval = 500, int maxCount = 20)
        {
            HtmlDocument doc = new HtmlDocument();
            int count = 0;
            object element = null;
            while (count < maxCount)
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

                count++;
                Thread.Sleep(interval);
            }
            return null;
        }
    }
}








