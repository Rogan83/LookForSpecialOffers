using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LookForSpecialOffers
{
    internal enum KindOfSearchElement
    {
        FindElementByCssSelector,
        FindElementsByCssSelector,
        FindElementByClassName,
        FindElementByXPath,
        SelectSingleNode,
        SelectNodes
    }
}
