using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LookForSpecialOffers.Enums
{
    internal enum KindOfSearchElement
    {
        FindElementByCssSelector,
        FindElementsByCssSelector,
        FindElementByClassName,
        FindElementByXPath,
        FindElementByID,
        SelectSingleNode,
        SelectNodes
    }
}
