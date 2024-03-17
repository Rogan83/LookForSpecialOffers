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
        FindElementsByClassName,
        FindElementByXPath,
        FindElementsByXPath,
        FindElementByID,
        FindElementsByID,
        FindElementByTagName,
        FindElementsByTagName,
        SelectSingleNode,
        SelectNodes
    }
}
