using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LookForSpecialOffers.Models
{
    internal class DetailPage
    {
        internal string Url { get; private set; }
        internal string Info { get; private set; }

        internal DetailPage(string url, string info)
        {
            Url = url;
            Info = info;
        }
    }
}
