using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LookForSpecialOffers
{
    internal class ProduktFavorite
    {
        public string Name { get; set; }
        public double PriceCap { get; set; }

        public ProduktFavorite(string name, double pricePerKg)
        {
            Name = name;
            PriceCap = pricePerKg;
        }
    }
}
