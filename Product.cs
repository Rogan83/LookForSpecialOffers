using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LookForSpecialOffers
{
    internal class Product
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public double OldPrice { get; set; }
        public double NewPrice { get; set; }
        public double PricePerKgOrLiter { get; set; }
        public string Badge { get; set; }
        public string OfferStartDate { get; set; }

        public Product(string name, string description, double oldPrice, double newPrice, double pricePerKgOrLiter, string badge, string offerStartDate)
        {
            Name = name;
            Description = description;
            OldPrice = oldPrice;
            NewPrice = newPrice;
            PricePerKgOrLiter = pricePerKgOrLiter;
            Badge = badge;
            OfferStartDate = offerStartDate;
        }
    }
}
