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
        public double PricePerKgOrLiter1 { get; set; }
        public double PricePerKgOrLiter2 { get; set; }
        public string Badge { get; set; }
        public string OfferStartDate { get; set; }

        public Product(string name, string description, double oldPrice, double newPrice, double pricePerKgOrLiter1, double pricePerKgOrLiter2, string badge, string offerStartDate)
        {
            Name = name;
            Description = description;
            OldPrice = oldPrice;
            NewPrice = newPrice;
            PricePerKgOrLiter1 = pricePerKgOrLiter1;
            PricePerKgOrLiter2 = pricePerKgOrLiter2;
            Badge = badge;
            OfferStartDate = offerStartDate;
        }
    }
}
