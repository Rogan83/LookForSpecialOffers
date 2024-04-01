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

        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType()) return false;

            Product other = (Product)obj;

            return Name == other.Name && Description == other.Description && OldPrice == other.OldPrice &&
                   NewPrice == other.NewPrice && PricePerKgOrLiter == other.PricePerKgOrLiter &&
                   Badge == other.Badge && OfferStartDate == other.OfferStartDate;
        }

        public override int GetHashCode()
        {
            unchecked // Overflow ist in diesem Fall beabsichtigt
            {
                int hash = 17; // Eine beliebige Startzahl

                // Multiplizieren und Addieren von Hashcodes der Felder
                hash = hash * 23 + (Name?.GetHashCode() ?? 0);
                hash = hash * 23 + (Description?.GetHashCode() ?? 0);
                hash = hash * 23 + OldPrice.GetHashCode();
                hash = hash * 23 + NewPrice.GetHashCode();
                hash = hash * 23 + PricePerKgOrLiter.GetHashCode();
                hash = hash * 23 + (Badge?.GetHashCode() ?? 0);
                hash = hash * 23 + (OfferStartDate?.GetHashCode() ?? 0);

                return hash;
            }
        }
    }
}
