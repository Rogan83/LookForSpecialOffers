using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LookForSpecialOffers.Models
{
    internal class Market
    {
        public string Name { get; set; }
        public bool IsSelected { get; set; }

        public Market(string name, bool isSelected)
        {
            Name = name;
            IsSelected = isSelected;
        }
    }
}
