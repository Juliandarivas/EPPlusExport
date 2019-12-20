using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlusExport.Models
{
    public class Customer
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string Country { get; set; }

        public Customer(int id, string name, string email, string country)
        {
            Id = id;
            Name = name;
            Email = email;
            Country = country;
        }
    }
}
