using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmsService
{
   public  class VehicleDetail
    {
        public int VehicalId { get; set; }
        public DateTime RenewalDate { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string PolicyNumber { get; set; }
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string MakeDescription { get; set; }
        public string ModelDescription { get; set; }
        public string UserId { get; set; }
        public string Email { get; set; }
        public string PhoneNumber { get; set; }

        public string CountryCode { get; set; }

        public bool IsCustomEmail { get; set; }
    }
}
