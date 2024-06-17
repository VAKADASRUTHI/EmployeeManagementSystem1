using Newtonsoft.Json;
using System;

namespace EmployeeManagementSystem.Models
{
   public class Address
   {
        [JsonProperty(PropertyName = "street", NullValueHandling = NullValueHandling.Ignore)]
         public string Street { get; set; }

        [JsonProperty(PropertyName = "city", NullValueHandling = NullValueHandling.Ignore)]
           public string City { get; set; }

        [JsonProperty(PropertyName = "state", NullValueHandling = NullValueHandling.Ignore)]
          public string State { get; set; }

        [JsonProperty(PropertyName = "postalCode", NullValueHandling = NullValueHandling.Ignore)]
           public string PostalCode { get; set; }

        [JsonProperty(PropertyName = "country", NullValueHandling = NullValueHandling.Ignore)]
          public string Country { get; set; }
   }
}