using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailCompleteApp.Models
{
    public class Location
    {
        public string firmName {  get; set; }
        public string address { get; set; }
        public Location(string firmName, string address ) 
        {
            this.firmName = firmName;
            this.address = address;
        }
    }
}
