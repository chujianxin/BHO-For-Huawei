using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace COM.HPE
{
    public class AW
    {
        public string comment { get; set; }
        public string name { get; set; }
        public List<Record> record { get; set; }
       
    }
    public class Record
    {

        public List<Element> elements { get; set; }
        public string fullUrl { get; set; }
        public string map { get; set; }
        public string page { get; set; }
        public string url { get; set; }
        public string website { get; set; }
        public string iframesrc { get; set; }
         
    }
    public class Element
    {
        public string id { get; set; }
        public string _class { get; set; } 
        public string action { get; set; }
        public string comment { get; set; }
        public string name { get; set; }
        public string value { get; set; }
        public string xpath { get; set; }
        public string type { get; set; }
        public string noattrxpath { get; set; }
        public string csspath { get; set; } 
    }
     
}
