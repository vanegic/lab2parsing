using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace lab2
{
    class Threat 
    {
        public static List<Threat> threats = new List<Threat>();
        public int Id { get; set; }
        public string Name { get; set; }
        public string Notice { get; set; }
        public string Source { get; set; }
        public string Influence { get; set; }
        public bool ConfidentityThreat { get; set; } = false;
        public bool IntegrityThreat { get; set; } = false;
        public bool AccessThreat { get; set; } = false;
        public string CreationDate { get; set; }
        public string ChangeDate { get; set; }
        public string Changes { get; set; } = " ";
    }
}
