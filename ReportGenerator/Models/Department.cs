using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.Models
{
    internal class Department
    {

        public Department(string name, int open, int due, int pastDue)
        {
            this.name = name;
            this.open = open;
            this.due = due;
            this.pastDue = pastDue;
        }
    
        public string name { get; set; }

        public int open { get; set; }
        public int due { get; set; }
        public int pastDue {  get; set; }

    }
}
