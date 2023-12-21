using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.Models
{
    internal class Finding
    {
        

        public Finding(string id_Number, string title, string owner, string department, DateTime start_Date, DateTime due_Date)
        {
            this.id_Number = id_Number;
            this.title = title;
            this.owner = owner;
            this.department = department;
            this.start_Date = start_Date;
            this.due_Date = due_Date;


        }

        public string id_Number { get; }
        public string title { get;  }
        public string owner { get; set; }
        public string department { get; set; }
        public int days_Overdue { get; set; }
        public int days_From_entered { get; set; }

        public DateTime start_Date { get;  }
        public DateTime due_Date { get;  }
        public int late_status { get; set; }

        //
    }
}
