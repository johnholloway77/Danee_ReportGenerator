using ReportGenerator.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Navigation;

namespace ReportGenerator.Services
{
    internal class Findings_Calculator
    {
        public static List<Finding> calculator(List<Finding> raw_Findings)
        {

            List<Finding> findings = raw_Findings;

            DateTime today = DateTime.Today;

            foreach (Finding finding in findings)
            {
                //duedate is earlier than today. Past due
                if (DateTime.Compare(finding.due_Date, today) < 0)
                {
                    finding.late_status = 2; //if past due equals 2 
                }
                else if (DateTime.Compare(finding.due_Date, today) == 0)
                {
                    finding.late_status = 1; //finding is due today.
                }
                else if (DateTime.Compare(finding.due_Date, today) > 0)
                {
                    finding.late_status = 0; //finding not due.
                }

                TimeSpan days = today - finding.due_Date;

                finding.days_Overdue = (int)days.TotalDays;

                days = today - finding.start_Date;

                finding.days_From_entered = (int)days.TotalDays;


            }

            return findings;

        }

        public static List<string> unique(List<Finding> raw_Findings)
        {
            List<string> departments = new List<string>();

            foreach (Finding finding in raw_Findings)
            {
                departments.Add(finding.department);
            }
            departments = departments.Distinct().ToList();

            foreach (String _ in departments) {
                Console.WriteLine(_);
            }
            return departments;
        }

        public static List<Department> department_calculator(List<Finding> findings, List<string> departments)
        {
           // List<Tuple<string, int, int, int>> depart_List = new List<Tuple<string, int, int, int>>();
           List<Department> depart_list = new List<Department>();

            foreach(String department in departments)
            {
                int open = 0;
                int due = 0;
                int pastDue = 0;

                foreach(Finding finding in findings)
                {
                    if(finding.department == department)
                    {
                        switch(finding.late_status)
                        {
                            case 0:
                                open++;
                                break;
                            case 1: 
                                due++; 
                                break;                            
                            case 2:
                                pastDue++;
                                break;
                        }
                    }
                }

                
                depart_list.Add(new Department(department, open, due, pastDue));
            }

            return depart_list;
        }
    
    }

  
  
}
