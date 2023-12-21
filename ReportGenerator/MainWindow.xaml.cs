using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ReportGenerator.services;
using ReportGenerator.Models;
using ReportGenerator.Services;
using OfficeOpenXml;
using System.Runtime.CompilerServices;

namespace ReportGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<Finding> findings;
        List<string> department_Names;
        List<Department> departments;
        //public ExcelPackage analysisExcel;
        private enum status { open, due, pastdue, closed };
        
        
        //public string text = "Please load an excel file";

        public MainWindow()
        {
            InitializeComponent();
            
            MyTextBlock.Text = "Please load an excel file";
            //this.text2 = text;
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel documents (.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                string fileLocation = openFileDialog.FileName;

               if (fileLocation != null)
                {
                    findings = Excel_Manager.loader(fileLocation);
                }
            }

            if (findings != null)
            {
                Console.WriteLine("Findings populated");
                findings = Findings_Calculator.calculator(findings);

                department_Names = Findings_Calculator.unique(findings);

                departments =  Findings_Calculator.department_calculator(findings, department_Names);


                MyTextBlock.Text = "Excel successfully loaded and processed. Please save file.";

                /*
                foreach (Department _ in departments)
                {
                    Console.WriteLine($"Department: {_.name} \n\tTotal Open: {_.open} \n\tTotal Due: {_.due} \n\tTotal Overdue: {_.pastDue}\n");
                }


                foreach (Finding _ in findings)
                {
                    Console.WriteLine($"Number: {_.id_Number} \nTitle: {_.title} \n\tStartdate: {_.start_Date.ToString()} \n\tDueDate: {_.due_Date.ToString()}  \n\tDays Until Due: {_.days_Overdue} \n\tDays from Entered: {_.days_From_entered} \n\tStatus: {(status)_.late_status}\n");
                }
               */
            }
            else
            {
                MyTextBlock.Text = "Unable to load an analyse excel file.";
            }
        }



        private void btnSaveFile_Click(object sender, RoutedEventArgs e)
        {
            int i = 0;

            if (findings != null)
            {


                Console.WriteLine("Save clicked");
                string save_Location;

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";

                if(saveFileDialog.ShowDialog() == true)
                {
                    save_Location = saveFileDialog.FileName;
                    Console.WriteLine($"{save_Location}");

                    i = Excel_Manager.builder(findings, departments, save_Location);
                    Console.WriteLine($"File saved to {save_Location}");

                    switch (i)
                    {
                        case 0:
                            MyTextBlock.Text = $"File has been saved to {save_Location}";
                            break;
                        case 1:
                            MyTextBlock.Text = $"Unable to save overwrite existing file. Please Ensure you do not have the file open.";
                            break;
                        case 2:
                            MyTextBlock.Text = $"Unable to save file";
                            break;
                        default:
                            MyTextBlock.Text = "File written";
                            break;

                    }
                    
                }

                

            }
            else
            {
                Console.WriteLine("No file loaded!");
                MyTextBlock.Text = "No file loaded! Please load excel file before saving";
            }
        }

        
    }
}
