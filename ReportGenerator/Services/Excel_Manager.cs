using OfficeOpenXml;
using OfficeOpenXml.Style;
using ReportGenerator.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Drawing;
using System.Windows.Media;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Drawing.Chart;

namespace ReportGenerator.services
{
    internal class Excel_Manager
    {
        private enum status { Open, Due, PastDue, Closed };


        public static List<Finding> loader(string address)
        {
            List<Finding> findings = new List<Finding>();
            //DateTime today = DateTime.Today;

            try
            {

                using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(@address)))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    var my_Worksheet = package.Workbook.Worksheets.First(); //use default worksheet
                    var totalRows = my_Worksheet.Dimension.End.Row;
                    var totalColumns = my_Worksheet.Dimension.End.Column;




                    for (int i = 4; i < totalRows; i++)
                    {
                        Finding _ = new Finding(
                            my_Worksheet.Cells[$"C{i}"].Value?.ToString().Trim(), my_Worksheet.Cells[$"D{i}"].Value?.ToString().Trim(), my_Worksheet.Cells[$"P{i}"].Value?.ToString().Trim(), my_Worksheet.Cells[$"U{i}"].Value?.ToString().Trim(), DateTime.Parse(my_Worksheet.Cells[$"B{i}"].Value?.ToString().Trim()), DateTime.Parse(my_Worksheet.Cells[$"L{i}"].Value?.ToString().Trim())
                            );


                        findings.Add(_);


                    }


                }
                return findings;
            }
            catch 
            {
                Console.WriteLine("Unable to parse data");
                
                return null;
            }
            
            
        }


        public static int builder(List<Finding> findings, List<Department> departments, string save_location)
        {
            using (ExcelPackage package = new ExcelPackage(@save_location))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                try
                {
                    package.Workbook.Worksheets.Delete($"Analysis {DateTime.Now.ToShortDateString()}");
                    Console.WriteLine($"Deleted Worksheet: Analysis {DateTime.Now.ToShortDateString()}");
                }
                catch
                {
                    Console.WriteLine("No Worksheet to delete");
                }

                var sheet = package.Workbook.Worksheets.Add($"Analysis {DateTime.Now.ToShortDateString()}");

                //ExcelChart chart = sheet.Drawings.AddChart("Chart", eChartType.ColumnClustered);

                int row_Number = 0; 
                
                //headers and all that jazz
                sheet.Cells["A1"].Value = $"Report generated on {DateTime.Now}";

                sheet.Cells["A2"].Value = "Date Entered";
                sheet.Cells["B2"].Value = "Due Date";
                sheet.Cells["C2"].Value = "Findings Number";
                sheet.Cells["D2"].Value = "Findings Title";
                sheet.Cells["E2"].Value = "Days Overdue:";
                sheet.Cells["F2"].Value = "Days From Entered";
                sheet.Cells["G2"].Value = "Status";
                sheet.Cells["H2"].Value = "Department";
                sheet.Cells["I2"].Value = "Owner";
                
                //style for headers.
                sheet.Row(1).Style.Font.Bold = true;
                sheet.Row(1).Height = 20;
                sheet.Row(2).Style.Font.Bold = true;
                sheet.Row(2).Height = 20;

                sheet.Column(1).Width = 13;
                sheet.Column(2).Width = 13;
                sheet.Column(3).Width = 15;
                sheet.Column(4).Width = 60;
                sheet.Column(5).Width = 15;
                sheet.Column(6).Width = 15;
                sheet.Column(7).Width = 11;
                sheet.Column(8).Width = 15;
                sheet.Column(9).Width = 15;

                row_Number = 3;
                for(int i = 0; i < findings.Count; i++) 
                {
                    sheet.Cells[$"A{3 + i}"].Value = findings[i].start_Date.ToShortDateString();
                    sheet.Cells[$"B{3 + i}"].Value = findings[i].due_Date.ToShortDateString();
                    sheet.Cells[$"C{3 + i}"].Value = findings[i].id_Number;
                    sheet.Cells[$"D{3 + i}"].Value = findings[i].title;
                    sheet.Cells[$"E{3 + i}"].Value = findings[i].days_Overdue;
                    sheet.Cells[$"F{3 + i}"].Value = findings[i].days_From_entered;
                    sheet.Cells[$"G{3 + i}"].Value = (status)findings[i].late_status;
                    sheet.Cells[$"H{3 + i}"].Value = findings[i].department;
                    sheet.Cells[$"I{3 + i}"].Value = findings[i].owner;

                    sheet.Cells[$"G{3 + i}"].Style.Font.Bold = true;
                    sheet.Cells[$"G{3 + i}"].Style.Font.Color.SetColor(System.Drawing.Color.White);
                    sheet.Cells[$"G{3 + i}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[$"G{3 + i}"].Style.Fill.PatternType = ExcelFillStyle.Solid;

                    switch (findings[i].late_status)
                    {
                        case 0:
                            sheet.Cells[$"G{3 + i}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);
                            break;

                        case 1:
                            sheet.Cells[$"G{3 + i}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            break;
                        case 2:
                            sheet.Cells[$"G{3 + i}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                            break;
                        default:
                            break;

                    }

                    row_Number++;

                }

                row_Number++;

                sheet.Cells[$"A{row_Number}"].Value = "Department";
                sheet.Cells[$"B{row_Number}"].Value = "Open";
                sheet.Cells[$"C{row_Number}"].Value = "Due";
                sheet.Cells[$"D{row_Number}"].Value = "Past Due";

                sheet.Row(row_Number).Style.Font.Bold = true;
                sheet.Row(row_Number).Height = 20;

                row_Number++;


                foreach(Department department in departments)
                {
                    sheet.Cells[$"A{row_Number}"].Value = department.name;
                    sheet.Cells[$"B{row_Number}"].Value = department.open;
                    sheet.Cells[$"C{row_Number}"].Value = department.due;
                    sheet.Cells[$"D{row_Number}"].Value = department.pastDue;
                    row_Number++;
                }

               // var pivotTable = package.Workbook.Worksheets[$"Analysis {DateTime.Now.ToShortDateString()}"].PivotTables[0];
               


                try
                {
                    //create file at the specified location
                    if (File.Exists(save_location))
                    {
                        try
                        {
                            File.Delete(save_location);
                        }
                        catch
                        {

                            return 1; // return error: file open can't overwrite
                        }

                    }


                    FileStream fs = File.Create(save_location);
                    fs.Close();

                    File.WriteAllBytes(save_location, package.GetAsByteArray());
                    package.Dispose();

                    return 0; //return successfully
                }
                catch {
                    return 2; // return error
                }


            }

            
        }

    }
}
