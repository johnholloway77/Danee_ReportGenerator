using OfficeOpenXml;
using OfficeOpenXml.Style;
using ReportGenerator.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml;
using OfficeOpenXml.Table;

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

            int findingsCount = findings.Count;




            //Lets export the data as a pivot table
            using (ExcelPackage package = new ExcelPackage(@save_location))
            {
                
                //Try-catch block to delete an existing sheet if overwriting existing file
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
                

                int row_Number = 0;
                var sheet = package.Workbook.Worksheets.Add($"Analysis {DateTime.Now.ToShortDateString()}");

                //headers and all that jazz
                sheet.Cells["D1"].Value = $"Report generated on {DateTime.Now}";

                using (ExcelRange Rng = sheet.Cells[$"A3:I{findingsCount + 3}"])
                {
                    ExcelTable table = sheet.Tables.Add(Rng, "tblFindings");

                    //Set column positions and names
                    table.Columns[0].Name = "Date Entered";
                    table.Columns[1].Name = "Findings Number";
                    table.Columns[2].Name = "ID Number";
                    table.Columns[3].Name = "Findings Title";
                    table.Columns[4].Name = "Days Overdue:";
                    table.Columns[5].Name = "Days from Entered";
                    table.Columns[6].Name = "Status";
                    table.Columns[7].Name = "Department";
                    table.Columns[8].Name = "Owner";


                    //show the filter thingimajig
                    table.ShowHeader = true;
                    table.ShowFilter = true;
                   
                }


                row_Number = 4;

                for (int i = 0 ; i < findings.Count; i++)
                {
                    using(ExcelRange Rng = sheet.Cells[$"A{i + 4}"])
                    {
                        Rng.Value = findings[i].start_Date.ToShortDateString();
                    }
                    using (ExcelRange Rng = sheet.Cells[$"B{i + 4}"])
                    {
                        Rng.Value = findings[i].due_Date.ToShortDateString();
                    }
                    using (ExcelRange Rng = sheet.Cells[$"C{i + 4}"])
                    {
                        Rng.Value = findings[i].id_Number;
                    }
                    using (ExcelRange Rng = sheet.Cells[$"D{i + 4}"])
                    {
                        Rng.Value = findings[i].title;
                    }
                    using (ExcelRange Rng = sheet.Cells[$"E{i + 4}"])
                    {
                        Rng.Value = findings[i].days_Overdue;
                    }
                    using (ExcelRange Rng = sheet.Cells[$"F{i + 4}"])
                    {
                        Rng.Value = findings[i].days_From_entered;
                    }
                    using (ExcelRange Rng = sheet.Cells[$"G{i + 4}"])
                    {
                        Rng.Value = (status)findings[i].late_status;

                        Rng.Style.Font.Bold = true;
                        Rng.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        Rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        Rng.Style.Fill.PatternType = ExcelFillStyle.Solid;

                        switch (findings[i].late_status)
                        {
                            case 0:
                                Rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);
                                break;

                            case 1:
                                Rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                                break;
                            case 2:
                                Rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                break;
                            default:
                                break;

                        }
                    }
                    using (ExcelRange Rng = sheet.Cells[$"H{i + 4}"])
                    {
                        Rng.Value = findings[i].department;
                    }
                    using (ExcelRange Rng = sheet.Cells[$"I{i + 4}"])
                    {
                        Rng.Value = findings[i].owner;
                        
                    }


                    row_Number++;

                    

                }

                row_Number++;
                row_Number++;


                using (ExcelRange Rng = sheet.Cells[$"A{row_Number}:D{row_Number + departments.Count}"])
                {
                    ExcelTable deptTable = sheet.Tables.Add(Rng, "deptTable");
                    deptTable.Columns[0].Name = "Department";
                    deptTable.Columns[1].Name = "Open";
                    deptTable.Columns[2].Name = "Due";
                    deptTable.Columns[3].Name = "Past Due";

                }

                row_Number++;
                int start = row_Number;

                foreach (var (department, i) in departments.Select((department, i) => (department, i)))
                {



                    using (ExcelRange Rng = sheet.Cells[$"A{i + start}"])
                    {
                        Rng.Value = department.name;
                    }
                    using (ExcelRange Rng = sheet.Cells[$"B{i + start}"])
                    {
                        Rng.Value = department.open;
                    }

                    using (ExcelRange Rng = sheet.Cells[$"C{i + start}"])
                    {
                        Rng.Value = department.due;
                    }
                    using (ExcelRange Rng = sheet.Cells[$"D{i + start}"])
                    {
                        Rng.Value = department.pastDue;
                    }

                    row_Number++;

                }







                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns(); //wtf does this do?




                //ExcelChart chart = sheet.Drawings.AddChart("Chart", eChartType.ColumnClustered);

                // }




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


                       //FileStream fs = File.Create(save_location);
                       // fs.Close();

                       // File.WriteAllBytes(save_location, package.GetAsByteArray());
                       // package.Dispose();

                    package.SaveAs(save_location);

                    return 0; //return successfully
                }
                catch {
                    return 2; // return error
                }


            }

            
        }

    }
}
