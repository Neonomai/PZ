using System;
using System.IO;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        string jsonFilePath = "result.json"; 
        string excelFilePath = "output.xlsx"; 

        
        var jsonData = File.ReadAllText(jsonFilePath);
        var jsonObject = JObject.Parse(jsonData);

        
        using (ExcelPackage package = new ExcelPackage())
        {
            
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Отчет");

           
            worksheet.Cells[1, 1].Value = "Hostname";
            worksheet.Cells[1, 2].Value = "Date";
            worksheet.Cells[1, 3].Value = "Problem";
            worksheet.Cells[1, 4].Value = "Level";

            
            using (var range = worksheet.Cells[1, 1, 1, 4])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            
            var messages = jsonObject["messages"] as JArray;
            if (messages != null)
            {
                int rowIndex = 2;
                foreach (var message in messages)
                {
                    
                    string hostname = "Graylogbot"; 
                    string date = message["date"]?.ToString();
                    string problem = string.Empty;
                    string level = string.Empty;

                    var textArray = message["text"] as JArray;
                    if (textArray != null)
                    {
                        foreach (var item in textArray)
                        {
                            if (item.Type == JTokenType.String)
                            {
                                problem += item.ToString();
                            }
                            else if (item.Type == JTokenType.Object && item["type"] != null)
                            {
                                if (item["type"].ToString() == "bold")
                                {
                                    level += item["text"].ToString();
                                }
                            }
                        }
                    }

                    
                    worksheet.Cells[rowIndex, 1].Value = hostname;
                    worksheet.Cells[rowIndex, 2].Value = date;
                    worksheet.Cells[rowIndex, 3].Value = problem;
                    worksheet.Cells[rowIndex, 4].Value = level;
                    rowIndex++;
                }
            }

            
            FileInfo excelFile = new FileInfo(excelFilePath);
            package.SaveAs(excelFile);
        }

        Console.WriteLine("Excel-файл успешно создан.");
    }
}
