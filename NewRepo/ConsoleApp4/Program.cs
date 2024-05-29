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

        string jsonFilePath = "result.json"; // Укажите путь к вашему JSON-файлу
        string excelFilePath = "output.xlsx"; // Укажите путь для сохранения Excel-файла

        // Загрузка JSON-данных
        var jsonData = File.ReadAllText(jsonFilePath);
        var jsonObject = JObject.Parse(jsonData);

        // Создание нового Excel-документа
        using (ExcelPackage package = new ExcelPackage())
        {
            // Создание нового листа
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Отчет");

            // Установка заголовков таблицы
            worksheet.Cells[1, 1].Value = "Hostname";
            worksheet.Cells[1, 2].Value = "Date";
            worksheet.Cells[1, 3].Value = "Problem";
            worksheet.Cells[1, 4].Value = "Level";

            // Установка стилей заголовков
            using (var range = worksheet.Cells[1, 1, 1, 4])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            // Заполнение таблицы данными из JSON
            var messages = jsonObject["messages"] as JArray;
            if (messages != null)
            {
                int rowIndex = 2;
                foreach (var message in messages)
                {
                    // Извлечение данных
                    string hostname = "Graylogbot"; // Укажите hostname, если он доступен в JSON
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

                    // Заполнение строки таблицы
                    worksheet.Cells[rowIndex, 1].Value = hostname;
                    worksheet.Cells[rowIndex, 2].Value = date;
                    worksheet.Cells[rowIndex, 3].Value = problem;
                    worksheet.Cells[rowIndex, 4].Value = level;
                    rowIndex++;
                }
            }

            // Сохранение Excel-файла
            FileInfo excelFile = new FileInfo(excelFilePath);
            package.SaveAs(excelFile);
        }

        Console.WriteLine("Excel-файл успешно создан.");
    }
}
