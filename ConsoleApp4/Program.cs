using System;
using System.IO;
using Newtonsoft.Json.Linq;
using Novacode;
using static System.Runtime.InteropServices.JavaScript.JSType;

class Program
{
    static void Main(string[] args)
    {
        string jsonFilePath = "path/to/yourfile.json";
        string odtFilePath = "path/to/output.odt";

        // Load JSON data
        var jsonData = File.ReadAllText(jsonFilePath);
        var jsonObject = JObject.Parse(jsonData);

        // Create a new ODT document
        using (DocX document = DocX.Create(odtFilePath))
        {
            // Extract data from JSON
            string id = jsonObject["id"].ToString();
            string type = jsonObject["type"].ToString();
            string date = jsonObject["date"].ToString();
            string dateUnixTime = jsonObject["date_unixtime"].ToString();
            string from = jsonObject["from"].ToString();
            string fromId = jsonObject["from_id"].ToString();
            var textArray = jsonObject["text"] as JArray;

            // Add content to the document
            document.InsertParagraph($"ID: {id}");
            document.InsertParagraph($"Type: {type}");
            document.InsertParagraph($"Date: {date}");
            document.InsertParagraph($"Date (Unix Time): {dateUnixTime}");
            document.InsertParagraph($"From: {from}");
            document.InsertParagraph($"From ID: {fromId}");
            document.InsertParagraph("Text:");

            foreach (var item in textArray)
            {
                if (item.Type == JTokenType.String)
                {
                    document.InsertParagraph(item.ToString());
                }
                else if (item.Type == JTokenType.Object && item["type"] != null && item["type"].ToString() == "bold")
                {
                    var boldText = item["text"].ToString();
                    var paragraph = document.InsertParagraph();
                    paragraph.Append(boldText).Bold();
                }
            }

            // Save the document
            document.Save();
        }

        Console.WriteLine("ODT file created successfully.");
    }
}