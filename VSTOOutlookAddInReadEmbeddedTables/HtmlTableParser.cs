using System;
using System.Configuration;
using System.Linq;
using HtmlAgilityPack;

public class HtmlTableParser
{
    public static string[] ExtractValuesFromHtml(string htmlContent)
    {
        // Load the HTML content into an HtmlDocument
        var doc = new HtmlDocument();
        doc.LoadHtml(htmlContent);

        // Read settings from App.config
        string targetRowLabel = ConfigurationManager.AppSettings["TargetRowLabel"];
        string[] headers = ConfigurationManager.AppSettings["Headers"].Split(',');

        // Find the table - assuming there is only one table in the HTML
        var table = doc.DocumentNode.SelectSingleNode("//table");

        // Get all rows of the table
        var rows = table.SelectNodes(".//tr");

        // Find the row that contains the target label
        var targetDataRow = rows.FirstOrDefault(row => row.SelectNodes(".//td").Any(cell => cell.InnerText.Contains(targetRowLabel)));
        if (targetDataRow == null)
        {
            throw new Exception("The specified row label was not found in any row.");
        }

        // Array to store the values to be returned
        string[] values = new string[headers.Length];

        // Find the header row and determine the indices of the required columns
        var headerRow = rows[0];
        var headerCells = headerRow.SelectNodes(".//th").Select(node => node.InnerText.Trim()).ToList();

        // Loop through headers and extract the required values
        for (int i = 0; i < headers.Length; i++)
        {
            int columnIndex = headerCells.IndexOf(headers[i]);
            if (columnIndex == -1)
            {
                throw new ArgumentException($"Header '{headers[i]}' not found.");
            }
            var cell = targetDataRow.SelectNodes(".//td")[columnIndex];
            values[i] = cell.InnerText.Trim();
        }

        return values;
    }
}
