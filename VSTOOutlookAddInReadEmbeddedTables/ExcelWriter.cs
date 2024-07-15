using OfficeOpenXml;
using System.Configuration;
using System.IO;

public class ExcelWriter
{
    static ExcelWriter()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public static void WriteToExcel(string[] values)
    {
        string filePath = ConfigurationManager.AppSettings["ExcelFilePath"];
        FileInfo fileInfo = new FileInfo(filePath);

        using (var package = new ExcelPackage(fileInfo))
        {
            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets.Count > 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add("Sheet1");

            int row = worksheet.Dimension?.End.Row + 1 ?? 1;

            for (int col = 1; col <= values.Length; col++)
            {
                worksheet.Cells[row, col].Value = values[col - 1];
            }

            package.Save();
        }
    }
}
