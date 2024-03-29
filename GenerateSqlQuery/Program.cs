using CsvHelper;
using Microsoft.Office.Interop.Excel;
using System.Formats.Asn1;
using System.Globalization;

class Program
{
    static void Main()
    {
        Console.WriteLine("Enter the directory containing the Excel file:");
        string excelDirectory = Console.ReadLine();

        Console.WriteLine("Enter the directory to save the SQL file:");
        string sqlDirectory = Console.ReadLine();

        Console.WriteLine("Enter the SQL query using Excel columns as placeholders:");
        string sqlQuery = Console.ReadLine();

        Console.WriteLine("Choose the file type (1 for Excel, 2 for CSV):");
        int fileTypeChoice = int.Parse(Console.ReadLine());

        if (fileTypeChoice == 1)
        {
            try
            {
                List<Dictionary<string, object>> excelData = ReadExcel(excelDirectory);

                string sqlResult = GenerateSqlQueries(excelData, sqlQuery);

                string sqlFilePath = Path.Combine(sqlDirectory, "output.sql");
                File.WriteAllText(sqlFilePath, sqlResult);

                Console.WriteLine($"SQL file generated successfully at: {sqlFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
        else if (fileTypeChoice == 2)
        {
            try
            {
                List<Dictionary<string, object>> excelData = ReadExcel(excelDirectory);

                string sqlResult = GenerateSqlQueries(excelData, sqlQuery);

                string sqlFilePath = Path.Combine(sqlDirectory, "output.sql");
                File.WriteAllText(sqlFilePath, sqlResult);

                Console.WriteLine($"SQL file generated successfully at: {sqlFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("Invalid choice. Exiting the program.");
        }

        
    }

    static List<Dictionary<string, object>> ReadExcel(string excelDirectory)
    {
        List<Dictionary<string, object>> excelData = new List<Dictionary<string, object>>();

        string excelFilePath = Path.Combine(excelDirectory, "input.xlsx");

        var excelApp = new Microsoft.Office.Interop.Excel.Application();
        var workbooks = excelApp.Workbooks;
        var workbook = workbooks.Open(excelFilePath);
        var worksheet = (Worksheet)workbook.Sheets[1];

        var lastRow = worksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

        // Assuming the first row contains column headers
        var columns = new List<string>();
        for (int col = 1; col <= worksheet.UsedRange.Columns.Count; col++)
        {
            columns.Add(((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, col]).Text.ToString());
        }

        for (int row = 2; row <= lastRow; row++)
        {
            var rowData = new Dictionary<string, object>();

            for (int col = 1; col <= columns.Count; col++)
            {
                rowData[columns[col - 1]] = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, col]).Text;
            }

            excelData.Add(rowData);
        }

        workbook.Close();
        excelApp.Quit();

        return excelData;
    }

    static List<Dictionary<string, object>> ReadCsv(string csvDirectory)
    {
        List<Dictionary<string, object>> csvData = new List<Dictionary<string, object>>();

        Console.WriteLine("Enter the CSV file name:");
        string csvFileName = Console.ReadLine();

        string csvFilePath = Path.Combine(csvDirectory, csvFileName);

        using (var reader = new StreamReader(csvFilePath))
        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            var records = csv.GetRecords<dynamic>();

            foreach (var record in records)
            {
                var rowData = new Dictionary<string, object>();

                foreach (var property in record.GetType().GetProperties())
                {
                    rowData[property.Name] = property.GetValue(record);
                }

                csvData.Add(rowData);
            }
        }

        return csvData;
    }

    static string GenerateSqlQueries(List<Dictionary<string, object>> excelData, string sqlQuery)
    {
        StringWriter sw = new StringWriter();

        foreach (var rowData in excelData)
        {
            string formattedSql = sqlQuery;

            foreach (var kvp in rowData)
            {
                formattedSql = formattedSql.Replace($"{{{{{kvp.Key}}}}}", kvp.Value.ToString());
            }

            sw.WriteLine(formattedSql);
        }

        return sw.ToString();
    }
}
