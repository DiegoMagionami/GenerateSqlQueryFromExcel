# ExcelToSQLConverter

GenerateSqlQueryFromExcel is a .NET 6 console application that reads an Excel or CSV file, prompts the user for a SQL query, and generates SQL statements for each row in the file using the provided query.

## Prerequisites

- .NET 6 SDK
- Microsoft.Office.Interop.Excel NuGet package

## Installation

1. Clone the repository:

   git clone https://github.com/DiegoMagionami/GenerateSqlQueryFromExcel.git

2. Navigate to the project directory:

	cd GenerateSqlQueryFromExcel

3. Build the application

	dotnet build
	
## Usage

1. Run the application:

	dotnet run
	
2. The application will prompt you for the following information:

Directory containing the Excel file.
Directory to save the SQL file.
SQL query using Excel columns as placeholders.

3. Provide the required information, and the application will generate a SQL file with the results in the specified directory.

## Example
Enter the directory containing the Excel file: C:\Path\To\ExcelFile
Enter the directory to save the SQL file: C:\Path\To\Output
Enter the SQL query using Excel columns as placeholders: INSERT INTO TableName (Column1, Column2) VALUES ('{{ColumnName1}}', '{{ColumnName2}}')
SQL file generated successfully at: C:\Path\To\Output\output.sql

## Notes
Ensure that Microsoft Excel is installed on the machine where the application runs since it uses the Microsoft.Office.Interop.Excel library.
License
This project is licensed under the MIT License - see the LICENSE file for details.