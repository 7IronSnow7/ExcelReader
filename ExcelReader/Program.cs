using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ExcelReader
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Welcome to the Excel File Reader App.");
            Console.WriteLine("Please provide the filename of your excel file (e.g your-file.xlsx).");
            Console.WriteLine("Disclaimer - If no filename is given, 'your-file.xlsx' is used. Additionally, if file does not exist, a new one is created.");

            string filePath = Console.ReadLine();

            if (filePath == string.Empty) filePath = "your-file.xlsx";

            if (!File.Exists(filePath))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = doc.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    Sheets sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet() { Id = doc.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                    sheets.Append(sheet);
                    workbookPart.Workbook.Save();
                }
            }
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().First();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                if (!sheetData.Elements<Row>().Any())
                {
                    Console.WriteLine("Current Excel file is empty");
                }
                else
                {
                    Console.WriteLine("This is the current content of your file:");
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            Console.Write(cell.CellValue.Text + "\t");
                        }
                        Console.WriteLine();
                    }
                }
            }
            Console.WriteLine("\nWould you like to add a new row? (y/n)");
            string response = Console.ReadLine();

            if (response.ToLower() == "y")
            {
                Console.WriteLine("Enter data seperated by commas (e.g., 1,John,Doe):)");
                string[] rowData = Console.ReadLine().Split(',');

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, true))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().First();
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    Row newRow = new Row();
                    foreach (string cellData in rowData)
                    {
                        Cell cell = new Cell()
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(cellData)
                        };
                        newRow.Append(cell);
                    }
                    sheetData.Append(newRow);
                }
            }
            Console.WriteLine("\nThank your for using the Excel File Reader App! Run the program again to continue using.");
        }
    }
}