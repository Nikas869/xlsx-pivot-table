using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;

namespace XlsxPivotTable
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 1)
            {
                return;
            }

            var absoluteFilePath = Path.GetFullPath(args[0]);

            using (var streamDocument = SpreadsheetDocument.Create($"{Path.GetFileNameWithoutExtension(absoluteFilePath)}_pivot.xlsx", SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = streamDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                var sheets = workbookPart.Workbook.AppendChild(new Sheets());

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Test Sheet" };
                sheets.Append(sheet);

                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }
        }
    }
}
