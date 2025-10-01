using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace OpenXmlHelpers
{
    // Non-generic wrapper: ensures a SharedStringTablePart exists
    public static class NonGenerics
    {
        public static SharedStringTablePart GetOrAddSharedStringTablePart(WorkbookPart workbookPart)
        {
            var sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sstPart == null)
            {
                sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
                sstPart.SharedStringTable = new SharedStringTable();
            }
            return sstPart;
        }

        // Example: add a string and return its index
        public static int InsertSharedString(WorkbookPart workbookPart, string text)
        {
            var sstPart = GetOrAddSharedStringTablePart(workbookPart);
            var sst = sstPart.SharedStringTable;

            // Check if string already exists
            int i = 0;
            foreach (var item in sst.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                    return i;
                i++;
            }

            // Otherwise, add it
            sst.AppendChild(new SharedStringItem(new Text(text)));
            return i;
        }

        public static WorkbookPart GetOrAddWorkbookPart(SpreadsheetDocument spreadSheet) {
            return spreadSheet.WorkbookPart ?? spreadSheet.AddWorkbookPart();
        }

        // Given a WorkbookPart, inserts a new worksheet.
        public static WorksheetPart GetOrAddWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();

            // Try to find the sheet by name
            Sheet existingSheet = sheets.Elements<Sheet>()
                                        .FirstOrDefault(s => s.Name == sheetName);

            if (existingSheet != null)
            {
                return (WorksheetPart)(workbookPart.GetPartById(existingSheet.Id));
            }

            // If not found, create a new worksheet
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Any())
            {
                sheetId = sheets.Elements<Sheet>().Max(s => s.SheetId.Value) + 1;
            }

            Sheet newSheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(newWorksheetPart),
                SheetId = sheetId,
                Name = sheetName
            };

            sheets.Append(newSheet);

            return newWorksheetPart;
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        public static Cell GetOrAddCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row = sheetData.Elements<Row>()
                               .FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == rowIndex);

            if (row == null)
            {
                row = new Row { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is a cell with the specified reference, return it.
            Cell existingCell = row.Elements<Cell>()
                                   .FirstOrDefault(c => c.CellReference != null &&
                                                        c.CellReference.Value == cellReference);
            if (existingCell != null)
                return existingCell;

            // Otherwise, create a new cell in the correct order.
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference != null ? cell.CellReference.Value : null,
                                   cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            Cell newCell = new Cell { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            return newCell;
        }
        //Sets Cell Value by SharedStringTable index
        public static void SetCellValueString(Cell cell, int index)
        {
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            return;
        }

    }
}
