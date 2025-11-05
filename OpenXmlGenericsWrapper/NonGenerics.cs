using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;


namespace OpenXmlGenericsWrapper
{
    // Non-generic wrapper: ensures a SharedStringTablePart exists
    public static class NonGenerics
    {

        public static void InsertBeforeCell(Row row, Cell newChild, Cell referenceChild)
        {
            row.InsertBefore<Cell>(newChild, referenceChild);
            return;
        }

        public static void InsertBeforeRow(SheetData sheetData, Row newRow, Row referenceRow)
        {
            sheetData.InsertBefore<Row>(newRow, referenceRow);
            return;
        }

        public enum DataType
        {
            Boolean,
            Number,
            Error,
            SharedString,
            String,
            InlineString,
            Date
        }

        public static void SetCellDatatype(Cell cell, DataType datatype)
        {
            switch (datatype)
            {
                case DataType.Boolean:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
                    break;
                case DataType.Number:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    break;
                case DataType.Error:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Error);
                    break;
                case DataType.SharedString:
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    break;
                case DataType.String:
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    break;
                case DataType.InlineString:
                    cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
                    break;
                case DataType.Date:
                    cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                    break;
                default:
                    cell.DataType = null;
                    break;
            }

            return;
        }

        public static void SetCellReference(Cell cell, string value)
        {
            cell.CellReference = new StringValue(value);
            return;
        }

        public enum PartType
        {
            SharedString,
            Worksheet,
            WorkbookStyles
        }
        // Add a WorksheetPart to the WorkbookPart
        public static OpenXmlPart AddNewPart(WorkbookPart workbookPart, PartType partType)
        {
            switch (partType)
            {
                case PartType.Worksheet:
                    return workbookPart.AddNewPart<WorksheetPart>();
                case PartType.SharedString:
                    return workbookPart.AddNewPart<SharedStringTablePart>();
                case PartType.WorkbookStyles:
                    return workbookPart.AddNewPart<WorkbookStylesPart>();
                default:
                    throw new ArgumentException("Unsupported part type");
            }
        }

    }
}
