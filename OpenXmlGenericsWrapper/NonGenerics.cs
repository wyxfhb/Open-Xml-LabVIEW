using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;


namespace OpenXmlGenericsWrapper
{
    // Non-generic wrapper: ensures a SharedStringTablePart exists
    public static class NonGenerics
    {

        public static Cell InsertBeforeCell(Row row, Cell newChild, Cell referenceChild)
        {
            row.InsertBefore<Cell>(newChild, referenceChild);
            return newChild;
        }
        public enum DataType
        {
            SharedString
        }

        public static Cell SetCellDatatype(Cell cell, DataType datatype)
        {
            switch (datatype)
            {
                case DataType.SharedString:
                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    break;
                    // You can add more cases here for other DataType values if needed
            }

            return cell;
        }

        public static Cell SetCellReference(Cell cell, string value)
        {
            cell.CellReference = new StringValue(value);
            return cell;
        }

        public enum PartType
        {
            Worksheet,
            SharedString
        }
        // Add a WorksheetPart to the WorkbookPart
        public static OpenXmlPart AddNewPart(WorkbookPart workbookPart,PartType partType) {
            switch (partType)
            {
                case PartType.Worksheet:
                    return workbookPart.AddNewPart<WorksheetPart>();
                case PartType.SharedString:
                    return workbookPart.AddNewPart<SharedStringTablePart>();
                default:
                    throw new ArgumentException("Unsupported part type");
            }
        }

    }
}
