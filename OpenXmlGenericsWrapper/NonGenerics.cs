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
        //Boolean => new CellValues("b");
        //Number => new CellValues("n");
        //Error => new CellValues("e");
        //SharedString => new CellValues("s");
        //String => new CellValues("str");
        //InlineString => new CellValues("inlineStr");
        //Date => new CellValues("d");
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

        public static Cell SetCellDatatype(Cell cell, DataType datatype)
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
        public static OpenXmlPart AddNewPart(WorkbookPart workbookPart, PartType partType)
        {
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
