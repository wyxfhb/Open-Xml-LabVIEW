using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlGenericsWrapper;

namespace OpenXmlGenericsWrapper.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestSetCellDatatype1()
        {
            UnderlineValues underlineValues = new UnderlineValues("single");
            Underline underline = new Underline();
            underline.Val = underlineValues;
            Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
            cell.CellValue = new CellValue("2006-01-01T06:30:00.123Z");
            Boolean equal = (cell.DataType.Equals(CellValues.Boolean)) ;
            Assert.AreEqual(cell.DataType, CellValues.Boolean);
        }
    }
}
