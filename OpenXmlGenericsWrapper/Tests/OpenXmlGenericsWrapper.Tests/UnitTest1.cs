using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlGenericsWrapper;
using DocumentFormat.OpenXml;

namespace OpenXmlGenericsWrapper.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestSetCellDatatype1()
        {
            Fonts fonts1 = new Fonts() { Count = (UInt32Value)1U };
            Font font1 = new Font();
            Font font2 = new Font();

            fonts1.Append(font1);
            int cnt = fonts1.ChildElements.Count;
            fonts1.Append(font2);
            cnt = fonts1.ChildElements.Count;

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
