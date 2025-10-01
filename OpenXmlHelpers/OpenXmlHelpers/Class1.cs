using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace OpenXmlHelpers
{
    public static class NonGenerics
    {
        // Non-generic wrapper: ensures a SharedStringTablePart exists
        public static SharedStringTablePart GetOrAddSharedStringTablePart(WorkbookPart workbookPart)
        {
            var sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sstPart == null)
            {
                sstPart = workbookPart.AddNewPart<SharedStringTablePart>();
                sstPart.SharedStringTable = new SharedStringTable();
                sstPart.SharedStringTable.Save();
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
            sst.Save();
            return i;
        }
    }
}
