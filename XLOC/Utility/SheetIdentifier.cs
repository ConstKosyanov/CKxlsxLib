using DocumentFormat.OpenXml.Spreadsheet;

namespace XLOC.Utility
{
    public class SheetIdentifier
    {
        public SheetIdentifier(Sheet sheet)
        {
            Id = sheet.Id;
            Name = sheet.Name;
        }

        public string Id { get; set; }
        public string Name { get; set; }
    }
}