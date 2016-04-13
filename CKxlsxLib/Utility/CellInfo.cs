using DocumentFormat.OpenXml.Spreadsheet;

namespace XLOC.Utility
{
    public class CellInfo
    {
        public CellInfo(string reference, xlContentType? type, object value, int? sharedId)
        {
            Reference = reference;
            ContentType = type;
            Value = value;
            SharedId = sharedId;
        }

        public string Reference { get; internal set; }
        public xlContentType? ContentType { get; internal set; }
        public object Value { get; internal set; }
        public int? SharedId { get; internal set; }
    }
}