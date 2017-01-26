using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace XLOC.Utility.Events
{
    public class CellReadingErrorEventArgs : EventArgs
    {
        public Exception Exception { get; internal set; }
        public string Reference { get; internal set; }
        public uint Row { get; set; }
        public CellValues? SourceType { get; internal set; }
        public xlContentType? OutputType { get; internal set; }
        public object Value { get; internal set; }
        public SheetIdentifier Sheet { get; set; }
    }
}