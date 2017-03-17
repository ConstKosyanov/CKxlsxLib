using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLOC.Utility.Exceptions
{
    [Serializable]
    class CellReadingException : Exception
    {
        public CellReadingException(Exception innerException) : base(string.Empty, innerException) { }

        public XlContentType OutputType { get; internal set; }
        public string Reference { get; internal set; }
        public CellValues? SourceType { get; internal set; }
        public string Value { get; internal set; }
    }
}