using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qXlsxLib.Utility.Events
{
    public class CellReadingErrorEventArgs : EventArgs
    {
        public Exception Exception { get; internal set; }
        public string Reference { get; internal set; }
        public CellValues? SourceType { get; internal set; }
        public xlContentType? OutputType { get; internal set; }
        public object Value { get; internal set; }
    }
}