using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CKxlsxLib
{
    public class CKxlsxLibCellReadingErrorEventArgs : EventArgs
    {
        public Exception Exception { get; internal set; }
        public string Reference { get; internal set; }
        public xlContentType? SourceType { get; internal set; }
        public xlContentType? OutputType { get; internal set; }
        public object Value { get; internal set; }
    }
}