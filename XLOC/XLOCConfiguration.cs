using XLOC.Utility.Events;
using System;
using DocumentFormat.OpenXml.Packaging;

namespace XLOC
{
    public class XLOCConfiguration
    {
        public EventHandler<SheetValidationErrorEventArgs> ValidationFailureEvent { get; set; }
        public EventHandler<CellReadingErrorEventArgs> CellReadingErrorEvent { get; set; }
        public uint[] Sheets { get; set; }
        public bool ContinueOnRowReadingError { get; set; } = true;
        public Utility.SkipModeEnum SkipMode { get; set; } = Utility.SkipModeEnum.None;
        public int? SkipCount { get; set; }
        public bool AutoDispose { get; set; } = true;
        internal SpreadsheetDocument Document { get; set; }

        internal XLOCConfiguration AddDocument(SpreadsheetDocument document)
        {
            Document = document;
            return this;
        }
    }
}