using XLOC.Utility.Events;
using System;

namespace XLOC
{
    public class XLOCConfiguration
    {
        public EventHandler<SheetValidationErrorEventArgs> ValidationFailureEvent { get; set; }
        public EventHandler<CellReadingErrorEventArgs> CellReadingErrorEvent { get; set; }
        public uint[] Sheets { get; set; }
        public bool ContinueOnRowReadingError { get; set; } = true;
    }
}
