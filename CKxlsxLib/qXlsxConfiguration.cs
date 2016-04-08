using qXlsxLib.Utility.Events;
using System;

namespace qXlsxLib
{
    public class qXlsxConfiguration
    {
        public EventHandler<SheetValidationErrorEventArgs> ValidationFailureEvent { get; set; }
        public EventHandler<CellReadingErrorEventArgs> CellReadingErrorEvent { get; set; }
        public uint[] Sheets { get; set; }
        public bool ContinueOnRowReadingError { get; set; } = true;
    }
}
