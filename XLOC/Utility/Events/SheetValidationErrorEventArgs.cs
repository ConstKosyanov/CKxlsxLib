using System;
using System.Collections.Generic;

namespace XLOC.Utility.Events
{
    public class SheetValidationErrorEventArgs : EventArgs
    {
        public SheetValidationErrorEventArgs(IEnumerable<string> missingFields, IEnumerable<Exception> exceptions, SheetIdentifier sheet)
        {
            MissingFields = missingFields;
            Exceptions = exceptions;
            Sheet = sheet;
        }

        public IEnumerable<string> MissingFields { get; set; }
        public IEnumerable<Exception> Exceptions { get; set; }
        public SheetIdentifier Sheet { get; set; }
    }
}