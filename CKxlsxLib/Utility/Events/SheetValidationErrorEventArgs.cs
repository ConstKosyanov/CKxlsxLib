using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLOC.Utility.Events
{
    public class SheetValidationErrorEventArgs : EventArgs
    {
        public SheetValidationErrorEventArgs(IEnumerable<string> missingFields, IEnumerable<Exception> exceptions)
        {
            MissingFields = missingFields;
            Exceptions = exceptions;
        }
        public IEnumerable<string> MissingFields { get; set; }
        public IEnumerable<Exception> Exceptions { get; set; }
    }
}