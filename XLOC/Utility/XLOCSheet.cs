using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLOC.Utility
{
    class XLOCSheet<T> : IGrouping<SheetIdentifier, T> where T : new()
    {
        #region Var
        //=================================================
        IEnumerable<T> enumerable;
        SheetIdentifier sheetIdentifier;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        public XLOCSheet(SheetIdentifier sheetIdentifier, IEnumerable<T> enumerable)
        {
            this.sheetIdentifier = sheetIdentifier;
            this.enumerable = enumerable;
        }
        //=================================================
        #endregion

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

        public SheetIdentifier Key => sheetIdentifier;

        public IEnumerator<T> GetEnumerator() => enumerable.GetEnumerator();

        internal static IGrouping<SheetIdentifier, T> Create(SheetIdentifier sheetIdentifier, IEnumerable<T> enumerable) => new XLOCSheet<T>(sheetIdentifier, enumerable);
    }
}